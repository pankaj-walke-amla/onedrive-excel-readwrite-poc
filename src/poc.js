const axios = require("axios");
const ExcelJS = require("exceljs");
const fs = require("fs");
const path = require("path");

const GRAPH_BASE_URL = "https://graph.microsoft.com/v1.0";

function loadEnvFile(fileName) {
  const filePath = path.join(process.cwd(), fileName);
  if (!fs.existsSync(filePath)) {
    return;
  }

  const content = fs.readFileSync(filePath, "utf8");
  const lines = content.split(/\r?\n/);

  for (const line of lines) {
    const trimmed = line.trim();
    if (!trimmed || trimmed.startsWith("#")) {
      continue;
    }

    const equalsIndex = trimmed.indexOf("=");
    if (equalsIndex === -1) {
      continue;
    }

    const key = trimmed.slice(0, equalsIndex).trim();
    if (!key || process.env[key]) {
      continue;
    }

    let value = trimmed.slice(equalsIndex + 1).trim();
    if (
      (value.startsWith('"') && value.endsWith('"')) ||
      (value.startsWith("'") && value.endsWith("'"))
    ) {
      value = value.slice(1, -1);
    }

    process.env[key] = value;
  }
}

function readEnvValueFromFile(fileName, keyName) {
  const filePath = path.join(process.cwd(), fileName);
  if (!fs.existsSync(filePath)) {
    return undefined;
  }

  const content = fs.readFileSync(filePath, "utf8");
  const lines = content.split(/\r?\n/);

  for (const line of lines) {
    const trimmed = line.trim();
    if (!trimmed || trimmed.startsWith("#")) {
      continue;
    }

    const equalsIndex = trimmed.indexOf("=");
    if (equalsIndex === -1) {
      continue;
    }

    const key = trimmed.slice(0, equalsIndex).trim();
    if (key !== keyName) {
      continue;
    }

    let value = trimmed.slice(equalsIndex + 1).trim();
    if (
      (value.startsWith('"') && value.endsWith('"')) ||
      (value.startsWith("'") && value.endsWith("'"))
    ) {
      value = value.slice(1, -1);
    }

    return value;
  }

  return undefined;
}

// Allow running with .env, and support .env.example for this POC setup.
loadEnvFile(".env");
loadEnvFile(".env.example");

function getRequiredEnv(name) {
  const value = process.env[name];
  if (!value || !value.trim()) {
    throw new Error(`${name} is required.`);
  }
  return value.trim();
}

function validateAccessTokenFormat(token) {
  // Graph access tokens are JWTs and should have three dot-separated segments.
  if (token.split(".").length !== 3) {
    throw new Error(
      "GRAPH_ACCESS_TOKEN appears malformed. Use the full JWT token value (header.payload.signature) with no placeholder text or truncation."
    );
  }
}

function isJwtLikeToken(token) {
  return !!token && token.split(".").length === 3;
}

function resolveAccessToken() {
  const cliToken = process.argv[2] ? process.argv[2].trim() : "";
  if (cliToken) {
    return cliToken;
  }

  const envToken = process.env.GRAPH_ACCESS_TOKEN
    ? process.env.GRAPH_ACCESS_TOKEN.trim()
    : "";

  if (isJwtLikeToken(envToken)) {
    return envToken;
  }

  const dotEnvToken = (readEnvValueFromFile(".env", "GRAPH_ACCESS_TOKEN") || "").trim();
  if (isJwtLikeToken(dotEnvToken)) {
    if (envToken) {
      console.warn("WARN: Terminal GRAPH_ACCESS_TOKEN is malformed. Using token from .env.");
    }
    return dotEnvToken;
  }

  const exampleToken = (readEnvValueFromFile(".env.example", "GRAPH_ACCESS_TOKEN") || "").trim();
  if (isJwtLikeToken(exampleToken)) {
    if (envToken) {
      console.warn("WARN: Terminal GRAPH_ACCESS_TOKEN is malformed. Using token from .env.example.");
    }
    return exampleToken;
  }

  return getRequiredEnv("GRAPH_ACCESS_TOKEN");
}

function getErrorMessage(error) {
  if (error.response) {
    const status = error.response.status;
    const data = JSON.stringify(error.response.data);
    return `Graph request failed with ${status}: ${data}`;
  }
  return error.message;
}

function isResourceLockedError(error) {
  return error && error.response && error.response.status === 423;
}

function getRetryDelayMs(error, attempt) {
  const retryAfter = Number(error && error.response && error.response.headers
    ? error.response.headers["retry-after"]
    : undefined);

  if (Number.isFinite(retryAfter) && retryAfter > 0) {
    return retryAfter * 1000;
  }

  return attempt * 1000;
}

function sleep(ms) {
  return new Promise((resolve) => setTimeout(resolve, ms));
}

function buildFallbackFileName(fileName) {
  const extension = path.extname(fileName);
  const baseName = extension ? fileName.slice(0, -extension.length) : fileName;
  return `${baseName}-${Date.now()}${extension || ".xlsx"}`;
}

async function withLockRetry(action, description, maxAttempts = 3) {
  for (let attempt = 1; attempt <= maxAttempts; attempt += 1) {
    try {
      return await action();
    } catch (error) {
      if (!isResourceLockedError(error) || attempt === maxAttempts) {
        throw error;
      }

      const delayMs = getRetryDelayMs(error, attempt);
      console.warn(
        `WARN: ${description} is locked (attempt ${attempt}/${maxAttempts}). Retrying in ${delayMs}ms.`
      );
      await sleep(delayMs);
    }
  }
}

async function createWorkbookBinary() {
  const workbook = new ExcelJS.Workbook();
  const worksheet = workbook.addWorksheet("Data");
  worksheet.getCell("A1").value = "Initial value";
  worksheet.getCell("B1").value = "Created by POC";

  const buffer = await workbook.xlsx.writeBuffer();
  return Buffer.from(buffer);
}

async function run() {
  const accessToken = resolveAccessToken();
  validateAccessTokenFormat(accessToken);
  const fileName = process.env.POC_FILE_NAME || `onedrive-excel-poc-${Date.now()}.xlsx`;
  const writeValue = process.env.POC_VALUE || `Updated at ${new Date().toISOString()}`;

  const graph = axios.create({
    baseURL: GRAPH_BASE_URL,
    headers: {
      Authorization: `Bearer ${accessToken}`,
      "Content-Type": "application/json",
    },
    timeout: 30000,
  });

  const workbookBinary = await createWorkbookBinary();

  let actualFileName = fileName;
  let uploadResponse;

  try {
    uploadResponse = await withLockRetry(
      () => axios.put(
        `${GRAPH_BASE_URL}/me/drive/root:/${encodeURIComponent(actualFileName)}:/content`,
        workbookBinary,
        {
          headers: {
            Authorization: `Bearer ${accessToken}`,
            "Content-Type":
              "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
          },
          timeout: 30000,
        }
      ),
      `Upload for ${actualFileName}`
    );
  } catch (error) {
    if (!isResourceLockedError(error)) {
      throw error;
    }

    actualFileName = buildFallbackFileName(fileName);
    console.warn(
      `WARN: ${fileName} is locked. Falling back to ${actualFileName}.`
    );

    uploadResponse = await withLockRetry(
      () => axios.put(
        `${GRAPH_BASE_URL}/me/drive/root:/${encodeURIComponent(actualFileName)}:/content`,
        workbookBinary,
        {
          headers: {
            Authorization: `Bearer ${accessToken}`,
            "Content-Type":
              "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
          },
          timeout: 30000,
        }
      ),
      `Upload for ${actualFileName}`
    );
  }

  const itemId = uploadResponse.data.id;
  if (!itemId) {
    throw new Error("Upload succeeded but OneDrive item id was not returned.");
  }

  const createSessionResponse = await withLockRetry(
    () => graph.post(
      `/me/drive/items/${itemId}/workbook/createSession`,
      { persistChanges: true }
    ),
    "Workbook createSession"
  );

  const workbookSessionId = createSessionResponse.data.id;
  if (!workbookSessionId) {
    throw new Error("Workbook session id was not returned.");
  }

  const sessionHeaders = { "workbook-session-id": workbookSessionId };
  let readResponse;

  try {
    await withLockRetry(
      () => graph.patch(
        `/me/drive/items/${itemId}/workbook/worksheets/Data/range(address='A1')`,
        { values: [[writeValue]] },
        { headers: sessionHeaders }
      ),
      "Workbook write"
    );

    readResponse = await withLockRetry(
      () => graph.get(
        `/me/drive/items/${itemId}/workbook/worksheets/Data/range(address='A1')`,
        { headers: sessionHeaders }
      ),
      "Workbook read"
    );
  } finally {
    try {
      await graph.post(
        `/me/drive/items/${itemId}/workbook/closeSession`,
        null,
        { headers: sessionHeaders }
      );
    } catch (closeError) {
      console.warn(`WARN: Failed to close workbook session: ${getErrorMessage(closeError)}`);
    }
  }

  const readBackValue = readResponse.data && readResponse.data.values
    ? readResponse.data.values[0][0]
    : undefined;

  if (readBackValue !== writeValue) {
    throw new Error(
      `Verification failed. Expected "${writeValue}" but got "${readBackValue}".`
    );
  }

  console.log("POC_SUCCESS");
  console.log(`FileName: ${actualFileName}`);
  console.log(`ItemId: ${itemId}`);
  console.log(`WebUrl: ${uploadResponse.data.webUrl}`);
  console.log(`WrittenValue: ${writeValue}`);
  console.log(`ReadBackValue: ${readBackValue}`);
}

run().catch((error) => {
  console.error(`POC_FAILED: ${getErrorMessage(error)}`);
  process.exit(1);
});

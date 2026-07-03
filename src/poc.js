const axios = require("axios");
const ExcelJS = require("exceljs");

const GRAPH_BASE_URL = "https://graph.microsoft.com/v1.0";

function getRequiredEnv(name) {
  const value = process.env[name];
  if (!value || !value.trim()) {
    throw new Error(`${name} is required.`);
  }
  return value.trim();
}

function getErrorMessage(error) {
  if (error.response) {
    const status = error.response.status;
    const data = JSON.stringify(error.response.data);
    return `Graph request failed with ${status}: ${data}`;
  }
  return error.message;
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
  const accessToken = process.argv[2] || getRequiredEnv("GRAPH_ACCESS_TOKEN");
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

  const uploadResponse = await axios.put(
    `${GRAPH_BASE_URL}/me/drive/root:/${encodeURIComponent(fileName)}:/content`,
    workbookBinary,
    {
      headers: {
        Authorization: `Bearer ${accessToken}`,
        "Content-Type":
          "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
      },
      timeout: 30000,
    }
  );

  const itemId = uploadResponse.data.id;
  if (!itemId) {
    throw new Error("Upload succeeded but OneDrive item id was not returned.");
  }

  const createSessionResponse = await graph.post(
    `/me/drive/items/${itemId}/workbook/createSession`,
    { persistChanges: true }
  );

  const workbookSessionId = createSessionResponse.data.id;
  if (!workbookSessionId) {
    throw new Error("Workbook session id was not returned.");
  }

  const sessionHeaders = { "workbook-session-id": workbookSessionId };

  await graph.patch(
    `/me/drive/items/${itemId}/workbook/worksheets/Data/range(address='A1')`,
    { values: [[writeValue]] },
    { headers: sessionHeaders }
  );

  const readResponse = await graph.get(
    `/me/drive/items/${itemId}/workbook/worksheets/Data/range(address='A1')`,
    { headers: sessionHeaders }
  );

  await graph.post(
    `/me/drive/items/${itemId}/workbook/closeSession`,
    null,
    { headers: sessionHeaders }
  );

  const readBackValue = readResponse.data && readResponse.data.values
    ? readResponse.data.values[0][0]
    : undefined;

  if (readBackValue !== writeValue) {
    throw new Error(
      `Verification failed. Expected "${writeValue}" but got "${readBackValue}".`
    );
  }

  console.log("POC_SUCCESS");
  console.log(`FileName: ${fileName}`);
  console.log(`ItemId: ${itemId}`);
  console.log(`WebUrl: ${uploadResponse.data.webUrl}`);
  console.log(`WrittenValue: ${writeValue}`);
  console.log(`ReadBackValue: ${readBackValue}`);
}

run().catch((error) => {
  console.error(`POC_FAILED: ${getErrorMessage(error)}`);
  process.exit(1);
});

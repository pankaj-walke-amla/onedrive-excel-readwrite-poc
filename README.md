# OneDrive Excel Read/Write POC

Working POC that uploads an Excel file to OneDrive and performs verified read/write on cell `A1` using Microsoft Graph with an access token.

## Implementation

1. Create a Graph access token for delegated permissions that can read/write files (for example `Files.ReadWrite` or broader equivalent).
2. Export the token as `GRAPH_ACCESS_TOKEN`.
3. Run the script to:
   - create an Excel file in memory
   - upload it to your OneDrive root
   - write to worksheet `Data`, cell `A1`
   - read `A1` back and verify value equality
4. On success, the script prints file name, item id, OneDrive web URL, written value, and read-back value.

## Project setup

```powershell
npm install
```

## Run

```powershell
$env:GRAPH_ACCESS_TOKEN="YOUR_TOKEN"
$env:POC_FILE_NAME="onedrive-excel-poc.xlsx"   # optional
$env:POC_VALUE="Hello from POC"                # optional
npm start
```

## Expected successful output

```text
POC_SUCCESS
FileName: ...
ItemId: ...
WebUrl: ...
WrittenValue: ...
ReadBackValue: ...
```

If token/scope/session is invalid, script exits with `POC_FAILED: ...` and the Graph error payload.

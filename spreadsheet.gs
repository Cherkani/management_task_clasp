const SpreadsheetService = (() => {
  let cachedSpreadsheet = null;

  const getSpreadsheet = () => {
    if (!cachedSpreadsheet) {
      cachedSpreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    }
    return cachedSpreadsheet;
  };

  const normalizeHeaders = (headers = []) => {
    const filtered = (headers || []).filter(Boolean).map(String);
    if (!filtered.length) {
      return ["id"];
    }
    return filtered[0] === "id" ? filtered : ["id", ...filtered];
  };

  const getSheetOrCreate = (sheetName, headers = []) => {
    let sheet = getSpreadsheet().getSheetByName(sheetName);
    if (sheet) {
      return sheet;
    }
    sheet = getSpreadsheet().insertSheet(sheetName);
    sheet.appendRow(normalizeHeaders(headers));
    return sheet;
  };

  const getSheetValues = (sheetName) => {
    const sheet = getSpreadsheet().getSheetByName(sheetName);
    if (!sheet) {
      return [];
    }
    return sheet.getDataRange().getDisplayValues();
  };

  const rowsToObjects = (values) => {
    if (!values || values.length < 2) {
      return [];
    }
    const headers = values[0];
    return values.slice(1).map((row) =>
      headers.reduce((acc, key, idx) => {
        acc[key] = row[idx] ?? "";
        return acc;
      }, {})
    );
  };

  return {
    getSpreadsheet,
    getSheetOrCreate,
    getSheetValues,
    normalizeHeaders,
    rowsToObjects,
  };
})();

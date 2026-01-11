const toJsonString = (payload) => JSON.stringify(payload);

const handleError = (error, fallback = null) => {
  console.error(error);
  return toJsonString({
    error: error.message || error.toString(),
    data: fallback,
  });
};

const executeAction = ({
  action,
  sheetName,
  id = null,
  data = null,
  headers = [],
}) => {
  if (!sheetName) {
    return handleError(new Error("sheetName is required"));
  }

  try {
    const orm = new ORM(sheetName, headers);
    let res;

    switch (action) {
      case "add":
        res = orm.create(data);
        break;
      case "update":
        res = orm.updateById(id, data);
        break;
      case "delete":
        res = orm.deleteById(id);
        break;
      case "getAll":
        res = orm.readAll();
        break;
      case "getOne":
        res = orm.readById(id);
        break;
      default:
        throw new Error(`Unknown action "${action}"`);
    }

    return toJsonString(res);
  } catch (error) {
    return handleError(error);
  }
};

const getSheetData = (name) => {
  try {
    const values = SpreadsheetService.getSheetValues(name);
    const objects = SpreadsheetService.rowsToObjects(values);
    return toJsonString(objects);
  } catch (error) {
    return handleError(error, []);
  }
};

const getSheetNamesAndHeaders = () => {
  const sheets = SpreadsheetService.getSpreadsheet().getSheets();
  const sheetNames = sheets.map((sheet) => sheet.getName());

  const headers = sheets.map((sheet) => {
    const dataRange = sheet.getDataRange();
    const headers = dataRange.getValues()[0];
    return { [sheet.getName()]: headers };
  });
  return { sheetNames, headers };
};

const addupdateItemToSheet = (sheetName, data) => {
  try {
    return executeAction({
      action: "add",
      sheetName,
      id: data.id || null,
      data,
    });
  } catch (error) {
    return handleError(error);
  }
};

const getDropdowns = (sheetNames) => {
  const dropdowns = {};
  sheetNames.forEach((sheetName) => {
    dropdowns[sheetName] = SpreadsheetService.getSheetValues(sheetName);
  });
  return dropdowns;
};

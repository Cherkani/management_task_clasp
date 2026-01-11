const SPREADSHEET_ID = "1AxZWkXf1BH2SW6awq0vchyl9OZVanCtxVXKJ_HW_fD4";

const executeAction = ({
  action,
  sheetName,
  id = null,
  data = null,
  headers = null,
}) => {
  let res = null;
  try {
    const orm = new ORM(sheetName, headers);

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
        return false;
    }

    return JSON.stringify(res);
  } catch (error) {
    console.log(error);
    return error;
  }
};

const getSheetData = (name) => {
  const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(name);
  const dataRange = sheet.getDataRange();
  const data = dataRange.getDisplayValues();
  const heads = data.shift();
  const obj = data.map((r) =>
    heads.reduce((o, k, i) => ((o[k] = r[i] || ""), o), {})
  );
  return JSON.stringify(obj);
};

const getSheetNamesAndHeaders = () => {
  const sheets = SpreadsheetApp.openById(SPREADSHEET_ID).getSheets();
  const sheetNames = sheets.map((sheet) => sheet.getName());

  const headers = sheets.map((sheet) => {
    const dataRange = sheet.getDataRange();
    const headers = dataRange.getValues()[0];
    return { [sheet.getName()]: headers };
  });
  return { sheetNames, headers };
};

const addupdateItemToSheet = (sheetName, data) => {
  // check if sheet exists else create it
  let sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(sheetName);
  if (!sheet) {
    sheet = SpreadsheetApp.openById(SPREADSHEET_ID).insertSheet(sheetName);
    sheet.getRange("A1").setValue("id");
    sheet.getRange("B1").setValue("name");
  }

  try {
    return executeAction({
      action: "add",
      sheetName,
      id: data.id || null,
      data,
    });
  } catch (error) {
    return error;
  }
};

const getDropdowns = (sheetNames) => {
  const dropdowns = {};
  sheetNames.forEach((sheetName) => {
    const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(sheetName);
    if (sheet) {
      dropdowns[sheetName] = sheet.getDataRange().getValues();
    } else {
      dropdowns[sheetName] = [];
    }
  });
  return dropdowns;
};

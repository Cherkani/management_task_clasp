const SPREADSHEET_ID = "1AxZWkXf1BH2SW6awq0vchyl9OZVanCtxVXKJ_HW_fD4";
const CALENDAR_ID = null; // Set to a specific calendar ID if you don't want to use the default one.
const DEFAULT_EVENT_DURATION_MINUTES = 60;

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
    const result = executeAction({
      action: "add",
      sheetName,
      id: data.id || null,
      data,
    });
    syncTaskToCalendar(sheetName, data);
    return result;
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

const syncTaskToCalendar = (sheetName, data) => {
  if (!data) return;

  try {
    const eventTitle =
      data.title ||
      data.name ||
      data.taskName ||
      data.summary ||
      `${sheetName || "Task"}${data.id ? ` #${data.id}` : ""}`;
    const startDate = buildDate(
      data.startDate || data.start || data.date || data.dueDate,
      data.startTime || data.time
    );
    const endDate =
      buildEndDate(
        data.endDate || data.end || data.dueDate || data.date,
        data.endTime,
        startDate
      ) || (startDate ? new Date(startDate.getTime() + DEFAULT_EVENT_DURATION_MINUTES * 60 * 1000) : null);

    if (!eventTitle || !startDate || !endDate) {
      return;
    }

    const calendar = CALENDAR_ID
      ? CalendarApp.getCalendarById(CALENDAR_ID)
      : CalendarApp.getDefaultCalendar();

    const description =
      data.description || data.details || data.notes || data.comment || "";

    calendar.createEvent(eventTitle, startDate, endDate, {
      description,
      location: data.location || "",
    });
  } catch (error) {
    console.error("Failed to sync task to calendar", error);
  }
};

const buildDate = (dateValue, timeValue) => {
  if (!dateValue) return null;
  const date =
    dateValue instanceof Date ? new Date(dateValue.getTime()) : new Date(dateValue);
  if (isNaN(date.getTime())) return null;
  if (!timeValue) return date;

  const [hoursPart, minutesPart] = String(timeValue).split(":");
  const hours = Number(hoursPart);
  const minutes = Number(minutesPart || 0);
  if (!isNaN(hours)) {
    date.setHours(hours, isNaN(minutes) ? 0 : minutes, 0, 0);
  }
  return date;
};

const buildEndDate = (dateValue, timeValue, fallbackStart) => {
  const endDate = buildDate(dateValue, timeValue);
  if (endDate) {
    return endDate;
  }
  if (fallbackStart) {
    return new Date(
      fallbackStart.getTime() + DEFAULT_EVENT_DURATION_MINUTES * 60 * 1000
    );
  }
  return null;
};

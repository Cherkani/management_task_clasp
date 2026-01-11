const SPREADSHEET_ID = "1AxZWkXf1BH2SW6awq0vchyl9OZVanCtxVXKJ_HW_fD4";
const CALENDAR_ID = null; // Set to a specific calendar ID if you don't want to use the default one.
const CALENDAR_EVENT_COLUMN = "calendarEventId";
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
        handleCalendarSync({
          type: "add",
          sheetName,
          id: data?.id,
          data,
        });
        break;
      case "update":
        res = orm.updateById(id, data);
        handleCalendarSync({
          type: "update",
          sheetName,
          id,
          data,
        });
        break;
      case "delete":
        handleCalendarSync({
          type: "delete",
          sheetName,
          id,
        });
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
    sheet.getRange("C1").setValue(CALENDAR_EVENT_COLUMN);
  } else {
    ensureCalendarEventColumn(sheet);
  }

  try {
    const result = executeAction({
      action: "add",
      sheetName,
      id: data.id || null,
      data,
    });
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
  if (!data) return null;

  try {
    const existingEventId =
      data[CALENDAR_EVENT_COLUMN] ||
      data.eventId ||
      getCalendarEventIdFromSheet(sheetName, data.id);

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
      return null;
    }

    const calendar = getCalendar();

    const description =
      data.description || data.details || data.notes || data.comment || "";
    const location = data.location || "";

    if (existingEventId) {
      const existingEvent = calendar.getEventById(existingEventId);
      if (existingEvent) {
        existingEvent.setTitle(eventTitle);
        existingEvent.setTime(startDate, endDate);
        existingEvent.setDescription(description);
        existingEvent.setLocation(location);
        return existingEventId;
      }
    }

    const event = calendar.createEvent(eventTitle, startDate, endDate, {
      description,
      location,
    });
    return event?.getId() || null;
  } catch (error) {
    console.error("Failed to sync task to calendar", error);
    return null;
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

const handleCalendarSync = ({ type, sheetName, id, data }) => {
  if (!sheetName) return;
  const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(sheetName);
  if (!sheet) return;
  ensureCalendarEventColumn(sheet);

  if (type === "delete") {
    deleteCalendarEvent(sheetName, id);
    return;
  }

  if (!data) return;
  if (!data.id && id) {
    data.id = id;
  }
  const eventId = syncTaskToCalendar(sheetName, data);
  if (eventId && id) {
    updateCalendarEventIdInSheet(sheetName, id, eventId);
  }
};

const ensureCalendarEventColumn = (sheet) => {
  const lastColumn = sheet.getLastColumn() || 1;
  const headers = sheet.getRange(1, 1, 1, lastColumn).getValues()[0];
  if (!headers.includes(CALENDAR_EVENT_COLUMN)) {
    sheet.insertColumnAfter(lastColumn);
    sheet.getRange(1, lastColumn + 1).setValue(CALENDAR_EVENT_COLUMN);
  }
};

const getCalendarEventIdFromSheet = (sheetName, recordId) => {
  if (!recordId) return null;
  const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(sheetName);
  if (!sheet) return null;
  ensureCalendarEventColumn(sheet);
  const dataRange = sheet.getDataRange();
  const values = dataRange.getValues();
  if (!values.length) return null;
  const headers = values[0];
  const idIndex = headers.indexOf("id");
  const eventIndex = headers.indexOf(CALENDAR_EVENT_COLUMN);
  if (idIndex === -1 || eventIndex === -1) return null;
  for (let i = 1; i < values.length; i++) {
    if (String(values[i][idIndex]) === String(recordId)) {
      return values[i][eventIndex] || null;
    }
  }
  return null;
};

const updateCalendarEventIdInSheet = (sheetName, recordId, eventId) => {
  if (!recordId || !eventId) return;
  const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(sheetName);
  if (!sheet) return;
  ensureCalendarEventColumn(sheet);
  const dataRange = sheet.getDataRange();
  const values = dataRange.getValues();
  const headers = values[0];
  const idIndex = headers.indexOf("id");
  const eventIndex = headers.indexOf(CALENDAR_EVENT_COLUMN);
  if (idIndex === -1 || eventIndex === -1) return;
  for (let i = 1; i < values.length; i++) {
    if (String(values[i][idIndex]) === String(recordId)) {
      sheet
        .getRange(i + 1, eventIndex + 1)
        .setValue(eventId);
      break;
    }
  }
};

const deleteCalendarEvent = (sheetName, recordId) => {
  if (!recordId) return;
  const eventId = getCalendarEventIdFromSheet(sheetName, recordId);
  if (!eventId) return;
  try {
    const calendar = getCalendar();
    const event = calendar.getEventById(eventId);
    if (event) {
      event.deleteEvent();
    }
    clearCalendarEventId(sheetName, recordId);
  } catch (error) {
    console.error("Unable to delete calendar event", error);
  }
};

const clearCalendarEventId = (sheetName, recordId) => {
  const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(sheetName);
  if (!sheet) return;
  ensureCalendarEventColumn(sheet);
  const dataRange = sheet.getDataRange();
  const values = dataRange.getValues();
  const headers = values[0];
  const idIndex = headers.indexOf("id");
  const eventIndex = headers.indexOf(CALENDAR_EVENT_COLUMN);
  if (idIndex === -1 || eventIndex === -1) return;
  for (let i = 1; i < values.length; i++) {
    if (String(values[i][idIndex]) === String(recordId)) {
      sheet.getRange(i + 1, eventIndex + 1).setValue("");
      break;
    }
  }
};

const getCalendar = () => {
  return CALENDAR_ID
    ? CalendarApp.getCalendarById(CALENDAR_ID)
    : CalendarApp.getDefaultCalendar();
};

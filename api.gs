const SPREADSHEET_ID = "1AxZWkXf1BH2SW6awq0vchyl9OZVanCtxVXKJ_HW_fD4";
const CALENDAR_ID = null; // Set to a specific calendar ID if you don't want to use the default one.
const CALENDAR_EVENT_COLUMN = "calendarEventId";
const DEFAULT_EVENT_DURATION_MINUTES = 60;
const DEFAULT_CALENDAR_RANGE_DAYS = 7;
const DEFAULT_SHEET_NAME = "Tasks";
const DEFAULT_TASK_HEADERS = [
  "id",
  "title",
  "description",
  "category",
  "startDate",
  "startTime",
  "endDate",
  "endTime",
  "duration",
  "status",
  CALENDAR_EVENT_COLUMN,
];
const FINANCE_SHEET_NAME = "Finance";
const FINANCE_HEADERS = [
  "id",
  "date",
  "type",
  "category",
  "description",
  "amount",
];

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
    sheet.getRange(1, 1, 1, DEFAULT_TASK_HEADERS.length).setValues([DEFAULT_TASK_HEADERS]);
    sheet.setFrozenRows(1);
  } else {
    ensureTaskSheetStructure(sheet);
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
    const normalized = normalizeTaskPayload(data);
    const existingEventId =
      data[CALENDAR_EVENT_COLUMN] ||
      data.eventId ||
      getCalendarEventIdFromSheet(sheetName, data.id);

    const eventTitle =
      normalized.title ||
      normalized.name ||
      normalized.taskName ||
      normalized.summary ||
      `${sheetName || "Task"}${data.id ? ` #${data.id}` : ""}`;
    const startDate = buildDate(
      normalized.startDate || normalized.start || normalized.date || normalized.dueDate,
      normalized.startTime || normalized.time
    );
    const endDate =
      buildEndDate(
        normalized.endDate || normalized.end || normalized.dueDate || normalized.date,
        normalized.endTime,
        startDate
      ) || (startDate ? new Date(startDate.getTime() + DEFAULT_EVENT_DURATION_MINUTES * 60 * 1000) : null);

    if (!eventTitle || !startDate || !endDate) {
      return null;
    }

    const calendar = getCalendar();

    const description =
      normalized.description || normalized.details || normalized.notes || normalized.comment || "";
    const location = normalized.location || "";

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
  ensureTaskSheetStructure(sheet);
  const recordId = id || data?.id;

  if (type === "delete") {
    deleteCalendarEvent(sheetName, recordId);
    return;
  }

  if (!data) return;
  if (!data.id && recordId) {
    data.id = recordId;
  }
  const eventId = syncTaskToCalendar(sheetName, data);
  if (eventId && recordId) {
    updateCalendarEventIdInSheet(sheetName, recordId, eventId);
  }
};

const ensureCalendarEventColumn = (sheet) => {
  ensureTaskSheetStructure(sheet);
};

const getCalendarEventIdFromSheet = (sheetName, recordId) => {
  if (!recordId) return null;
  const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(sheetName);
  if (!sheet) return null;
  ensureTaskSheetStructure(sheet);
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
  ensureTaskSheetStructure(sheet);
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
  ensureTaskSheetStructure(sheet);
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

const getCalendarEventsForRange = (options = {}) => {
  const calendar = options.calendarId
    ? CalendarApp.getCalendarById(options.calendarId)
    : getCalendar();
  const start = options.startDate ? new Date(options.startDate) : new Date();
  const end = options.endDate
    ? new Date(options.endDate)
    : new Date(start.getTime() + DEFAULT_CALENDAR_RANGE_DAYS * 24 * 60 * 60 * 1000);
  const events = calendar.getEvents(start, end);
  const payload = events.map((event) => ({
    id: event.getId(),
    title: event.getTitle(),
    start: event.getStartTime()?.toISOString(),
    end: event.getEndTime()?.toISOString(),
    description: event.getDescription(),
    location: event.getLocation(),
    url: event.getUrl ? event.getUrl() : "",
    attendees: event.getGuestList
      ? event
          .getGuestList()
          .map((guest) => ({ email: guest.getEmail(), name: guest.getName() }))
      : [],
  }));
  return JSON.stringify(payload);
};

const scheduleTask = (payload = {}) => {
  const sheetName = payload.sheetName || DEFAULT_SHEET_NAME;
  let sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(sheetName);
  if (!sheet) {
    sheet = SpreadsheetApp.openById(SPREADSHEET_ID).insertSheet(sheetName);
  }
  ensureTaskSheetStructure(sheet);
  const normalized = normalizeTaskPayload(payload);
  if (!normalized.startDate || !normalized.startTime || !normalized.endTime) {
    throw new Error("Start date, start time, and end time are required.");
  }
  const start = buildDate(normalized.startDate, normalized.startTime);
  const end = buildEndDate(normalized.endDate || normalized.startDate, normalized.endTime, start);
  if (!start || !end || end <= start) {
    throw new Error("Please provide a valid time range.");
  }
  normalized.duration =
    normalized.duration ||
    calculateDurationHours(start, end).toFixed(2);
  normalized.status = normalized.status || "Planned";
  executeAction({
    action: "add",
    sheetName,
    data: normalized,
  });
  return {
    success: true,
    taskId: normalized.id,
    message: "Task scheduled.",
  };
};

const getScheduledTasks = ({ sheetName = DEFAULT_SHEET_NAME, date = null } = {}) => {
  if (!date) {
    return getScheduledTasksInRange({ sheetName });
  }
  const formatted = formatDateString(date);
  return getScheduledTasksInRange({
    sheetName,
    startDate: formatted,
    endDate: formatted,
  });
};

const getScheduledTasksInRange = ({
  sheetName = DEFAULT_SHEET_NAME,
  startDate = null,
  endDate = null,
} = {}) => {
  const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(sheetName);
  if (!sheet) return JSON.stringify([]);
  ensureTaskSheetStructure(sheet);
  const range = sheet.getDataRange();
  const values = range.getValues();
  if (values.length <= 1) return JSON.stringify([]);
  const headers = values.shift().map((header) => String(header || "").trim());
  const startKey = startDate ? formatDateString(startDate) : null;
  const endKey = endDate ? formatDateString(endDate) : null;
  const tasks = values
    .filter((row) => row.some((cell) => cell !== ""))
    .map((row) => mapRow(headers, row))
    .filter((row) => {
      if (!startKey && !endKey) return true;
      const rowDate = formatDateString(row.startDate || row.date || row.dueDate || "");
      if (!rowDate) return false;
      if (startKey && rowDate < startKey) return false;
      if (endKey && rowDate > endKey) return false;
      return true;
    })
    .map((row) => {
      const start = buildDate(row.startDate || row.date, row.startTime);
      const end = buildDate(row.endDate || row.date, row.endTime);
      return {
        id: row.id,
        title: row.title || row.name || row.task || "Untitled task",
        description: row.description || row.notes || "",
        category: row.category || "",
        status: row.status || "",
        start: start ? start.toISOString() : "",
        end: end ? end.toISOString() : "",
        calendarEventId: row[CALENDAR_EVENT_COLUMN] || "",
      };
    })
    .sort((a, b) => {
      if (!a.start || !b.start) return 0;
      return new Date(a.start) - new Date(b.start);
    });
  return JSON.stringify(tasks);
};

const normalizeTaskPayload = (data = {}) => {
  const normalized = { ...data };
  normalized.title =
    data.title || data.name || data.task || data.taskName || "Untitled task";
  normalized.description =
    data.description || data.details || data.notes || data.comment || "";
  normalized.category = data.category || data.type || "";
  normalized.startDate = formatDateString(data.startDate || data.date || data.dueDate || new Date());
  normalized.endDate = formatDateString(data.endDate || data.date || data.dueDate || normalized.startDate);
  normalized.startTime = formatTimeString(data.startTime || data.time || "09:00");
  normalized.endTime = formatTimeString(data.endTime || data.finishTime || "");
  if (!normalized.endTime && normalized.startTime) {
    const start = buildDate(normalized.startDate, normalized.startTime);
    if (start) {
      const fallbackEnd = new Date(start.getTime() + DEFAULT_EVENT_DURATION_MINUTES * 60 * 1000);
      normalized.endTime = formatTimeString(
        Utilities.formatDate(fallbackEnd, Session.getScriptTimeZone(), "HH:mm")
      );
      normalized.endDate = formatDateString(fallbackEnd);
    }
  }
  normalized.location = data.location || "";
  return normalized;
};

const formatDateString = (value) => {
  if (!value) return "";
  if (value instanceof Date) {
    return Utilities.formatDate(value, Session.getScriptTimeZone(), "yyyy-MM-dd");
  }
  const parsed = new Date(value);
  if (isNaN(parsed.getTime())) {
    return "";
  }
  return Utilities.formatDate(parsed, Session.getScriptTimeZone(), "yyyy-MM-dd");
};

const formatTimeString = (value) => {
  if (!value) return "";
  if (value instanceof Date) {
    return Utilities.formatDate(value, Session.getScriptTimeZone(), "HH:mm");
  }
  const stringValue = String(value).trim();
  if (!stringValue) return "";
  if (/^\d{1,2}:\d{2}$/.test(stringValue)) {
    const [hours, minutes] = stringValue.split(":");
    const normalizedHours = (`0${hours}`).slice(-2);
    return `${normalizedHours}:${minutes}`;
  }
  const parsed = new Date(stringValue);
  if (isNaN(parsed.getTime())) {
    return "";
  }
  return Utilities.formatDate(parsed, Session.getScriptTimeZone(), "HH:mm");
};

const calculateDurationHours = (start, end) => {
  if (!start || !end) return 0;
  const diff = Math.max(end.getTime() - start.getTime(), 0);
  return diff / (1000 * 60 * 60);
};

const mapRow = (headers, row) => {
  return headers.reduce((acc, header, idx) => {
    acc[header] = row[idx];
    return acc;
  }, {});
};

const ensureTaskSheetStructure = (sheet) => {
  if (!sheet) return;
  if (sheet.getLastRow() === 0) {
    sheet.getRange(1, 1, 1, DEFAULT_TASK_HEADERS.length).setValues([DEFAULT_TASK_HEADERS]);
    sheet.setFrozenRows(1);
    return;
  }
  let lastColumn = sheet.getLastColumn() || DEFAULT_TASK_HEADERS.length;
  const headerRange = sheet.getRange(1, 1, 1, lastColumn);
  const current = headerRange.getValues()[0];
  const normalized = current.map((header) => String(header || "").trim());
  normalized.forEach((header, idx) => {
    if (header.toLowerCase() === "name") {
      sheet.getRange(1, idx + 1).setValue("title");
      normalized[idx] = "title";
    }
  });
  DEFAULT_TASK_HEADERS.forEach((required) => {
    if (!normalized.includes(required)) {
      sheet.insertColumnAfter(lastColumn);
      lastColumn += 1;
      sheet.getRange(1, lastColumn).setValue(required);
      normalized.push(required);
    }
  });
};

const deleteTask = ({ sheetName = DEFAULT_SHEET_NAME, id }) => {
  if (!id) throw new Error("Task id is required.");
  executeAction({
    action: "delete",
    sheetName,
    id,
  });
  return { success: true };
};

const ensureFinanceSheetStructure = () => {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  let sheet = ss.getSheetByName(FINANCE_SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(FINANCE_SHEET_NAME);
    sheet.getRange(1, 1, 1, FINANCE_HEADERS.length).setValues([FINANCE_HEADERS]);
    sheet.setFrozenRows(1);
    return sheet;
  }
  if (sheet.getLastRow() === 0) {
    sheet.getRange(1, 1, 1, FINANCE_HEADERS.length).setValues([FINANCE_HEADERS]);
    sheet.setFrozenRows(1);
  }
  let lastColumn = sheet.getLastColumn() || FINANCE_HEADERS.length;
  const headerRange = sheet.getRange(1, 1, 1, lastColumn);
  const current = headerRange.getValues()[0].map((header) => String(header || "").trim());
  FINANCE_HEADERS.forEach((required) => {
    if (!current.includes(required)) {
      sheet.insertColumnAfter(lastColumn);
      lastColumn += 1;
      sheet.getRange(1, lastColumn).setValue(required);
    }
  });
  return sheet;
};

const addFinanceEntry = (entry = {}) => {
  const sheet = ensureFinanceSheetStructure();
  const payload = {
    date: formatDateString(entry.date || new Date()),
    type: entry.type || "Expense",
    category: entry.category || "",
    description: entry.description || "",
    amount: Number(entry.amount) || 0,
  };
  executeAction({
    action: "add",
    sheetName: FINANCE_SHEET_NAME,
    data: payload,
    headers: FINANCE_HEADERS,
  });
  return { success: true };
};

const updateTaskStatus = ({ id, status, sheetName = DEFAULT_SHEET_NAME }) => {
  if (!id) throw new Error("Task id is required.");
  const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(sheetName);
  if (!sheet) throw new Error("Sheet not found.");
  ensureTaskSheetStructure(sheet);
  const dataRange = sheet.getDataRange();
  const values = dataRange.getValues();
  if (values.length <= 1) return { success: false };
  const headers = values[0];
  const idIndex = headers.indexOf("id");
  const statusIndex = headers.indexOf("status");
  if (idIndex === -1 || statusIndex === -1) throw new Error("Missing headers.");
  for (let i = 1; i < values.length; i++) {
    if (String(values[i][idIndex]) === String(id)) {
      values[i][statusIndex] = status || "";
      sheet
        .getRange(i + 1, statusIndex + 1)
        .setValue(status || "");
      return { success: true };
    }
  }
  return { success: false };
};

const scheduleRecurringTasks = ({
  occurrences = 1,
  frequency = "none",
  ...payload
}) => {
  const normalizedFrequency = (frequency || "none").toLowerCase();
  const increments = {
    daily: 1,
    weekly: 7,
    biweekly: 14,
  };
  let created = 0;
  let currentDate = new Date(payload.date || payload.startDate || new Date());
  for (let i = 0; i < Math.max(1, occurrences); i++) {
    const cloned = {
      ...payload,
      date: formatDateString(currentDate),
      startDate: formatDateString(currentDate),
    };
    scheduleTask(cloned);
    created++;
    const increment = increments[normalizedFrequency];
    if (!increment) break;
    currentDate = new Date(currentDate.getTime() + increment * 24 * 60 * 60 * 1000);
  }
  return { success: true, created };
};

const getFinanceEntries = () => {
  const sheet = ensureFinanceSheetStructure();
  const range = sheet.getDataRange();
  const values = range.getValues();
  if (values.length <= 1) return JSON.stringify([]);
  const headers = values.shift();
  const entries = values
    .filter((row) => row.some((cell) => cell !== ""))
    .map((row) =>
      headers.reduce((acc, header, idx) => {
        acc[String(header || "").trim()] = row[idx];
        return acc;
      }, {})
    )
    .map((row) => ({
      id: row.id,
      date: row.date ? formatDateString(row.date) : "",
      type: row.type || row["Type (Income/Expense)"] || "",
      category: row.category || row.Category || "",
      description: row.description || row.Description || "",
      amount: Number(row.amount || row.Amount || 0),
    }))
    .sort((a, b) => {
      const aDate = new Date(a.date || 0).getTime();
      const bDate = new Date(b.date || 0).getTime();
      return bDate - aDate;
    });
  return JSON.stringify(entries);
};

const deleteFinanceEntry = (id) => {
  if (!id) throw new Error("Entry id is required.");
  executeAction({
    action: "delete",
    sheetName: FINANCE_SHEET_NAME,
    id,
  });
  return { success: true };
};

const syncMissingTaskEvents = ({ sheetName = DEFAULT_SHEET_NAME } = {}) => {
  const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(sheetName);
  if (!sheet) return { synced: 0 };
  ensureTaskSheetStructure(sheet);
  const dataRange = sheet.getDataRange();
  const values = dataRange.getValues();
  if (values.length <= 1) return { synced: 0 };
  const headers = values[0].map((header) => String(header || "").trim());
  const idIndex = headers.indexOf("id");
  const calendarIndex = headers.indexOf(CALENDAR_EVENT_COLUMN);
  if (idIndex === -1 || calendarIndex === -1) return { synced: 0 };
  let synced = 0;
  for (let i = 1; i < values.length; i++) {
    const row = values[i];
    const recordId = row[idIndex];
    const currentEvent = row[calendarIndex];
    if (!recordId || currentEvent) continue;
    const record = mapRow(headers, row);
    record.id = recordId;
    const eventId = syncTaskToCalendar(sheetName, record);
    if (eventId) {
      sheet.getRange(i + 1, calendarIndex + 1).setValue(eventId);
      synced++;
    }
  }
  return { synced };
};

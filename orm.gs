class ORM {
  constructor(sheetName, headers = []) {
    this.sheetName = sheetName;
    this.sheet = SpreadsheetService.getSheetOrCreate(sheetName, headers);
    this.headers = this.ensureHeaders(headers);
  }

  ensureHeaders(headers = []) {
    const normalized = SpreadsheetService.normalizeHeaders(headers);
    const lastColumn = this.sheet.getLastColumn();

    if (!lastColumn) {
      this.sheet.getRange(1, 1, 1, normalized.length).setValues([normalized]);
      return normalized;
    }

    const current = this.sheet.getRange(1, 1, 1, lastColumn).getValues()[0];
    const hasAnyHeader = current.some((value) => value !== "");
    if (!hasAnyHeader) {
      this.sheet.getRange(1, 1, 1, normalized.length).setValues([normalized]);
      return normalized;
    }

    const missing = normalized.filter((header) => !current.includes(header));
    if (missing.length) {
      this.sheet
        .getRange(1, current.length + 1, 1, missing.length)
        .setValues([missing]);
      return [...current, ...missing];
    }

    return current;
  }

  getHeaders() {
    if (!this.headers || !this.headers.length) {
      this.headers = this.ensureHeaders();
    }
    return this.headers;
  }

  create(data) {
    const id = this.getNextId() || 1;
    data["id"] = id;
    const headers = this.getHeaders();

    if (this.sheet.getLastColumn() !== headers.length) {
      this.sheet
        .getRange(1, 1, 1, headers.length)
        .setValues([this.getHeaders()]);
    }
    const dataRange = this.sheet.getDataRange();
    const values = dataRange.getValues();
    const newRow = [];

    for (const header of headers) {
      newRow.push(data[header] || "");
    }
    this.sheet.appendRow(newRow);
    values.push(newRow);
    return values;
  }

  readAll() {
    const dataRange = this.sheet.getDataRange();
    return dataRange.getValues();
  }

  readAllObjects() {
    return SpreadsheetService.rowsToObjects(this.readAll());
  }

  readById(id) {
    if (!id) throw new Error("ID is required");
    const dataRange = this.sheet.getDataRange();
    const values = dataRange.getValues();
    const headers = values[0];
    for (let i = 1; i < values.length; i++) {
      if (values[i][0] == id) {
        const record = {};
        for (let j = 0; j < headers.length; j++) {
          record[headers[j]] = values[i][j];
        }
        return record;
      }
    }
    return null;
  }

  updateById(id, data) {
    if (!id) throw new Error("ID is required");
    const dataRange = this.sheet.getDataRange();
    const values = dataRange.getValues();
    const headers = this.getHeaders();
    for (let i = 1; i < values.length; i++) {
      if (values[i][0] == id) {
        for (const key in data) {
          const columnIndex = headers.indexOf(key);
          if (columnIndex !== -1) {
            values[i][columnIndex] = data[key];
          }
        }
        dataRange.setValues(values);
        break;
      }
    }
    return values;
  }

  deleteById(id) {
    if (!id) throw new Error("ID is required");
    const lastRow = this.sheet.getLastRow();
    if (lastRow < 2) {
      return [];
    }
    for (let i = 2; i <= lastRow; i++) {
      const cellValue = this.sheet.getRange(i, 1).getValue();
      if (cellValue == id) {
        this.sheet.deleteRow(i);
        break;
      }
    }

    return this.sheet.getDataRange().getValues();
  }

  getNextId() {
    const dataRange = this.sheet.getDataRange();
    const values = dataRange.getValues();
    let maxId = 0;
    for (let i = 1; i < values.length; i++) {
      const id = values[i][0];
      if (id > maxId) {
        maxId = id;
      }
    }
    return maxId + 1;
  }
}

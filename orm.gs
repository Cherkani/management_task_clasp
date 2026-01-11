// Define your ORM class
class ORM {
  constructor(sheetName, headers = null) {
    const sheet =
      SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
    if (!sheet) {
      this.sheet = SpreadsheetApp.getActiveSpreadsheet().insertSheet(sheetName);
      headers = ["id", ...headers];
      this.sheet.appendRow(headers);
    } else {
      this.sheet = sheet;
    }
  }

  // Create a new record
  create(data) {
    const id = this.getNextId() || 1;
    data["id"] = id;
    const dataRange = this.sheet.getDataRange();
    const values = dataRange.getValues();
    const headers = values[0];
    const newRow = [];

    for (const header of headers) {
      newRow.push(data[header] || "");
    }
    this.sheet.appendRow(newRow);
    values.push(newRow);
    return values;
  }

  // Read all records
  readAll() {
    const dataRange = this.sheet.getDataRange();
    return dataRange.getValues();
    // const headers = values[0];
    // const records = [];
    // //Returning data from multi-dimensional array
    // for (let i = 1; i < values.length; i++) {
    //   const record = {};
    //   for (let j = 0; j < headers.length; j++) {
    //     record[headers[j]] = values[i][j];
    //   }
    //   records.push(record);
    // }
    // return records;
  }

  // Read a specific record by ID
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

  // Update a record by ID
  updateById(id, data) {
    if (!id) throw new Error("ID is required");
    const dataRange = this.sheet.getDataRange();
    const values = dataRange.getValues();
    const headers = values[0];
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

  // Delete a record by ID
  deleteById(id) {
    if (!id) throw new Error("ID is required");
    const dataRange = this.sheet.getDataRange();
    let values = dataRange.getValues();
    for (let i = 1; i < values.length; i++) {
      if (values[i][0] == id) {
        this.sheet.deleteRow(i + 1);
        values = dataRange.getValues();
        values.pop();
        break;
      }
    }

    return values;
  }

  // Get the next ID
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

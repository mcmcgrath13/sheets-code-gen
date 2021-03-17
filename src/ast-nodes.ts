// build out classes for ast nodes here or maybe just use type script
class Workbook {
  constructor(ss: SpreadsheetApp.Spreadsheet) {
    this.name = ss.getName();
    this.url = ss.getUrl();
    this.ranges = [];

    const sheets = ss.getSheets();
    for (const i in sheets) {
      let sheet = sheets[i];
      let sheetName = sheet.getName();
      let range = sheet.getDataRange();
      let rows = range.getNumRows();
      for (const i of utils.range(range.getNumRows())) {
        for (const j of utils.range(range.getNumColumns())) {
          let cell = range.getCell(i + 1, j + 1);
          let r = new Range(cell, sheetName);
          this.ranges.push(r);
        }
      }
    }
  }
}

class Range {
  constructor(range: SpreadsheetApp.Range, sheetName: string) {
    this.row = range.getRow();
    this.column = range.getColumn();
    this.numRows = range.getNumRows();
    this.numColumns = range.getNumColumns();
    this.sheet = sheetName;
    this.formula = new Formula(range.getFormula());
    this.format = range.getNumberFormat();
    this.values = range.getValues();
    this.name = ""; // not accessible from range, TODO: figure out
    this.note = range.getNote();
  }

  isCell() {
    return this.numRows === 1 && this.numColumns === 1;
  }
}

class Formula {
  constructor(formula: string) {
    Logger.log(formula);
    this.parsed = parseFormula(formula);
    Logger.log(this.parsed);
  }
}

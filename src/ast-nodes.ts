// build out classes for ast nodes here or maybe just use type script
class Workbook {
  constructor(ss: SpreadsheetApp.Spreadsheet) {
    this.name = ss.getName();
    this.url = ss.getUrl();
    this.sheets = [];
    this.ranges = [];

    const sheets = ss.getSheets();
    for (const i in sheets) {
      let s = sheets[i];
      let sheet = new Sheet(s);
      this.sheets.push(sheet);
      for (const i of utils.range(sheet.numRows)) {
        for (const j of utils.range(sheet.numColumns)) {
          let cell = s.getRange(i + 1, j + 1);
          if (cell.getFormula()) {
            let range = new Range(cell, sheet);
            this.ranges.push(range);
          }
        }
      }
    }
  }
}

class Sheet {
  constructor(sheet: SpreadsheetApp.Sheet) {
    this.name = sheet.getName();
    this.values = sheet.getDataRange().getValues(); // TODO: how to handle values which are the result of a formula - leave in? reference formula? reference range?
    this.numRows = sheet.getLastRow();
    this.numColumns = sheet.getLastColumn();
  }
}

class Range {
  constructor(range: SpreadsheetApp.Range, sheet: Sheet) {
    this.row = range.getRow();
    this.column = range.getColumn();
    this.numRows = 1;
    this.numColumns = 1;
    this.sheet = sheet;
    let formulaTokens = parseFormula(range.getFormula());
    this.formula = new Formula(formulaTokens, {
      value: "__TOP__",
      type: "toptoken",
      subtype: "start",
    });
    this.format = range.getNumberFormat();
    this.name = ""; // not accessible from range, TODO: figure out
    this.note = range.getNote();
  }

  isCell() {
    return this.numRows === 1 && this.numColumns === 1;
  }
}

interface FormulaToken {
  type: String;
  subtype: String;
  value: String;
}

class Formula {
  constructor(formulaTokens: FormulaToken[], head: FormulaToken) {
    this.head = head.value === "" ? head.type : head.value;
    this.args = [];
    while (formulaTokens.length > 0) {
      let token = formulaTokens.shift();
      if (token.subtype === "start") {
        this.args.push(new Formula(formulaTokens, token));
      } else if (token.subtype === "stop") {
        break;
      } else {
        this.args.push(token.value);
      }
    }
  }

  print() {
    let str = "";

    if (this.head !== "__TOP__" && this.head !== "subexpression") {
      str += this.head;
    }
    if (this.head !== "__TOP__") {
      str += "(";
    }

    str += this.args
      .map((arg) => {
        return arg instanceof Formula ? arg.print() : arg;
      })
      .join("");

    if (this.head !== "__TOP__") {
      str += ")";
    }

    return str;
  }
}

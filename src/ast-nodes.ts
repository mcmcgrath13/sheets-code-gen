// build out classes for ast nodes here or maybe just use type script
class Workbook {
  constructor(ss: SpreadsheetApp.Spreadsheet) {
    this.name = ss.getName();
    this.url = ss.getUrl();
    this.sheets = [];

    const sheets = ss.getSheets();
    for (const i in sheets) {
      let s = sheets[i];
      let sheet = new Sheet(s);
      this.sheets.push(sheet);
    }
  }
}

class Sheet {
  constructor(sheet: SpreadsheetApp.Sheet) {
    this.name = sheet.getName();
    this.values = sheet.getDataRange().getValues(); // TODO: how to handle values which are the result of a formula - leave in? reference formula? reference range?
    this.numRows = sheet.getLastRow();
    this.numColumns = sheet.getLastColumn();
    this.ranges = [];

    for (const i of utils.range(this.numRows)) {
      let lastRange = null;
      for (const j of utils.range(this.numColumns)) {
        let cell = sheet.getRange(i + 1, j + 1);
        if (cell.getFormula()) {
          let range = new Range(cell);
          if (!lastRange || (lastRange && !lastRange.mergeColumn(range))) {
            this.ranges.push(range);
            lastRange = range;
          }
        } else {
          lastRange = null;
        }
      }
    }

    for (var i = 0; i < this.ranges.length; i++) {
      let range = this.ranges[i];
      let neighIdx = utils.findRangeBelow(range, this.ranges);
      while (neighIdx !== -1 && range.mergeRow(this.ranges[neighIdx])) {
        this.ranges.splice(neighIdx, 1);
        neighIdx = utils.findRangeBelow(range, this.ranges);
      }
    }
  }
}

class Range {
  constructor(range: SpreadsheetApp.Range) {
    this.row = range.getRow();
    this.column = range.getColumn();
    this.numRows = 1;
    this.numColumns = 1;
    let formulaTokens = parseFormula(range.getFormulaR1C1());
    this.formula = new Formula(formulaTokens, {
      value: "__TOP__",
      type: "toptoken",
      subtype: "start",
    });
    this.format = range.getNumberFormat();
    this.name = ""; // not accessible from range, TODO: figure out
    this.note = range.getNote();
  }

  isCell(): boolean {
    return this.numRows === 1 && this.numColumns === 1;
  }

  mergeRow(otherRange: Range): boolean {
    if (
      this.formula.print() === otherRange.formula.print() &&
      this.format === otherRange.format &&
      this.name === otherRange.name
    ) {
      this.numRows += 1;
      return true;
    } else {
      return false;
    }
  }

  mergeColumn(otherRange: Range): boolean {
    if (
      this.formula.print() === otherRange.formula.print() &&
      this.format === otherRange.format &&
      this.name === otherRange.name
    ) {
      this.numColumns += 1;
      return true;
    } else {
      return false;
    }
  }
}

interface FormulaToken {
  type: String;
  subtype: String;
  value: String;
}

class Formula {
  constructor(formulaTokens: FormulaToken[], head: FormulaToken) {
    this.head = head.value === "" ? head.type : head.value.toUpperCase();
    this.args = [];
    while (formulaTokens.length > 0) {
      let token = formulaTokens.shift();
      if (token.subtype === "start") {
        this.args.push(new Formula(formulaTokens, token));
      } else if (token.subtype === "stop") {
        break;
      } else if (token.subtype === "range") {

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
        return arg['print'] ? arg.print() : arg;
      })
      .join("");

    if (this.head !== "__TOP__") {
      str += ")";
    }

    return str;
  }
}

class RangeReference {
  constructor(r1c1: String) {
    let re = /^(?:R(?<row>[0-9\-\[\]]+))?(?:C(?<column>[0-9\-\[\]]+))?$/;
    let matches = r1c1.split(':').map(r => r.match(re));
    this.start = new CellReference(matches[0].groups.row, matches[0].groups.column)
    if (matches.length === 2) {
      this.stop = new CellReference(matches[1].groups.row, matches[1].groups.column)
    } else {
      this.stop = this.start;
    }
  }

  isCell() {
    return this.start === this.stop;
  }

  print() {
    return this.isCell() ? this.start.print() : `${this.start.print()}:${this.stop.print()}`
  }
}

const parseCellRef = (str: String | undefined) => {
  if (!str) {
    return { isRelative: false, value: 0 }; // not on the grid this would be the whole row/column
  }

  let re = /^\[(?<val>\-?[0-9]+)\]$/;
  let match = str.match(re);
  if (match) {
    return { isRelative: true, value: parseInt(rowMatch.groups.val) };
  } else {
    return { isRelative: false, value: parseInt(str) };
  }
}

class CellReference {
  constructor(row: String, column: String) {
    let re = /^\[(?<val>\-?[0-9]+)\]$/;
    let rowRes = parseCellRef(row);
    this.rowIsRelative = rowRes.isRelative;
    this.row = rowRes.value;
    let columnRes = parseCellRef(column);
    this.columnIsRelative = columnRes.isRelative;
    this.column = columnRes.value;
  }

  print() {
    return 'CELL'
  }
}

const DEFAULT_NUMBER_FORMAT = "0.###############";

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
        if (
          cell.getFormula() ||
          cell.getNumberFormat() !== DEFAULT_NUMBER_FORMAT
        ) {
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

  print(ab: ?boolean) {
    return this.formula.print()
      ? this.formula.print(
          ab,
          this.row,
          this.column
        )
      : this.format;
  }

  printAddress(ab: ?boolean) {
    const format = (row, column) => {
      if (ab) {
        return utils.getAlpha(column) + row;
      } else {
        return `R${row}C${column}`;
      }
    }

    let str = format(this.row, this.column);
    if (!this.isCell()) {
      str += ':' + format(this.row + this.numRows - 1, this.column + this.numColumns - 1);
    }

    return str;
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
        this.args.push(new RangeReference(token.value));
      } else {
        this.args.push(token.value);
      }
    }
  }

  print(
    ab: ?boolean,
    row: ?number,
    column: ?number
  ) {
    let str = "";

    if (this.head !== "__TOP__" && this.head !== "subexpression") {
      str += this.head;
    }
    if (this.head !== "__TOP__") {
      str += "(";
    }

    str += this.args
      .map((arg) => {
        return arg["print"]
          ? arg.print(ab, row, column)
          : arg;
      })
      .join("");

    if (this.head !== "__TOP__") {
      str += ")";
    }

    return str;
  }
}

interface RangeMatch {
  row: String | undefined;
  column: String | undefined;
  sheet: String | undefined;
}

class RangeReference {
  constructor(r1c1: String) {
    let re = /^(?:(?<sheet>.*)!)?(?:R(?<row>[0-9\-\[\]]+))?(?:C(?<column>[0-9\-\[\]]+))?$/;
    let matches = r1c1.split(":").map((r) => r.match(re));
    this.sheet = matches[0].groups.sheet;
    this.start = new CellReference(matches[0].groups);
    if (matches.length === 2) {
      this.stop = new CellReference(matches[1].groups);
    } else {
      this.stop = this.start;
    }
  }

  isCell() {
    return this.start === this.stop;
  }

  print(
    ab: ?boolean,
    row: ?number,
    column: ?number
  ) {
    let str = this.sheet ? `${this.sheet}!` : "";
    str += this.isCell()
      ? this.start.print(ab, row, column)
      : `${this.start.print(ab, row, column)}:${this.stop.print(
          ab,
          row,
          column
        )}`;
    return str;
  }
}

class CellReference {
  constructor(match: RangeMatch) {
    let re = /^\[(?<val>\-?[0-9]+)\]$/;
    this.row = new CellAddress(match.row);
    this.column = new CellAddress(match.column);
  }

  isOnlyRow() {
    return this.row.isEmpty();
  }

  isOnlyColumn() {
    return this.column.isEmpty();
  }

  print(ab: ?boolean, row, column) {
    if (ab) {
      return `${this.column.printAB(column, false)}${this.row.printAB(
        row,
        true
      )}`;
    } else {
      return `${this.row.print("R")}${this.column.print("C")}`;
    }
  }
}

class CellAddress {
  constructor(str: String | undefined) {
    if (!str) {
      this.isRelative = false;
      this.value = 0;
      return;
    }

    let re = /^\[(?<val>\-?[0-9]+)\]$/;
    let match = str.match(re);
    if (match) {
      this.isRelative = true;
      this.value = parseInt(match.groups.val);
    } else {
      this.isRelative = false;
      this.value = parseInt(str);
    }
  }

  isEmpty() {
    return this.value === 0 && !this.isRelative;
  }

  print(prefix) {
    if (this.isEmpty()) {
      return "";
    } else if (this.isRelative) {
      return `[${prefix}${this.value}]`;
    } else {
      return `${prefix}${this.value}`;
    }
  }

  printAB(anchor: number, isRow) {
    if (this.isEmpty()) return "";
    let str = this.isRelative ? "" : "$";
    let val = this.isRelative ? this.value + anchor : this.value;
    if (isRow) {
      str += `${val}`;
    } else {
      str += utils.getAlpha(val);
    }
    Logger.log(`value: ${this.value} anchor: ${anchor} str: ${str}`);
    return str;
  }
}

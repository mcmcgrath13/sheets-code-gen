const DEFAULT_NUMBER_FORMAT = "0.###############";

// build out classes for ast nodes here or maybe just use type script
class Workbook {
  name: string;
  url: string;
  sheets: Sheet[];

  constructor(ss: GoogleAppsScript.Spreadsheet.Spreadsheet, activeOnly: boolean, tabularData: boolean) {
    this.name = ss.getName();
    this.url = ss.getUrl();
    this.sheets = [];

    if (activeOnly) {
      this.sheets.push(new Sheet(ss.getActiveSheet(), tabularData));
    } else {
      const sheets = ss.getSheets();
      for (const i in sheets) {
        let s = sheets[i];
        let sheet = new Sheet(s, tabularData);
        this.sheets.push(sheet);
      }
    }
  }
}

class Sheet {
  name: string;
  values: any[][];
  numRows: number;
  numColumns: number;
  ranges: Range[];
  table: Table;

  constructor(sheet: GoogleAppsScript.Spreadsheet.Sheet, tabularData: boolean) {
    this.name = sheet.getName();
    this.values = sheet.getDataRange().getValues();
    this.numRows = sheet.getLastRow();
    this.numColumns = sheet.getLastColumn();
    this.ranges = [];

    for (const i of utils.range(this.numRows)) {
      for (const j of utils.range(this.numColumns)) {
        let cell = sheet.getRange(i + 1, j + 1);
        if (
          cell.getFormula() ||
          cell.getNumberFormat() !== DEFAULT_NUMBER_FORMAT
        ) {
          let range = new Range(cell);
          this.ranges.push(range);
          if (!range.formula.isEmpty()) {
            this.values[i][j] = range;
          }
        }
      }
    }

    if (tabularData) {
      this.table = new Table(this.name, this.values);
    }
  }

  collapseRanges() {
    // collapse left to right
    for (var i = 0; i < this.ranges.length; i++) {
      let range = this.ranges[i];
      let neighIdx = utils.findRangeRight(range, this.ranges);
      while (neighIdx !== -1 && range.mergeColumn(this.ranges[neighIdx])) {
        this.ranges.splice(neighIdx, 1);
        neighIdx = utils.findRangeRight(range, this.ranges);
      }
    }

    // collapse top to bottom
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
  row: number;
  column: number;
  numRows: number;
  numColumns: number;
  formula: Formula;
  format: string;
  name: string;
  note: string;

  constructor(range: GoogleAppsScript.Spreadsheet.Range) {
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

  print(ab?: boolean) {
    return this.formula.print()
      ? this.formula.print(ab, this.row, this.column)
      : this.format;
  }

  printAddress(ab?: boolean) {
    const format = (row: number, column: number) => {
      if (ab) {
        return utils.getAlpha(column) + row;
      } else {
        return `R${row}C${column}`;
      }
    };

    let str = format(this.row, this.column);
    if (!this.isCell()) {
      str +=
        ":" +
        format(this.row + this.numRows - 1, this.column + this.numColumns - 1);
    }

    return str;
  }
}

class Table {
  range: RangeReference;
  headers: string[];
  dataColumns: number[];
  derivedColumns: number[];

  constructor(sheet: string, values: any[][]) {
    if (values[0][0] == '' && values[0][1] == '') {
      throw 'Table header missing: must have header in the first row, only first column can be blank.'
    }
    this.headers = [];

    let j = 0;
    let numRows = Number.MAX_SAFE_INTEGER;
    while (j < values[0].length && typeof(values[0][j]) === 'string') {
      let header = values[0][j];
      if (header === '') {
        if (j === 0) {
          header = 'row_id';
        } else {
          break;
        }
      }
      let col_type = typeof values[1][j];
      let i = 2
      let isDerived = false;
      while (typeof values[i][j] === col_type) {
        if (values[i][j] instanceof Range) {
          isDerived = true;
          // formula must match
          if (values[i][j].print() !== values[i-1][j].print()) {
            break;
          }
        }
        i++;
        if (i >= values.length) break;
      }
      if (numRows !== Number.MAX_SAFE_INTEGER && numRows !== i) {
        break;
      }
      numRows = Math.min(i, numRows);
      j++;
      this.headers.push(header);
      if (isDerived) {
        this.derivedColumns.push(j);
      } else {
        this.dataColumns.push(j);
      }
    }

    let start = new CellReference({ row: "1", column: "1"});
    let stop = new CellReference({ row: numRows.toString(), column: j.toString()});
    this.range = new RangeReference(sheet, start, stop)
  }
}

interface FormulaToken {
  type: string;
  subtype: string;
  value: string;
}

type Arg = Formula | RangeReference | TokenValue;

class Formula {
  head: string;
  args: Arg[];

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
        this.args.push(RangeReference.fromR1C1(token.value));
      } else {
        this.args.push(new TokenValue(token.value));
      }
    }
  }

  isEmpty() {
    return this.head === "__TOP__" && this.args.length === 0;
  }

  print(ab?: boolean, row?: number, column?: number) {
    let str = "";

    if (this.head !== "__TOP__" && this.head !== "subexpression") {
      str += this.head;
    }
    if (this.head !== "__TOP__") {
      str += "(";
    }

    str += this.args
      .map((arg) => {
        return arg.print(ab, row, column);
      })
      .join("");

    if (this.head !== "__TOP__") {
      str += ")";
    }

    return str;
  }
}

class TokenValue {
  value: string;

  constructor(v: string) {
    this.value = v;
  }

  print() {
    return this.value;
  }
}

interface RangeMatch {
  row?: string;
  column?: string;
  sheet?: string;
}

class RangeReference {
  sheet: string;
  start: CellReference;
  stop: CellReference;

  constructor(sheet: string, start: CellReference, stop: CellReference) {
    this.sheet = sheet;
    this.start = start;
    this.stop = stop
  }

  static fromR1C1(r1c1: string) {
    let re = /^(?:(?<sheet>.*)!)?(?:R(?<row>[0-9\-\[\]]+))?(?:C(?<column>[0-9\-\[\]]+))?$/;
    let matches = r1c1.split(":").map((r) => r.match(re));
    const sheet = matches[0].groups.sheet;
    const start = new CellReference(matches[0].groups);
    let stop;
    if (matches.length === 2) {
      stop = new CellReference(matches[1].groups);
    } else {
      stop = start;
    }
    return new RangeReference(sheet, start, stop);
  }

  isCell() {
    return this.start === this.stop;
  }

  rowExtent(start: number): number[] {
    if (this.start.isOnlyColumn()) {
      return [1, Number.MAX_VALUE];
    }
    return [
      this.start.row.value + (this.start.row.isRelative ? start : 0),
      this.stop.row.value + (this.stop.row.isRelative ? start : 0),
    ];
  }

  columnExtent(start: number): number[] {
    if (this.start.isOnlyRow()) {
      return [1, Number.MAX_VALUE];
    }
    return [
      this.start.column.value + (this.start.column.isRelative ? start : 0),
      this.stop.column.value + (this.stop.column.isRelative ? start : 0),
    ];
  }

  print(ab?: boolean, row?: number, column?: number) {
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
  row: CellAddress;
  column: CellAddress;

  constructor(match: RangeMatch) {
    this.row = new CellAddress(match.row);
    this.column = new CellAddress(match.column);
  }

  isOnlyRow() {
    return this.column.isEmpty();
  }

  isOnlyColumn() {
    return this.row.isEmpty();
  }

  print(ab?: boolean, row?: number, column?: number) {
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
  isRelative: boolean;
  value: number;

  constructor(str?: string) {
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

  isEmpty(): boolean {
    return this.value === 0 && !this.isRelative;
  }

  print(prefix: string) {
    if (this.isEmpty()) {
      return "";
    } else if (this.isRelative) {
      return `[${prefix}${this.value}]`;
    } else {
      return `${prefix}${this.value}`;
    }
  }

  printAB(anchor: number, isRow: boolean): string {
    if (this.isEmpty()) return "";
    let str = this.isRelative ? "" : "$";
    let val = this.isRelative ? this.value + anchor : this.value;
    if (isRow) {
      str += `${val}`;
    } else {
      str += utils.getAlpha(val);
    }
    return str;
  }
}

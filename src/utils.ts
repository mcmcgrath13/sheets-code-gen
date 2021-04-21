const ALPHABET = [
  "A",
  "B",
  "C",
  "D",
  "E",
  "F",
  "G",
  "H",
  "I",
  "J",
  "K",
  "L",
  "M",
  "N",
  "O",
  "P",
  "Q",
  "R",
  "S",
  "T",
  "U",
  "V",
  "W",
  "X",
  "Y",
  "Z",
];

// get an array of indexes of the given length
const utils = {
  range(length: number) {
    return [...Array(length).keys()];
  },
  findRangeBelow(range: Range, ranges: Range[]): number {
    return ranges.findIndex((r) => {
      return (
        r.column === range.column &&
        r.numColumns === range.numColumns &&
        r.row === range.row + range.numRows
      );
    });
  },
  findRangeRight(range: Range, ranges: Range[]): number {
    return ranges.findIndex((r) => {
      return (
        r.row === range.row &&
        r.numRows === range.numRows &&
        r.column === range.column + range.numRows
      );
    });
  },
  getAlpha(num: number, str?: string): string {
    if (num < 0) return str;
    const idx = (num - 1) % ALPHABET.length;
    let s = str ? ALPHABET[idx] + str : ALPHABET[idx];
    return this.getAlpha(num - ALPHABET.length, s);
  },
};

const langs = {
  gs: {
    print(ast: Workbook) {
      return ast.sheets
        .map((sheet) => {
          sheet.collapseRanges();
          return (
            `${sheet.name}\n  ` +
            sheet.ranges
              .map((cell) => `${cell.printAddress(true)} = ${cell.print(true)}`)
              .join("\n  ")
          );
        })
        .join("\n");
    },
  },
  ast: {
    print(ast: Workbook) {
      return JSON.stringify(ast);
    },
  },
  jl: {
    toSnakeCase(str: string): string {
      return str.toLowerCase().replace(/\W/g, "_");
    },
    getVal(v: any): string {
      if (v === "") {
        return "nothing";
      } else if (typeof v === "string" || v instanceof String) {
        return `"${v}"`;
      } else {
        return v.toString();
      }
    },
    val2Str(v: any, sheetName: string): string {
      return v instanceof Range ? this.range2var(v, sheetName) : v;
    },
    vals2Str(vals: any[], sheetName: string): string {
      return (
        "[" +
        vals
          .map((row) => row.map((v) => this.val2Str(v, sheetName)).join(" "))
          .join("; ") +
        "]"
      );
    },
    ab2Var(sheet: string, address: string): string {
      if (address.includes("!")) {
        return this.toSnakeCase(address)
          .replace(/:/g, "")
          .replace(/\$/g, "")
          .replace(/!/g, "_");
      }
      return (
        this.toSnakeCase(sheet) +
        "_" +
        address.toLowerCase().replace(/:/g, "").replace(/\$/g, "")
      );
    },
    range2var(r: Range, sheetName: string): string {
      return this.ab2Var(sheetName, r.printAddress(true));
    },
    rangeRef2var(
      r: RangeReference,
      row: number,
      column: number,
      sheet: string
    ): string {
      return this.ab2Var(sheet, r.print(true, row, column));
    },
    inds2var(i: number, j: number, sheetName: string): string {
      return `${this.toSnakeCase(sheetName)}_${utils
        .getAlpha(j + 1)
        .toLowerCase()}${i + 1}`;
    },
    printFormula(f: Formula, r: Range, sheet: string): ValVar {
      let vars = [];
      let str = "";

      if (f.head !== "__TOP__" && f.head !== "subexpression") {
        str += f.head;
      }
      if (f.head !== "__TOP__") {
        str += "(";
      }

      str += f.args
        .map((arg) => {
          if (arg instanceof Formula) {
            const { text, vars: v } = this.printFormula(arg, r, sheet);
            vars.push(...v);
            return text;
          } else if (arg instanceof RangeReference) {
            let varName = this.rangeRef2var(arg, r.row, r.column, sheet);
            vars.push({
              var: varName,
              sheet: sheet,
              rowExtent: arg.rowExtent(r.row),
              columnExtent: arg.columnExtent(r.column),
            });
            return varName;
          } else {
            return arg.value;
          }
        })
        .join("");

      if (f.head !== "__TOP__") {
        str += ")";
      }

      return { text: str, vars };
    },
    printDfFormula(f: Formula, r: Range, sheet: string, table: Table): ValVar {
      let vars = [];
      let str = "";

      if (f.head !== "__TOP__" && f.head !== "subexpression") {
        str += f.head;
      }
      if (f.head !== "__TOP__") {
        str += "(";
      }

      str += f.args
        .map((arg) => {
          if (arg instanceof Formula) {
            const { text, vars: v } = this.printDfFormula(arg, r, sheet, table);
            vars.push(...v);
            return text;
          } else if (arg instanceof RangeReference) {
            let varName;
            const columnExtent = arg.columnExtent(r.column);
            const rowExtent = arg.rowExtent(r.row);
            if (table.containsRangeRef(arg, r, sheet)) {
              const tableRowExtent = table.range.rowExtent(0);
              if (columnExtent[0] === columnExtent[1]) {
                let headerIdx = table.headers[columnExtent[0] - 1];
                if (
                  rowExtent[0] === tableRowExtent[0] &&
                  rowExtent[1] === tableRowExtent[1]
                ) {
                  varName = `df.${headerIdx}`;
                } else if (
                  rowExtent[0] === rowExtent[1] &&
                  rowExtent[0] === tableRowExtent[0]
                ) {
                  varName = `row.${headerIdx}`;
                } else {
                  throw "Multi row (but not column) dataframe indexing not implemented";
                }
              } else {
                let headers = table.headers.filter(
                  (h, i) => i >= columnExtent[0] - 1 && i <= columnExtent[1] - 1
                );
                let headersIdx = `[${headers.map((h) => `:{h}`).join(", ")}]`;
                if (
                  rowExtent[0] === tableRowExtent[0] &&
                  rowExtent[1] === tableRowExtent[1]
                ) {
                  varName = `df[!, ${headersIdx}]`;
                } else if (
                  rowExtent[0] === rowExtent[1] &&
                  rowExtent[0] === tableRowExtent[0]
                ) {
                  varName = `row[!, ${headersIdx}]`;
                } else {
                  throw "Multi row (but not column) dataframe indexing not implemented";
                }
              }
            } else {
              varName = this.rangeRef2var(arg, r.row, r.column, sheet);
              vars.push({
                var: varName,
                sheet: sheet,
                rowExtent,
                columnExtent,
              });
            }

            return varName;
          } else {
            return arg.value;
          }
        })
        .join("");

      if (f.head !== "__TOP__") {
        str += ")";
      }

      return { text: str, vars };
    },
    getValVar(v: DepVar, sheetVals: Map<string, any[][]>): ValVar {
      const vals = sheetVals.get(v.sheet);
      const depVars: DepVar[] = [];

      // scalar
      if (
        v.rowExtent[0] === v.rowExtent[1] &&
        v.columnExtent[0] === v.columnExtent[1]
      ) {
        const varVal = vals[v.rowExtent[0] - 1][v.columnExtent[0] - 1];
        if (varVal instanceof Range) {
          depVars.push({
            sheet: v.sheet,
            var: this.range2var(varVal, v.sheet),
          });
          return { text: this.val2Str(varVal), vars: depVars };
        } else {
          return { text: this.getVal(varVal), vars: depVars };
        }
      }

      const varVals = [];
      for (
        let i = v.rowExtent[0] - 1;
        i < v.rowExtent[1] && i < vals.length;
        i++
      ) {
        const rowVals = [];
        for (
          let j = v.columnExtent[0] - 1;
          j < v.columnExtent[1] && j < vals[i].length;
          j++
        ) {
          const varVal = vals[i][j];
          if (varVal instanceof Range) {
            depVars.push({
              sheet: v.sheet,
              var: this.range2var(varVal, v.sheet),
            });
          } else {
            depVars.push({
              sheet: v.sheet,
              var: this.inds2var(i, j, v.sheet),
              rowExtent: [i + 1, i + 1],
              columnExtent: [j + 1, j + 1],
            });
          }
          rowVals.push(depVars.slice(-1)[0].var);
        }
        varVals.push(rowVals);
      }

      return { text: this.vals2Str(varVals, v.sheet), vars: depVars };
    },
    dfExpr(table: Table, values: any[][], sheet: string) {
      let text = "DataFrame([";
      let dataCols = table.dataColumns.map((col) => {
        return (
          `[` +
          values
            .slice(1, table.range.stop.row.value)
            .map((row) => this.getVal(row[col]))
            .join(", ") +
          "]"
        );
      });
      text +=
        dataCols.join(", ") +
        "], [" +
        table.dataColumns.map((col) => `:${table.headers[col]}`).join(", ") +
        "])";

      // TODO: add derived columns
      let derivedCols = table.derivedColumns.map((col) => {
        let r = values[1][col];
        let dfFormula = this.printDfFormula(r.formula, r, sheet, table);
        return {
          text: `df.${table.headers[col]} = map(row -> ${dfFormula.text}, eachrow(df))`,
          vars: dfFormula.vars,
        };
      });

      text += "\n\n" + derivedCols.map((d) => d.text).join("\n");
      let vars = derivedCols.reduce((acc, cur) => acc.concat(cur.vars), []);

      return { text, vars };
    },
    addToSorted(
      sorted: Expression[],
      v: ValVar,
      k: string,
      m: Map<string, ValVar>,
      sheetVals: Map<string, any[][]>
    ) {
      Logger.log(v);
      v.vars.forEach((vr) => {
        const varName = vr.var;
        if (!sorted.find((e) => e.lhs === varName)) {
          if (m.has(varName)) {
            sorted.push({ lhs: varName, rhs: m.get(varName).text });
          } else {
            const valVar = this.getValVar(vr, sheetVals);
            m.set(varName, valVar);
            this.addToSorted(sorted, valVar, varName, m, sheetVals);
          }
        }
      });
      if (!sorted.find((e) => e.lhs === k)) {
        sorted.push({ lhs: k, rhs: v.text });
      }
    },
    print(ast: Workbook): string {
      const sheetVals: Map<string, any[][]> = new Map();
      const exprs: Map<string, ValVar> = new Map();
      let tabular = false;

      ast.sheets.forEach((sheet) => {
        sheetVals.set(sheet.name, sheet.values);

        if (sheet.table) {
          tabular = true;
          let dfExpr = this.dfExpr(sheet.table, sheet.values, sheet.name);
          exprs.set("df", dfExpr);

          sheet.ranges
            .filter((cell) => cell.formula.print() !== "")
            .filter((cell) => !sheet.table.containsRange(cell))
            .forEach((cell) =>
              exprs.set(
                this.range2var(cell, sheet.name),
                this.printDfFormula(cell.formula, cell, sheet.name, sheet.table)
              )
            );
          // TODO, add expr to mapping, intervene on ranges in getValVar to see if it matches a column and substitute the column reference
        } else {
          sheet.ranges
            .filter((cell) => cell.formula.print() !== "")
            .forEach((cell) =>
              exprs.set(
                this.range2var(cell, sheet.name),
                this.printFormula(cell.formula, cell, sheet.name, sheet.table)
              )
            );
        }
      });

      const sorted: Expression[] = [];

      exprs.forEach((v, k, m) => {
        this.addToSorted(sorted, v, k, m, sheetVals);
      });

      return (
        "using SpreadsheetFunctions\n" +
        (tabular ? "using DataFrames\n" : "") +
        "\n" +
        sorted.map((s) => s.lhs + " = " + s.rhs).join("\n")
      );
    },
  },
};

interface Expression {
  lhs: string;
  rhs: string;
}

interface DepVar {
  sheet: string;
  var: string;
  rowExtent?: number[];
  columnExtent?: number[];
}

interface ValVar {
  text: string;
  vars: DepVar[];
}

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
  range(length) {
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
  getAlpha(num, str) {
    if (num < 0) return str;
    const idx = (num - 1) % ALPHABET.length;
    let s = str ? ALPHABET[idx] + str : ALPHABET[idx];
    return this.getAlpha(num - ALPHABET.length, s);
  },
};

const langs = {
  gs: {
    print(ast) {
      return ast.sheets
        .map((sheet) => {
          sheet.collapseRanges();
          return (
            `${sheet.name}\n  ` +
            sheet.ranges // TODO: collapsed view here
              .map((cell) => `${cell.printAddress(true)} = ${cell.print(true)}`)
              .join("\n  ")
          );
        })
        .join("\n");
    },
  },
  ast: {
    print(ast) {
      return JSON.stringify(ast);
    },
  },
  jl: {
    toSnakeCase(str: String) {
      return str.toLowerCase().replace(/\W/g, "_");
    },
    val2Str(v, sheetName) {
      return v
        ? v instanceof Range
          ? this.range2var(v, sheetName)
          : v.toString()
        : "nothing";
    },
    vals2Str(vals: Array, sheetName: String) {
      return (
        "[" +
        vals
          .map((row) => row.map((v) => this.val2Str(v, sheetName)).join(" "))
          .join("; ") +
        "]"
      );
    },
    ab2Var(sheet: String, address: String) {
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
    range2var(r: Range, sheetName: String) {
      return this.ab2Var(sheetName, r.printAddress(true));
    },
    printFormula(f: Formula, r: Range, sheet: String) {
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
            let abRef = arg.print(true, r.row, r.column);
            let varName = this.ab2Var(sheet, abRef);
            vars.push({
              var: varName,
              sheet: sheet,
              rowExtent: arg.rowExtent(r.row),
              columnExtent: arg.columnExtent(r.column),
            });
            return varName;
          } else {
            return arg;
          }
        })
        .join("");

      if (f.head !== "__TOP__") {
        str += ")";
      }

      return { text: str, vars };
    },
    getValVar(v, sheetVals: Map) {
      const vals = sheetVals.get(v.sheet);
      const depVars = [];

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
        }
        return { text: this.val2Str(varVal, v.sheet), vars: depVars };
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
          rowVals.push(varVal);
          if (varVal instanceof Range) {
            depVars.push({
              sheet: v.sheet,
              var: this.range2var(varVal, v.sheet),
            });
          }
        }
        varVals.push(rowVals);
      }

      return { text: this.vals2Str(varVals, v.sheet), vars: depVars };
    },
    addToSorted(sorted, v, k, m, sheetVals) {
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
    print(ast) {
      const sheetVals = new Map();
      const exprs = new Map();
      ast.sheets.forEach((sheet) => {
        sheetVals.set(sheet.name, sheet.values);

        sheet.ranges
          .filter((cell) => cell.formula.print() !== "")
          .forEach((cell) =>
            exprs.set(
              this.range2var(cell, sheet.name),
              this.printFormula(cell.formula, cell, sheet.name)
            )
          );
      });

      const sorted = [];

      exprs.forEach((v, k, m) => {
        this.addToSorted(sorted, v, k, m, sheetVals);
      });

      Logger.log(sorted);

      return sorted.map((s) => s.lhs + " = " + s.rhs).join("\n");
    },
  },
};

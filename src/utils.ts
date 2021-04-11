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
  getAlpha(num: number, str?: string): string {
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
    print(ast) {
      return JSON.stringify(ast);
    },
  },
  jl: {
    toSnakeCase(str: String) {
      return str.toLowerCase().replace(/\W/g, "_");
    },
    sheetVals2Arr(sheet: Sheet) {
      return (
        "[\n    " +
        sheet.values
          .map((row) =>
            row.map((v) => (v ? v.toString() : "nothing")).join(" ")
          )
          .join(";\n    ") +
        "\n  ]"
      );
    },
    range2var(r: Range, sheetName: String) {
      const sheetVar = this.toSnakeCase(sheetName);
      return (
        sheetVar + "_" + r.printAddress(true).toLowerCase().replace(/:/g, "_")
      );
    },
    printFormula(f: Formula, r: Range, valVar: String) {
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
            return this.printFormula(arg, r, valVar);
          } else if (arg instanceof RangeReference) {
            let valArr;
            if (arg.sheet) {
              valArr = this.toSnakeCase(arg.sheet) + "_vals";
            } else {
              valArr = valVar;
            }
            let idx;
            if (arg.isCell()) {
              if (arg.start.isOnlyRow()) {
                let row = arg.start.row.isRelative
                  ? r.row + arg.start.row.value
                  : arg.start.row.value;
                idx = row.toString() + ", :";
              } else if (arg.start.isOnlyColumn()) {
                let column = arg.start.column.isRelative
                  ? r.column + arg.start.column.value
                  : arg.start.column.value;
                idx = ":, " + column.toString();
              } else {
                let row = arg.start.row.isRelative
                  ? r.row + arg.start.row.value
                  : arg.start.row.value;
                let column = arg.start.column.isRelative
                  ? r.column + arg.start.column.value
                  : arg.start.column.value;
                idx = row.toString() + ", " + column.toString();
              }
              // TODO: how to translate range reference into inds - also what if val is really a var?
            } else {
              idx = ":,:";
            }
            return valArr + "[" + idx + "]";
          } else {
            return arg;
          }
        })
        .join("");

      if (f.head !== "__TOP__") {
        str += ")";
      }

      return str;
    },
    print(ast) {
      return ast.sheets
        .map((sheet) => {
          const valArrName = this.toSnakeCase(sheet.name) + "_vals";
          const valExpr = valArrName + " = " + this.sheetVals2Arr(sheet) + "\n";
          return (
            valExpr +
            sheet.ranges
              .filter((cell) => cell.formula.print() !== "")
              .map(
                (cell) =>
                  `${this.range2var(cell, sheet.name)} = ${this.printFormula(
                    cell.formula,
                    cell,
                    valArrName
                  )}`
              )
              .join("\n")
          );
        })
        .join("\n");
    },
  },
};

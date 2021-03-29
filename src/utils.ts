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
  getAlpha(num, str) {
    if (num < 0) return str;
    const idx = (num - 1) % ALPHABET.length;
    let s = str ? ALPHABET[idx] + str : ALPHABET[idx];
    return this.getAlpha(num - ALPHABET.length, s);
  },
  langs: {
    gs: {
      print(ast) {
        return ast.sheets
          .map((sheet) => {
            return `${sheet.name}\n  ` + sheet.ranges
              .map(
                (cell) =>
                  `${cell.printAddress(true)} = ${cell.print(true)}`
              )
              .join("\n  ");
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
      print(ast) {
        return "not implemented";
      },
    },
  },
};

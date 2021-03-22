// get an array of indexes of the given length
const utils = {
  range(length) {
    return [...Array(length).keys()];
  },
  findRangeBelow(range: Range, ranges: Range[]): number {
    return ranges.findIndex((r) => {
      return (
        r.column == range.column &&
        r.numColumns === range.numColumns &&
        r.row === range.row + range.numRows
      );
    });
  },
  langs: {
    gs: {
      print(ast) {
        return ast.sheets
          .map((sheet) => {
            return sheet.ranges
              .map(
                (cell) =>
                  `${sheet.name} (${cell.row}${
                    cell.numRows > 1 ? ":" + (cell.row + cell.numRows - 1) : ""
                  }, ${cell.column}${
                    cell.numColumns > 1
                      ? ":" + (cell.column + cell.numColumns - 1)
                      : ""
                  }) = ${cell.formula.print()}`
              )
              .join("\n");
          })
          .join("\n");
      },
    },
    ast: {
      print(ast) {
        return "not implemented";
      },
    },
    jl: {
      print(ast) {
        return "not implemented";
      },
    },
  },
};

/**
 * Creates a menu entry in the Google Sheets UI when the document is opened.
 * This method is only used by the regular add-on, and is never called by
 * the mobile add-on version.
 *
 * @param {object} e The event parameter for a simple onOpen trigger. To
 *     determine which authorization mode (ScriptApp.AuthMode) the trigger is
 *     running in, inspect e.authMode.
 */
function onOpen(e) {
  SpreadsheetApp.getUi()
    .createAddonMenu()
    .addItem("Start", "showSidebar")
    .addToUi();
}

/**
 * Runs when the add-on is installed.
 * This method is only used by the regular add-on, and is never called by
 * the mobile add-on version.
 *
 * @param {object} e The event parameter for a simple onInstall trigger. To
 *     determine which authorization mode (ScriptApp.AuthMode) the trigger is
 *     running in, inspect e.authMode. (In practice, onInstall triggers always
 *     run in AuthMode.FULL, but onOpen triggers may be AuthMode.LIMITED or
 *     AuthMode.NONE.)
 */
function onInstall(e) {
  onOpen(e);
}

/**
 * Opens a sidebar in the document containing the add-on's user interface.
 * This method is only used by the regular add-on, and is never called by
 * the mobile add-on version.
 */
function showSidebar() {
  let ui = HtmlService.createHtmlOutputFromFile("sidebar").setTitle(
    "Code Generation"
  );
  SpreadsheetApp.getUi().showSidebar(ui);
}

/**
 * Gets the user-selected text and translates it from the origin language to the
 * destination language. The languages are notated by their two-letter short
 * form.
 *
 * @param {string} lang The two-letter short for the target language.
 * @return {Object} Object containing the result of the code generation.
 */
const getGeneratedCode = (lang) => {
  let ast = getAST();
  let code = generateCode(ast, lang);
  return {
    code: code,
  };
};

const getAST = () => {
  // read the spreadsheet, parse
  let ss = SpreadsheetApp.getActiveSpreadsheet();
  const cells = [];
  const sheets = ss.getSheets();
  for (const i in sheets) {
    let sheet = sheets[i];
    let sheetName = sheet.getName();
    let range = sheet.getDataRange();
    let formulas = range.getFormulas();
    for (const i of getRange(range.getNumRows())) {
      for (const j of getRange(range.getNumColumns())) {
        if (formulas[i][j]) {
          let cell = range.getCell(i + 1, j + 1);
          cells.push({
            sheet: sheetName,
            row: cell.getRow(),
            column: cell.getColumn(),
            formula: formulas[i][j],
          });
        }
      }
    }
  }
  return cells;
};

const generateCode = (ast, lang) => {
  return ast
    .map((cell) => `${cell.sheet}(${cell.row}, ${cell.column}) = ${cell.formula}`)
    .reduce((acc, curr) => acc + `${curr}\n`, "");
};

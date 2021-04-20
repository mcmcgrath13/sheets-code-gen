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
const getGeneratedCode = (lang: string, activeOnly: boolean, tabularData: boolean) => {
  Logger.log(lang);
  Logger.log(activeOnly);
  let ast = getAST(activeOnly, tabularData);
  let code = generateCode(ast, lang);
  return {
    code: code,
  };
};

const getAST = (activeOnly: boolean, tabularData: boolean) => {
  // read the spreadsheet, parse
  let ss = SpreadsheetApp.getActiveSpreadsheet();
  let ast = new Workbook(ss, activeOnly, tabularData);
  return ast;
};

const generateCode = (ast: Workbook, lang: string) => {
  return langs[lang].print(ast);
};

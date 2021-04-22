# Google Sheets Code Generation

An add-on to generate linearized code from google sheets.

## Getting Started

Project was started by adapting [this tutorial on adding a translation sidebar to Docs](https://developers.google.com/workspace/add-ons/translate-addon-sample).

Key steps:
* Set up a Google Cloud Platform project to house the code
* Install [clasp](https://github.com/google/clasp)
* Log in to the project with `clasp login`
* Push the code to the project with `clasp push`
* [Install the add-on](https://developers.google.com/workspace/add-ons/translate-addon-sample#step_3_install_the_unpublished_add-on)
* Open a google sheets doc, click the `Add-ons` menu, select `sheets-code-gen` (might take a few seconds to appear), then `start`
* The sidebar with the code generation tool should now be open
* Select a target language, whether the data contains a table in the top left of a sheet, and whether to only generate code for the current sheet
* Click `Generate`, then within a few seconds, an output should appear

## Target Programs

The `gsheets-code-gen` is designed to provide a base abstract syntax tree (AST), which can then be projected into many programs.  Presently, the following target programs are implemented.

### AST

Return the raw AST as JSON.  This is useful for inspecting the structure of the worksheet and debugging other output programs.

### GSheets

Return a variant of the google spreadsheet formula language. Ranges with the [same formula are collapsed](#range-collapsing), presenting a consolidated view of the code present in the sheet.  Otherwise, the syntax of the language remains unchanged.  The references in the formula are those for the top-left most cell in the target range (e.g. `B2:B4 = A2` implies that `B2 = A2`, `B3 = A3`, and `B4 = A4`).  The dollar sign notation is used as usual to indicate absolutely valued references instead of relative (e.g. `B2:B4 = A$2` implies that `B2 = A2`, `B3 = A2`, and `B4 = A2`).

### Julia

Return Julia code, which leverages the [SpreadsheetFunctions.jl](https://github.com/mcmcgrath13/SpreadsheetFunctions.jl) package to translate/implement the google sheets functions (only a small subset of functions currently implemented).  If the `Data is Tabular` box is checked, any [detected tables](#table-detection) will be translated to data frames and any functions which reference the range covered by the dataframe will index into it.  Otherwise, each cell which contains a formula or is in the dependency graph of a formula is translated individually and composed into arrays to represent ranges.  Variables are named based on the a1 notation of the cell in the spreadsheet (e.g. cell `B4` on `Sheet 1` becomes `sheet_1_b4`), except for tables which are named with the sheet name and `table` (e.g. the table on 'May Expenses' becomes `may_expenses_table`).

## Key Algorithms

### Range Collapsing

Collapsing similar cells in the spreadsheet into one range allows for a more consolidated data structure and shorter code representation of the sheet.  This behavior is opt-in by calling `collapseRanges` on a `Sheet` node of the AST.

The range collapsing proceeds as follows:
1. Starting from the top-left cell in the `Sheet`, compare the next cell to the right.
  * If it matches*, combine it with the present cell, delete the `Range` node for the collapsed cell
  * If it doesn't match, select it and repeat step 1
2. Complete step 1 for every row in the `Sheet`
3. Starting in the top-left `Range` of the `Sheet`, check if there is a `Range` immediately below it with the same number of columns
  * If there is a range, check if it matches* and if so merge it and remove the merged range from the `Sheet` and repeat step 3
  * If not, select the next range, and repeate step 3

\* A `Range` is considered to match another `Range` if the R1C1 notation version of the formulas match, the number formatting matches, and the names match (or are both unnamed)

### Table Detection

The present implementation of table detection is quite primitive, it assumes the following:

* Tables are columnar
* There is one row of headers
  * The first column may have a blank header, indicating that column contains row identifiers
* Headers are unique strings
* All columns have the same length
* A column contains data of the same type
* Every cell in a column contains the equivalent formula or no formula
* Tables start at the top left of the sheet
* There is only one table per sheet
* There is no missing data (empty cells)
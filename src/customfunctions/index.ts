/**
 * Custom formula
 * @customfunction
 */
function GETSHEETS(cmd) {
  return SpreadsheetApp.getActive()
    .getSheets()
    .map((sheet) => [sheet[`get${cmd}`]()]);
}

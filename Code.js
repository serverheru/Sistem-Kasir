const SPREADSHEET_ID = "19fPeJLyEm_DtzYENLQsg3CpqcQXTIKzfd2liDqB9OHU";

function doGet() {
  return HtmlService
    .createTemplateFromFile("index")
    .evaluate()
    .setTitle("KopiKasir Pro")
    .addMetaTag("viewport", "width=device-width, initial-scale=1");
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

function getSheet(sheetName) {
  const spreadsheet = SpreadsheetApp.openById(SPREADSHEET_ID);
  return spreadsheet.getSheetByName(sheetName);
}

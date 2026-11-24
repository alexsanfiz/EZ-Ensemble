function doGet() {
  return HtmlService.createHtmlOutputFromFile("Index").setTitle("EZ-Ensemble");
}

// Run once from the editor so the demo sheet belongs to the project owner.
function setupDemo_() {
  const properties = PropertiesService.getScriptProperties();
  if (!properties.getProperty("DEMO_SHEET_ID")) {
    const spreadsheet = SpreadsheetApp.create("EZ-Ensemble Plumbing Demo");
    const sheet = spreadsheet.getSheets()[0].setName("Demo");
    sheet.getRange("A1").setValue("Hello from EZ-Ensemble!");
    properties.setProperty("DEMO_SHEET_ID", spreadsheet.getId());
  }
  console.log(SpreadsheetApp.openById(properties.getProperty("DEMO_SHEET_ID")).getUrl());
}

function getDemoMessage() {
  const id = PropertiesService.getScriptProperties().getProperty("DEMO_SHEET_ID");
  if (!id) throw new Error("Run setupDemo_ in the Apps Script editor first.");
  const sheet = SpreadsheetApp.openById(id).getSheetByName("Demo");
  if (!sheet) throw new Error("The demo spreadsheet needs a tab named Demo.");
  return sheet.getRange("A1").getDisplayValue();
}

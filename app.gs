function createReportTrigger() {
  createMonthlyReport();
  var cellOptionCounter = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Options').getRange("B2");
  var valueOptionCounter = cellOptionCounter.getValue();
  cellOptionCounter.setValue(++valueOptionCounter);
  if (valueOptionCounter > 6)
    cellOptionCounter.setValue(0);
}

function createMonthlyReport() {
  initMonthlyOptions();
  initPeriodTable('#66cc66');
}

function onOpen() {
  var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.addMenu('GoldenCode Report', [
    {name: 'Создать Ежемесячный Отчёт', functionName: 'createMonthlyReport'}
  ]);
}

function createTrigger1() {
  ScriptApp.newTrigger('createReportTrigger')
    .timeBased()
    .everyDays(1)
    .atHour(1)
    .create();
}

function createTrigger2() {
  ScriptApp.newTrigger('createReportTrigger')
    .timeBased()
    .everyDays(1)
    .atHour(2)
    .create();
}

function createTrigger3() {
  ScriptApp.newTrigger('createReportTrigger')
    .timeBased()
    .everyDays(1)
    .atHour(3)
    .create();
}

function deleteAllTriggers() {
  var allTriggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < allTriggers.length; i++) {
    ScriptApp.deleteTrigger(allTriggers[i]);
  }
}

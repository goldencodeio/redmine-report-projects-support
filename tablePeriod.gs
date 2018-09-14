function initPeriodTable(color) {
  writePeriodHeader(color);
  writePeriodUserRows(color);
}

function writePeriodHeader(color) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  sheet.getRange(1, 1).setValue('Проект').setBackground(color);
  var columnI = 2;
  REPORT.forEach(function(rep) {
    sheet.getRange(1, columnI++).setValue(rep.name).setBackground(color);
  });
}

function writePeriodUserRows(color) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var rowI = 2 + (OPTIONS.counter * 20);
  var allProjects = APIRequest('projects').projects;
  OPTIONS.projects = allProjects.filter(function(project) {
    return /SUP/.test(project.name);
  });
  OPTIONS.projects = OPTIONS.projects.splice(OPTIONS.counter * 20, 20);
  if (OPTIONS.projects.length === 0) {
    Browser.msgBox('По данному счётчику отсутсвуют проекты');
    return
  }
  OPTIONS.projects.forEach(function(project) {
    sheet.getRange(rowI++, 1).setValue(project.name).setBackground(color);
  });
  processReports();
}

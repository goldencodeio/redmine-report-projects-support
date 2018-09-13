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
  var rowI = 2;
  var allProjects = APIRequest('projects').projects;
  OPTIONS.projects = allProjects.filter(function(project) {
    return /SUP/.test(project.name);
  });
  OPTIONS.projects.forEach(function(project) {
    var nameProject = project.parent ? project.parent.name : project.name;
    sheet.getRange(rowI++, 1).setValue(nameProject).setBackground(color);
  });
}

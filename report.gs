var REPORT = [
  {
    code: 'legal_person',
    name: 'Юридическое\nлицо',
    manual: true
  },
  {
    code: 'lpr_notices',
    name: 'ЛПР уведомлен/\nне уведомлен',
    manual: true
  },
  {
    code: 'monthly_pay',
    name: 'Сумма в месяц,\nруб',
    manual: true
  },
  {
    code: 'plan_quantity_departure',
    name: 'планируемое\nколичество\nвыездов',
    manual: true
  },
  {
    code: 'fact_quantity_departure',
    name: 'фактическое\nколичество\nвыездов',
    manual: true
  },
  {
    code: 'contract_time_spend',
    name: 'Прописано в\nдоговоре,\nч/мес',
    manual: true
  },
  {
    code: 'max_time_spend',
    name: 'Макс. допустимые\nтрудозатраты,\nч/мес.',
    manual: true
  },
  {
    code: 'time_spend_open_task',
    name: 'Фактическое\nкол-во часов,\nч (с незакрытыми\nзадачами)',
    manual: false
  },
  {
    code: 'time_spend_closed_task',
    name: 'Только закрытые\nзадачи, ч',
    manual: false
  }
];

function processReports() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getActiveSheet();
  var rowI = 2 + (OPTIONS.counter * OPTIONS.quantity);
  var columnI = 2;

  OPTIONS.projects.forEach(function(project) {
    REPORT.forEach(function(report) {
      if (!report.manual) {
        var reportValue = getProjectReport(report.code, project);
        if (report.code === 'time_spend_open_task' && isNumeric(sheet.getRange(rowI, columnI - 1).getValue())) {
          if (reportValue > sheet.getRange(rowI, columnI - 1).getValue())
            sheet.getRange(rowI, 1, 1, sheet.getLastColumn()).setBackground('#f00');
          else
            sheet.getRange(rowI, 1, 1, sheet.getLastColumn()).setBackground('#fff');
        }
        sheet.getRange(rowI, columnI++).setValue(reportValue);
      } else {
        columnI++;
      }
    });

    columnI = 2;
    rowI++;
  });
}

function getProjectReport(report, project) {
  switch (report) {
    case 'time_spend_open_task':
      return getTimeSpendOpenTask(project);
      break;

    case 'time_spend_closed_task':
      return getTimeSpendClosedTask(project);
      break;
  }
}

function getTimeSpendOpenTask(project) {
  var issues = APIRequest('issues', {query: [
    {key: 'project_id', value: project.id},
    {key: 'status_id', value: 'open'},
    {key: 'tracker_id', value: 7},
    {key: 'created_on', value: getDateRage(OPTIONS.startDate, OPTIONS.finalDate)}
  ]});

  var timeEntries = [];

  issues.issues.forEach(function(issue) {
    var res = APIRequest('time_entries', {query: [
      {key: 'issue_id', value: issue.id},
      {key: 'activity_id', value: 33}
    ]});

    timeEntries = timeEntries.concat(res.time_entries);
  });

  return timeEntries.reduce(function(a, c) {
    return a + c.hours;
  }, 0);
}

function getTimeSpendClosedTask(project) {
  var allTasks = [];

  for (var i = 4; i <= 5; i++) {
    var issues = APIRequest('issues', {query: [
      {key: 'project_id', value: project.id},
      {key: 'status_id', value: 5},
      {key: 'tracker_id', value: 7},
      {key: 'cf_34', value: '1'},
      {key: 'cf_7', value: i},
      {key: 'created_on', value: getDateRage(OPTIONS.startDate, OPTIONS.finalDate)}
    ]});

    allTasks = allTasks.concat(issues.issues);
  }

  var timeEntries = [];

  allTasks.forEach(function(issue) {
    var res = APIRequest('time_entries', {query: [
      {key: 'issue_id', value: issue.id},
      {key: 'activity_id', value: 33}
    ]});

    timeEntries = timeEntries.concat(res.time_entries);
  });

  return timeEntries.reduce(function(a, c) {
    return a + c.hours;
  }, 0);
}

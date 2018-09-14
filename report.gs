var REPORT = [
  {
    code: 'legal_person',
    name: 'Юридическое\nлицо',
    manual: true
  },
  {
    code: 'monthly_pay',
    name: 'Сумма в месяц,\nруб',
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
    name: 'Фактическое кол-во\nчасов в августе,\nч (с незакрытыми\nзадачами)',
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
      {key: 'issue_id', value: issue.id}
    ]});

    timeEntries = timeEntries.concat(res.time_entries);
  });

  return timeEntries.reduce(function(a, c) {
    return a + c.hours;
  }, 0);
}

function getTimeSpendClosedTask(project) {
  var issues = APIRequest('issues', {query: [
    {key: 'project_id', value: project.id},
    {key: 'status_id', value: 5},
    {key: 'tracker_id', value: 7},
    {key: 'created_on', value: getDateRage(OPTIONS.startDate, OPTIONS.finalDate)}
  ]});

  var timeEntries = [];

  issues.issues.forEach(function(issue) {
    var res = APIRequest('time_entries', {query: [
      {key: 'issue_id', value: issue.id}
    ]});

    timeEntries = timeEntries.concat(res.time_entries);
  });

  return timeEntries.reduce(function(a, c) {
    return a + c.hours;
  }, 0);
}

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
  var rowI = 2 + (OPTIONS.counter * 20);
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
      return GetTimeSpendClosedTask(project);
      break;
  }
}

function getTimeSpendOpenTask(project) {
  var res = APIRequest('time_entries', {query: [
    {key: 'project_id', value: project.id},
    {key: 'spent_on', value: getDateRage(OPTIONS.startDate, OPTIONS.finalDate)}
  ]});

  var timeEntries = res.time_entries.filter(function(timeEntry) {
    if (!timeEntry.issue) return false;
    var res = APIRequestById('issues', timeEntry.issue.id);
    return (res.issue.status.id !== 5 && res.issue.tracker.id === 7);
  });

  return timeEntries.reduce(function(a, c) {
    return a + c.hours;
  }, 0);
}

function GetTimeSpendClosedTask(project) {
  var res = APIRequest('time_entries', {query: [
    {key: 'project_id', value: project.id},
    {key: 'spent_on', value: getDateRage(OPTIONS.startDate, OPTIONS.finalDate)}
  ]});

  var timeEntries = res.time_entries.filter(function(timeEntry) {
    if (!timeEntry.issue) return false;
    var res = APIRequestById('issues', timeEntry.issue.id);
    return (res.issue.status.id === 5 && res.issue.tracker.id === 7);
  });

  return timeEntries.reduce(function(a, c) {
    return a + c.hours;
  }, 0);
}

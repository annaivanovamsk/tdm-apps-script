/***** НАСТРОЙКИ *****/

const TRAFFIC_SHEET_NAME = 'Поведение трафика';

const SOURCES = {
  ad: 'Реклама',
  organic: 'Поиск',
  direct: 'Прямые заходы'
};

/**
 * Один раз вставь сюда ID счетчика и токен, запусти функцию saveMetrikaSettings().
 * Потом можешь удалить токен из кода.
 */
function saveMetrikaSettings() {
  PropertiesService.getScriptProperties().setProperties({
    METRIKA_COUNTER_ID: '92370926',
    METRIKA_TOKEN: 'y0__wgBEPjhu80IGJSdQiC1habSFzDO9aTVCHBMM2MrbYTQtMwQWgGMABTsU7yD'
  });
}


/***** ГЛАВНЫЕ ФУНКЦИИ *****/

/**
 * Запустить один раз.
 * Создает лист, оформление, графики и еженедельный триггер.
 */
function setupTrafficBehaviorReport() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(TRAFFIC_SHEET_NAME);

  if (!sh) {
    sh = ss.insertSheet(TRAFFIC_SHEET_NAME);
  }

  sh.clear();
  sh.clearFormats();

  drawTrafficTemplate_(sh);
  createTrafficWeeklyTrigger_();

  SpreadsheetApp.getUi().alert('Лист создан. Автообновление поставлено на понедельник.');
}


/**
 * Основная функция обновления.
 * Ее можно запускать вручную и она же будет работать по понедельникам.
 */
function updateTrafficBehaviorWeekly() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(TRAFFIC_SHEET_NAME) || ss.insertSheet(TRAFFIC_SHEET_NAME);

  if (sh.getLastRow() < 5) {
    drawTrafficTemplate_(sh);
  }

  const period = getPreviousWeekPeriod_();
  const result = fetchMetrikaTraffic_(period.date1, period.date2);

  writeCurrentWeek_(sh, period, result);
  upsertHistory_(sh, period, result.rows);
  rebuildHelperTables_(sh);
  rebuildTrafficCharts_(sh);
  writeTrafficConclusion_(sh, result.rows);

  SpreadsheetApp.flush();
}


/***** ПОЛУЧЕНИЕ ДАННЫХ ИЗ МЕТРИКИ *****/

function fetchMetrikaTraffic_(date1, date2) {
  const props = PropertiesService.getScriptProperties();
  const counterId = props.getProperty('METRIKA_COUNTER_ID');
  const token = props.getProperty('METRIKA_TOKEN');

  if (!counterId || !token) {
    throw new Error('Не указан METRIKA_COUNTER_ID или METRIKA_TOKEN. Сначала запусти saveMetrikaSettings().');
  }

  const metrics = [
    'ym:s:visits',
    'ym:s:users',
    'ym:s:avgVisitDurationSeconds',
    'ym:s:bounceRate',
    'ym:s:pageDepth',
    'ym:s:newUsers'
  ].join(',');

  const params = {
    ids: counterId,
    date1: date1,
    date2: date2,
    metrics: metrics,
    dimensions: 'ym:s:lastTrafficSource',
    filters: "ym:s:lastTrafficSource=.('ad','organic','direct') AND ym:s:isRobot=='No'",
    accuracy: 'full',
    lang: 'ru',
    limit: 100
  };

  const url = 'https://api-metrika.yandex.net/stat/v1/data?' + toQueryString_(params);

  const response = UrlFetchApp.fetch(url, {
    method: 'get',
    headers: {
      Authorization: 'OAuth ' + token
    },
    muteHttpExceptions: true
  });

  const code = response.getResponseCode();
  const text = response.getContentText();

  if (code >= 300) {
    throw new Error('Ошибка Метрики API: ' + code + ' — ' + text);
  }

  const json = JSON.parse(text);

  const rowsMap = {};

  Object.keys(SOURCES).forEach(id => {
    rowsMap[id] = {
      id: id,
      source: SOURCES[id],
      visits: 0,
      users: 0,
      avgTimeSec: 0,
      bounceRate: 0,
      pageDepth: 0,
      newUsers: 0
    };
  });

  if (json.data && json.data.length) {
    json.data.forEach(item => {
      const sourceId = item.dimensions[0].id;
      const m = item.metrics;

      if (!rowsMap[sourceId]) return;

      rowsMap[sourceId] = {
        id: sourceId,
        source: SOURCES[sourceId],
        visits: Math.round(m[0] || 0),
        users: Math.round(m[1] || 0),
        avgTimeSec: m[2] || 0,
        bounceRate: m[3] || 0,
        pageDepth: m[4] || 0,
        newUsers: Math.round(m[5] || 0)
      };
    });
  }

  const rows = Object.keys(SOURCES).map(id => rowsMap[id]);

  const totals = calculateTotals_(rows);

  return {
    rows: rows,
    totals: totals
  };
}


/***** ЗАПИСЬ НА ЛИСТ *****/

function drawTrafficTemplate_(sh) {
  sh.setName(TRAFFIC_SHEET_NAME);
  sh.setFrozenRows(5);

  sh.getRange('A1').setValue('Поведение трафика');
  sh.getRange('A2').setValue('Еженедельный отчет по качеству трафика');
  sh.getRange('A4').setValue('Период:');

  sh.getRange('A1').setFontSize(22).setFontWeight('bold');
  sh.getRange('A2').setFontStyle('italic').setFontColor('#555555');
  sh.getRange('A4').setFontWeight('bold');

  sh.setColumnWidth(1, 150);
  sh.setColumnWidth(2, 130);
  sh.setColumnWidth(3, 120);
  sh.setColumnWidth(4, 130);
  sh.setColumnWidth(5, 110);
  sh.setColumnWidth(6, 120);
  sh.setColumnWidth(7, 150);

  drawKpiCards_(sh);

  sh.getRange('A7').setValue('По источникам трафика');
  sh.getRange('A7').setFontSize(14).setFontWeight('bold');

  const sourceHeaders = [
    ['Источник', 'Визиты', 'Пользователи', 'Время на сайте', 'Отказы', 'Глубина', 'Новые пользователи']
  ];
  sh.getRange('A8:G8').setValues(sourceHeaders);
  styleHeader_(sh.getRange('A8:G8'));

  sh.getRange('A14').setValue('Динамика по неделям');
  sh.getRange('A14').setFontSize(14).setFontWeight('bold');

  const historyHeaders = [
    ['Неделя', 'Источник', 'Визиты', 'Время на сайте', 'Отказы', 'Глубина']
  ];
  sh.getRange('A15:F15').setValues(historyHeaders);
  styleHeader_(sh.getRange('A15:F15'));

  sh.getRange('A33').setValue('Краткий вывод');
  sh.getRange('A33').setFontSize(14).setFontWeight('bold');

  sh.getRange('A34:G37').merge();
  sh.getRange('A34').setWrap(true).setVerticalAlignment('top');

  sh.hideColumns(25, 10);
}


function drawKpiCards_(sh) {
  const cards = [
    { range: 'I1:J4', title: 'Визиты', color: '#EAF3FF' },
    { range: 'K1:L4', title: 'Пользователи', color: '#EAF7EA' },
    { range: 'M1:N4', title: 'Ср. время на сайте', color: '#F3EEFF' },
    { range: 'O1:P4', title: 'Отказы', color: '#FFF2E6' },
    { range: 'Q1:R4', title: 'Глубина просмотра', color: '#EAF8FA' },
    { range: 'S1:T4', title: 'Новые пользователи', color: '#FFF7DE' }
  ];

  cards.forEach(card => {
    const r = sh.getRange(card.range);
    r.merge();
    r.setBackground(card.color);
    r.setBorder(true, true, true, true, false, false, '#D9D9D9', SpreadsheetApp.BorderStyle.SOLID);
    r.setVerticalAlignment('middle');
    r.setHorizontalAlignment('center');
    r.setWrap(true);
    r.setFontWeight('bold');
    r.setValue(card.title + '\n—');
  });

  for (let col = 9; col <= 20; col++) {
    sh.setColumnWidth(col, 105);
  }
}


function writeCurrentWeek_(sh, period, result) {
  sh.getRange('A4').setValue('Период: ' + period.label);

  const rows = result.rows.map(r => [
    r.source,
    r.visits,
    r.users,
    secondsToSheetTime_(r.avgTimeSec),
    r.bounceRate / 100,
    r.pageDepth,
    r.newUsers
  ]);

  const total = result.totals;

  rows.push([
    'Итого',
    total.visits,
    total.users,
    secondsToSheetTime_(total.avgTimeSec),
    total.bounceRate / 100,
    total.pageDepth,
    total.newUsers
  ]);

  sh.getRange('A9:G12').clearContent();
  sh.getRange(9, 1, rows.length, 7).setValues(rows);

  sh.getRange('A12:G12').setFontWeight('bold');
  sh.getRange('D9:D12').setNumberFormat('[m]:ss');
  sh.getRange('E9:E12').setNumberFormat('0.0%');
  sh.getRange('F9:F12').setNumberFormat('0.00');

  updateKpiCards_(sh, total);
}


function updateKpiCards_(sh, total) {
  const cards = [
    { range: 'I1:J4', title: 'Визиты', value: formatNumber_(total.visits), color: '#EAF3FF' },
    { range: 'K1:L4', title: 'Пользователи', value: formatNumber_(total.users), color: '#EAF7EA' },
    { range: 'M1:N4', title: 'Ср. время на сайте', value: secondsToText_(total.avgTimeSec), color: '#F3EEFF' },
    { range: 'O1:P4', title: 'Отказы', value: formatPercentText_(total.bounceRate), color: '#FFF2E6' },
    { range: 'Q1:R4', title: 'Глубина просмотра', value: total.pageDepth.toFixed(2).replace('.', ','), color: '#EAF8FA' },
    { range: 'S1:T4', title: 'Новые пользователи', value: formatNumber_(total.newUsers), color: '#FFF7DE' }
  ];

  cards.forEach(card => {
    const r = sh.getRange(card.range);
    r.setBackground(card.color);
    r.setValue(card.title + '\n' + card.value + '\nза неделю');
    r.setFontWeight('bold');
    r.setHorizontalAlignment('center');
    r.setVerticalAlignment('middle');
  });
}


function upsertHistory_(sh, period, rows) {
  const startRow = 16;
  const lastRow = Math.max(sh.getLastRow(), startRow - 1);

  let existing = [];
  if (lastRow >= startRow) {
    existing = sh.getRange(startRow, 1, lastRow - startRow + 1, 8).getValues();
  }

  const keyToRow = {};
  existing.forEach((row, i) => {
    const week = row[0];
    const source = row[1];
    if (week && source) {
      keyToRow[week + '|' + source] = startRow + i;
    }
  });

  rows.forEach(r => {
    const values = [
      period.label,
      r.source,
      r.visits,
      secondsToSheetTime_(r.avgTimeSec),
      r.bounceRate / 100,
      r.pageDepth,
      '',
      period.date1
    ];

    const key = period.label + '|' + r.source;

    if (keyToRow[key]) {
      sh.getRange(keyToRow[key], 1, 1, 8).setValues([values]);
    } else {
      sh.appendRow(values);
    }
  });

  const newLastRow = sh.getLastRow();
  if (newLastRow >= startRow) {
    sh.getRange(startRow, 4, newLastRow - startRow + 1, 1).setNumberFormat('[m]:ss');
    sh.getRange(startRow, 5, newLastRow - startRow + 1, 1).setNumberFormat('0.0%');
    sh.getRange(startRow, 6, newLastRow - startRow + 1, 1).setNumberFormat('0.00');
  }

  sh.hideColumns(8);
}


/***** ВСПОМОГАТЕЛЬНЫЕ ТАБЛИЦЫ ДЛЯ ГРАФИКОВ *****/

function rebuildHelperTables_(sh) {
  const startRow = 16;
  const lastRow = sh.getLastRow();

  if (lastRow < startRow) return;

  const data = sh.getRange(startRow, 1, lastRow - startRow + 1, 8).getValues()
    .filter(r => r[0] && r[1]);

  const weeksMap = {};

  data.forEach(r => {
    const week = r[0];
    const source = r[1];
    const visits = Number(r[2]) || 0;
    const timeMinutes = Number(r[3]) * 24 * 60;
    const bounce = Number(r[4]) || 0;
    const depth = Number(r[5]) || 0;
    const dateKey = r[7];

    if (!weeksMap[week]) {
      weeksMap[week] = {
        week: week,
        dateKey: dateKey,
        sources: {}
      };
    }

    weeksMap[week].sources[source] = {
      visits: visits,
      timeMinutes: timeMinutes,
      bounce: bounce,
      depth: depth
    };
  });

  const weeks = Object.values(weeksMap).sort((a, b) => {
    return String(a.dateKey).localeCompare(String(b.dateKey));
  });

  const helperCol = 25; // Y
  sh.getRange(1, helperCol, 80, 12).clearContent();

  const sourceNames = ['Реклама', 'Поиск', 'Прямые заходы'];

  writeHelperTable_(sh, 1, helperCol, 'time', weeks, sourceNames);
  writeHelperTable_(sh, 12, helperCol, 'bounce', weeks, sourceNames);
  writeHelperTable_(sh, 23, helperCol, 'depth', weeks, sourceNames);
  writeComparisonHelper_(sh, 34, helperCol);
}


function writeHelperTable_(sh, row, col, metric, weeks, sourceNames) {
  sh.getRange(row, col).setValue(metric);
  sh.getRange(row + 1, col, 1, 4).setValues([['Неделя'].concat(sourceNames)]);

  const values = weeks.map(w => {
    return [
      w.week,
      getMetricValue_(w.sources, sourceNames[0], metric),
      getMetricValue_(w.sources, sourceNames[1], metric),
      getMetricValue_(w.sources, sourceNames[2], metric)
    ];
  });

  if (values.length) {
    sh.getRange(row + 2, col, values.length, 4).setValues(values);
  }
}


function writeComparisonHelper_(sh, row, col) {
  const sourceData = sh.getRange('A9:G11').getValues();

  sh.getRange(row, col, 1, 4).setValues([['Источник', 'Визиты', 'Время на сайте, мин.', 'Отказы']]);

  const values = sourceData.map(r => [
    r[0],
    r[1],
    Number(r[3]) * 24 * 60,
    Number(r[4])
  ]);

  sh.getRange(row + 1, col, values.length, 4).setValues(values);
}


function getMetricValue_(sources, sourceName, metric) {
  if (!sources[sourceName]) return 0;

  if (metric === 'time') return sources[sourceName].timeMinutes || 0;
  if (metric === 'bounce') return sources[sourceName].bounce || 0;
  if (metric === 'depth') return sources[sourceName].depth || 0;

  return 0;
}


/***** ГРАФИКИ *****/

function rebuildTrafficCharts_(sh) {
  sh.getCharts().forEach(chart => sh.removeChart(chart));

  const helperCol = 25;

  const timeRange = getHelperRange_(sh, 2, helperCol);
  const bounceRange = getHelperRange_(sh, 13, helperCol);
  const depthRange = getHelperRange_(sh, 24, helperCol);
  const comparisonRange = sh.getRange(34, helperCol, 4, 4);

  if (timeRange) {
    insertLineChart_(sh, timeRange, 'Среднее время на сайте по неделям', 7, 9);
  }

  if (bounceRange) {
    insertLineChart_(sh, bounceRange, 'Отказы по неделям', 7, 15);
  }

  if (depthRange) {
    insertLineChart_(sh, depthRange, 'Глубина просмотра', 22, 9);
  }

  insertColumnChart_(sh, comparisonRange, 'Сравнение источников', 22, 15);
}


function getHelperRange_(sh, headerRow, col) {
  let row = headerRow + 1;

  while (sh.getRange(row, col).getValue()) {
    row++;
  }

  const numRows = row - headerRow;

  if (numRows < 2) return null;

  return sh.getRange(headerRow, col, numRows, 4);
}


function insertLineChart_(sh, range, title, row, col) {
  const chart = sh.newChart()
    .setChartType(Charts.ChartType.LINE)
    .addRange(range)
    .setPosition(row, col, 0, 0)
    .setOption('title', title)
    .setOption('legend', { position: 'bottom' })
    .setOption('width', 520)
    .setOption('height', 260)
    .setOption('curveType', 'function')
    .build();

  sh.insertChart(chart);
}


function insertColumnChart_(sh, range, title, row, col) {
  const chart = sh.newChart()
    .setChartType(Charts.ChartType.COLUMN)
    .addRange(range)
    .setPosition(row, col, 0, 0)
    .setOption('title', title)
    .setOption('legend', { position: 'bottom' })
    .setOption('width', 520)
    .setOption('height', 260)
    .build();

  sh.insertChart(chart);
}


/***** АВТО-ВЫВОД *****/

function writeTrafficConclusion_(sh, rows) {
  const search = rows.find(r => r.id === 'organic');
  const ad = rows.find(r => r.id === 'ad');
  const direct = rows.find(r => r.id === 'direct');

  let text = '';

  if (search && ad) {
    if (search.avgTimeSec > ad.avgTimeSec && search.bounceRate < ad.bounceRate) {
      text += '• Поисковый трафик показывает более высокую вовлеченность: выше время на сайте и ниже отказы.\n';
    } else {
      text += '• Поисковый и рекламный трафик нужно смотреть в динамике по времени на сайте, отказам и глубине просмотра.\n';
    }
  }

  if (ad) {
    text += '• По рекламному трафику отслеживаем динамику по неделям: время на сайте, отказы и глубину просмотра.\n';
  }

  if (direct && direct.avgTimeSec < 20) {
    text += '• По прямым заходам короткое время на сайте — стоит смотреть качество этих визитов отдельно.\n';
  }

  text += '• Конверсии остаются в основном отчете, здесь фиксируем именно качество посещений сайта.';

  sh.getRange('A34').setValue(text);
}


/***** ТРИГГЕР *****/

function createTrafficWeeklyTrigger_() {
  const functionName = 'updateTrafficBehaviorWeekly';

  ScriptApp.getProjectTriggers().forEach(trigger => {
    if (trigger.getHandlerFunction() === functionName) {
      ScriptApp.deleteTrigger(trigger);
    }
  });

  ScriptApp.newTrigger(functionName)
    .timeBased()
    .onWeekDay(ScriptApp.WeekDay.MONDAY)
    .atHour(9)
    .create();
}


/***** УТИЛИТЫ *****/

function getPreviousWeekPeriod_() {
  const tz = Session.getScriptTimeZone() || 'Europe/Moscow';
  const today = new Date();

  const day = today.getDay(); // 0 воскресенье, 1 понедельник
  const daysFromMonday = (day + 6) % 7;

  const thisMonday = new Date(today);
  thisMonday.setDate(today.getDate() - daysFromMonday);
  thisMonday.setHours(0, 0, 0, 0);

  const lastMonday = new Date(thisMonday);
  lastMonday.setDate(thisMonday.getDate() - 7);

  const lastSunday = new Date(thisMonday);
  lastSunday.setDate(thisMonday.getDate() - 1);

  const date1 = Utilities.formatDate(lastMonday, tz, 'yyyy-MM-dd');
  const date2 = Utilities.formatDate(lastSunday, tz, 'yyyy-MM-dd');

  const label = Utilities.formatDate(lastMonday, tz, 'dd.MM') + '–' + Utilities.formatDate(lastSunday, tz, 'dd.MM.yyyy');

  return { date1, date2, label };
}


function calculateTotals_(rows) {
  const visits = rows.reduce((sum, r) => sum + r.visits, 0);
  const users = rows.reduce((sum, r) => sum + r.users, 0);
  const newUsers = rows.reduce((sum, r) => sum + r.newUsers, 0);

  const weighted = (field) => {
    if (!visits) return 0;
    return rows.reduce((sum, r) => sum + ((r[field] || 0) * r.visits), 0) / visits;
  };

  return {
    visits: visits,
    users: users,
    avgTimeSec: weighted('avgTimeSec'),
    bounceRate: weighted('bounceRate'),
    pageDepth: weighted('pageDepth'),
    newUsers: newUsers
  };
}


function toQueryString_(params) {
  return Object.keys(params)
    .map(key => encodeURIComponent(key) + '=' + encodeURIComponent(params[key]))
    .join('&');
}


function secondsToSheetTime_(seconds) {
  return Number(seconds || 0) / 86400;
}


function secondsToText_(seconds) {
  seconds = Math.round(seconds || 0);
  const min = Math.floor(seconds / 60);
  const sec = seconds % 60;
  return String(min).padStart(2, '0') + ':' + String(sec).padStart(2, '0');
}


function formatNumber_(num) {
  return String(Math.round(num || 0)).replace(/\B(?=(\d{3})+(?!\d))/g, ' ');
}


function formatPercentText_(value) {
  return (Number(value || 0)).toFixed(1).replace('.', ',') + '%';
}


function styleHeader_(range) {
  range
    .setBackground('#EAF3FF')
    .setFontWeight('bold')
    .setHorizontalAlignment('center')
    .setBorder(true, true, true, true, true, true);
}
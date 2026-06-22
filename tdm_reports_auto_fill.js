/**
 * TDM reports for Google Sheets + Yandex Metrica.
 *
 * Важно:
 * - токен Метрики хранится только в Script Properties: METRIKA_TOKEN;
 * - счетчик Метрики берется из строки "Счетчик Метрики" / "Счётчик Метрики" в листе;
 * - все запросы к API Метрики идут с attribution=automatic.
 */

const TDM_REPORTS_CONFIG = {
  spreadsheetId: '1zXDDfiRYHJE34iOJg-QfAC0bvoTKMIAuVMfsJS0z5Dg',
  metrikaApiUrl: 'https://api-metrika.yandex.net/stat/v1/data',
  metrikaGoalsUrl: 'https://api-metrika.yandex.net/management/v1/counter/%s/goals',
  tokenPropertyName: 'METRIKA_TOKEN',
  defaultMetrikaCounterId: '92370926',
  attribution: 'automatic',
  lang: 'ru',
  accuracy: 'full',
  timezone: '+03:00',
  sheets: {
    regionsCities: 'Регионы_города',
    regionsSettings: 'Настройки_Регионы',
    regionsSource: 'Регионы для продвижения',
    products: 'Товары_по_неделям',
    promotedPositions: 'Позиции для продвижения.'
  },
  electroEquipmentPages: [
    {
      id: 'electro-001',
      group: 'Новое товарное направление — электрооборудование',
      name: 'Электрооборудование',
      url: 'https://tdmtdm.ru/catalog/elektrooborudovanie/'
    },
    {
      id: 'electro-002',
      group: 'Новое товарное направление — электрооборудование',
      name: 'Вилки',
      url: 'https://tdmtdm.ru/catalog/elektrooborudovanie/vilki/'
    },
    {
      id: 'electro-003',
      group: 'Новое товарное направление — электрооборудование',
      name: 'Выключатели',
      url: 'https://tdmtdm.ru/catalog/elektrooborudovanie/vyklyuchateli/'
    },
    {
      id: 'electro-004',
      group: 'Новое товарное направление — электрооборудование',
      name: 'Рамки монтажные',
      url: 'https://tdmtdm.ru/catalog/elektrooborudovanie/ramki_montazhnye/'
    },
    {
      id: 'electro-005',
      group: 'Новое товарное направление — электрооборудование',
      name: 'Розетки',
      url: 'https://tdmtdm.ru/catalog/elektrooborudovanie/rozetki/'
    }
  ],
  markers: {
    inputBlock: 'Блок ввода',
    outputBlock: 'Блок вывода',
    counter: 'Счетчик Метрики'
  },
  goalNameParts: {
    threePages: ['просмотр 3', 'просмотр трех', '3 страницы', '3х страниц'],
    phoneClick: ['телефон', 'номер телефона', 'клик по номеру'],
    emailClick: ['email', 'e-mail', 'почт'],
    talkMe: ['talkme', 'чат', 'клиент написал'],
    ecommercePurchase: ['ecommerce', 'покуп', 'заказ'],
    ecommerceCart: ['ecommerce', 'корзин'],
    roistatLead: ['roistat', 'заяв'],
    roistatCall: ['roistat', 'звон']
  },
  regionsGoalKeys: ['threePages', 'emailClick', 'phoneClick', 'talkMe', 'ecommerceCart'],
  goalLabels: {
    threePages: 'Просмотр 3х страниц',
    emailClick: 'Клик: По email адресу',
    phoneClick: 'Клик: По номеру телефона',
    talkMe: 'TalkMe: Клиент написал в чат (онлайн)',
    ecommerceCart: 'Ecommerce: добавление в корзину'
  }
};

function onOpen() {
  try {
    addTdmWeeklyPlanMenu_();
  } catch (e) {
    Logger.log(e);
  }

  try {
    addEnoEmoMenu_();
  } catch (e) {
    Logger.log(e);
  }

  addTdmReportsMenu_();
}

function addTdmReportsMenu_() {
  SpreadsheetApp.getUi()
    .createMenu('ТДМ отчёты')
    .addItem('Обновить регионы/города', 'updateRegionsCitiesReport')
    .addItem('Обновить товары', 'updateProductsReport')
    .addItem('Обновить всё', 'updateAllReports')
    .addToUi();
}

function updateAllReports() {
  updateRegionsCitiesReport();
  updateProductsReport();
}

function updateRegionsCitiesReport() {
  const ss = SpreadsheetApp.openById(TDM_REPORTS_CONFIG.spreadsheetId);
  const sheet = getOrCreateRegionsCitiesSheet_(ss);
  prepareRegionsCitiesSheet_(sheet, ss);

  const counterId = getCounterIdFromSheet_(sheet);
  const period = {
    date1: '2026-04-01',
    date2: getCurrentAvailableDate_()
  };
  const rawGoals = getMetrikaGoals(counterId);
  const goals = mapGoalsByName(rawGoals);
  const goalReport = getGoalMappingReport_(goals);
  const regionsGoalReport = getSelectedRegionsGoalReport_(goals);
  const inputItems = getRegionsCitiesInput_(ss, sheet);
  const errors = [];

  Logger.log('ID счётчика: ' + counterId);
  Logger.log('Период: ' + period.date1 + ' - ' + period.date2);
  Logger.log('Модель атрибуции: ' + TDM_REPORTS_CONFIG.attribution);
  Logger.log('Найденные цели: ' + (goalReport.found.length ? goalReport.found.join(', ') : 'нет'));
  Logger.log('Цели, которые не удалось найти: ' + (goalReport.missing.length ? goalReport.missing.join(', ') : 'нет'));
  Logger.log('Цели для листа Регионы_города: ' + (regionsGoalReport.found.length ? regionsGoalReport.found.join(', ') : 'нет'));
  Logger.log('Цели Регионы_города, которые не удалось найти: ' + (regionsGoalReport.missing.length ? regionsGoalReport.missing.join(', ') : 'нет'));

  const rows = [];
  inputItems.forEach(function(item) {
    try {
      const filters = buildRegionCityFilter_(item);
      const data = fetchMetrikaData({
        counterId: counterId,
        date1: period.date1,
        date2: period.date2,
        dimensions: ['ym:s:datePeriod<group>', 'ym:s:regionAreaName', 'ym:s:regionCityName'],
        metrics: buildRegionsCitiesMetrics_(goals),
        params: { group: 'month' },
        filters: filters
      });

      (data.data || []).forEach(function(apiRow) {
        const m = apiRow.metrics || [];
        const month = dimensionName_(apiRow, 0);
        const region = dimensionName_(apiRow, 1) || item.region;
        const city = dimensionName_(apiRow, 2) || item.city || 'Не определён';
        const visits = num_(m[0]);
        const users = num_(m[1]);
        const leads = num_(m[2]);
        const bounces = num_(m[3]);
        const depth = num_(m[4]);
        const duration = num_(m[5]);
        const cr = visits ? leads / visits : 0;
        const quality = buildTrafficQuality_({
          visits: visits,
          leads: leads,
          cr: cr,
          bounces: bounces,
          depth: depth,
          duration: duration
        });
        const conclusion = buildRegionsCitiesConclusion_({
          visits: visits,
          leads: leads,
          cr: cr,
          bounces: bounces,
          depth: depth,
          duration: duration
        });

        rows.push([
          month,
          region,
          city,
          visits,
          users,
          '',
          '',
          '',
          '',
          '',
          leads,
          cr,
          '',
          bounces,
          depth,
          secondsToTime_(duration),
          quality,
          conclusion.text,
          conclusion.recommendation
        ]);
      });
    } catch (e) {
      const label = item.region + (item.city ? ' / ' + item.city : '');
      errors.push(label + ': ' + e.message);
      Logger.log('Ошибка по региону: ' + label + ': ' + e.message);
    }
  });

  writeRowsToSheet(sheet, rows, getRegionsCitiesOutputHeaders_());
  writeRegionsCitiesSummary_(sheet, rows);
  writeRegionsCitiesInsights_(sheet, rows, errors);
  buildRegionsCitiesCharts_(sheet);
  sheet.getRange('B5').setValue(new Date());
  sheet.getRange('D5').setValue(errors.length ? 'Обновлено с ошибками' : 'Обновлено');
  sheet.getRange('B6').setValue('Цели/заявки: ' + (regionsGoalReport.found.length ? regionsGoalReport.found.join(', ') : 'нет'));
  sheet.getRange('D6').setValue(regionsGoalReport.missing.length ? 'Не найдены: ' + regionsGoalReport.missing.join(', ') : 'Все цели найдены');

  Logger.log('Количество обработанных регионов/городов: ' + inputItems.length);
  Logger.log('Ошибки по регионам: ' + (errors.length ? errors.join(' | ') : 'нет'));
}

function getOrCreateRegionsCitiesSheet_(ss) {
  let sheet = ss.getSheetByName(TDM_REPORTS_CONFIG.sheets.regionsCities);
  if (!sheet) {
    sheet = ss.insertSheet(TDM_REPORTS_CONFIG.sheets.regionsCities);
  }
  return sheet;
}

function prepareRegionsCitiesSheet_(sheet, ss) {
  const existingCounterId = PropertiesService.getScriptProperties().getProperty('METRIKA_COUNTER_ID')
    || TDM_REPORTS_CONFIG.defaultMetrikaCounterId;
  const settingsSheet = getOrCreateRegionsSettingsSheet_(ss);
  const existingSettingsRows = readExistingRegionsCitiesInput_(settingsSheet);
  const existingInputRows = existingSettingsRows.length ? existingSettingsRows : readExistingRegionsCitiesInput_(sheet);
  const sourceRows = getRegionsCitiesInputFromSource_(ss);
  const inputRows = existingInputRows.length ? existingInputRows : sourceRows.map(function(item) {
    return [item.region, item.city || '', 'Да'];
  });
  prepareRegionsCitiesSettingsSheet_(settingsSheet, inputRows);
  const layout = getRegionsCitiesLayout_();

  sheet.getCharts().forEach(function(chart) {
    sheet.removeChart(chart);
  });
  sheet.clear();
  if (sheet.getMaxColumns() < 38) {
    sheet.insertColumnsAfter(sheet.getMaxColumns(), 38 - sheet.getMaxColumns());
  }

  const headers = getRegionsCitiesOutputHeaders_();
  sheet.getRange('A1').setValue('Регионы и города — трафик, заявки и качество');
  sheet.getRange('A1:T1')
    .setFontWeight('bold')
    .setFontSize(14)
    .setBackground('#1f4e79')
    .setFontColor('#ffffff');
  sheet.getRange('A3:F8').setValues([
    ['Настройки отчёта', '', '', '', '', ''],
    ['Период анализа', '01.04.2026 - текущая доступная дата', 'Счётчик Метрики', existingCounterId, 'Атрибуция', 'automatic'],
    ['Последнее обновление', new Date(), 'Статус обновления', 'Готов к обновлению', 'Директ API', 'Подключен'],
    ['Цели отчёта', 'Будут заполнены после обновления', 'Статус целей', 'Будет заполнен после обновления', '', ''],
    ['Список регионов', 'Вынесен на лист Настройки_Регионы', '', '', '', ''],
    ['', '', '', '', '', '']
  ]);

  sheet.getRange(layout.mainSummaryTitleRow, 1, 1, 12).setValues([['Главная сводка', '', '', '', '', '', '', '', '', '', '', '']]);
  sheet.getRange(layout.mainSummaryStartRow, 1, 5, 4).setValues([
    ['Всего визитов', '', 'Всего целей/заявок', ''],
    ['Общий CR', '', 'Средний отказ', ''],
    ['Средняя глубина', '', 'Среднее время на сайте', ''],
    ['Регион с максимумом заявок', '', 'Регион с максимумом визитов', ''],
    ['Регионов с трафиком без заявок', '', 'Регионов с малым объёмом данных', '']
  ]);

  sheet.getRange(layout.chartsTitleRow, 1, 1, 12).setValues([['Графики', '', '', '', '', '', '', '', '', '', '', '']]);
  sheet.getRange(layout.attentionTitleRow, 1, 1, 12).setValues([['Зона внимания', '', '', '', '', '', '', '', '', '', '', '']]);
  sheet.getRange(layout.attentionHeaderRow, 1, 1, 4).setValues([['Приоритет', 'Регион', 'Причина', 'Действие']]);
  sheet.getRange(layout.insightsTitleRow, 1, 1, 12).setValues([['Итоговые выводы', '', '', '', '', '', '', '', '', '', '', '']]);
  sheet.getRange(layout.insightsHeaderRow, 1, 1, 4).setValues([['Группа', 'Регион', 'Причина', 'Рекомендация']]);
  sheet.getRange(layout.outputTitleRow, 1, 1, 20).setValues([['Блок вывода / детальная таблица: месяц × регион × город', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '']]);
  sheet.getRange(layout.outputHeaderRow, 1, 1, headers.length).setValues([headers]);

  sheet.getRange('A3:F3').setFontWeight('bold').setBackground('#e8f0fe');
  sheet.getRange('A4:F8').setBackground('#f7faff');
  sheet.getRange(layout.mainSummaryTitleRow, 1, 1, 12).setFontWeight('bold').setBackground('#e6f4ea');
  sheet.getRange(layout.mainSummaryStartRow, 1, 5, 4).setFontWeight('bold');
  sheet.getRange(layout.chartsTitleRow, 1, 1, 12).setFontWeight('bold').setBackground('#f3f6fb');
  sheet.getRange(layout.attentionTitleRow, 1, 2, 4).setFontWeight('bold').setBackground('#fce8e6');
  sheet.getRange(layout.insightsTitleRow, 1, 2, 4).setFontWeight('bold').setBackground('#fff2cc');
  sheet.getRange(layout.outputTitleRow, 1, 2, 20).setFontWeight('bold').setBackground('#d9ead3');
  sheet.getRange('AA1:AI1').setValues([['Сводка по регионам', '', '', '', '', '', '', '', '']]);
  sheet.getRange('AA2:AI2').setValues([['Регион', 'Визиты', 'Цели/заявки', 'CR', 'Отказы', 'Глубина', 'Время', 'Вывод', 'Рекомендация']]);
  sheet.setFrozenRows(2);
  sheet.setFrozenColumns(2);
  sheet.getRange('A:AI').setWrap(true).setVerticalAlignment('middle');
  sheet.setColumnWidths(1, 1, 170);
  sheet.setColumnWidths(2, 8, 110);
  sheet.setColumnWidths(10, 10, 125);
  sheet.setColumnWidths(11, 4, 180);
  sheet.hideColumns(27, 12);
}

function getOrCreateRegionsSettingsSheet_(ss) {
  let sheet = ss.getSheetByName(TDM_REPORTS_CONFIG.sheets.regionsSettings);
  if (!sheet) {
    sheet = ss.insertSheet(TDM_REPORTS_CONFIG.sheets.regionsSettings);
  }
  return sheet;
}

function prepareRegionsCitiesSettingsSheet_(sheet, inputRows) {
  sheet.clear();
  sheet.getRange('A1:C1').setValues([['Блок ввода / входных данных', '', '']]);
  sheet.getRange('A2:C2').setValues([['Регион', 'Город', 'Включить в анализ']]);
  if (inputRows.length) {
    sheet.getRange(3, 1, inputRows.length, 3).setValues(inputRows);
  }
  sheet.getRange('A1:C2').setFontWeight('bold').setBackground('#e8f0fe');
  sheet.setFrozenRows(2);
  sheet.setColumnWidths(1, 2, 180);
  sheet.setColumnWidth(3, 140);
  sheet.getRange('A:C').setWrap(true);
}

function readExistingCounterIdFromSheet_(sheet) {
  try {
    const finder = sheet.createTextFinder('Счётчик Метрики').matchCase(false).findNext()
      || sheet.createTextFinder('Счетчик Метрики').matchCase(false).findNext();
    if (!finder) return '';

    const row = finder.getRow();
    const col = finder.getColumn();
    const values = sheet.getRange(row, col, 1, Math.max(1, sheet.getLastColumn() - col + 1)).getValues()[0];
    for (let i = 1; i < values.length; i++) {
      const value = String(values[i] || '').trim();
      if (/^\d+$/.test(value)) return value;
    }
  } catch (e) {
    return '';
  }
  return '';
}

function getRegionsCitiesLayout_() {
  return {
    mainSummaryTitleRow: 10,
    mainSummaryStartRow: 11,
    chartsTitleRow: 17,
    attentionTitleRow: 50,
    attentionHeaderRow: 51,
    attentionStartRow: 52,
    insightsTitleRow: 65,
    insightsHeaderRow: 66,
    insightsStartRow: 67,
    outputTitleRow: 100,
    outputHeaderRow: 101,
    helperSummaryHeaderRow: 2,
    helperSummaryStartRow: 3,
    helperNoLeadsHeaderRow: 2,
    helperNoLeadsStartRow: 3,
    helperMonthHeaderRow: 35,
    helperMonthStartRow: 36
  };
}

function readExistingRegionsCitiesInput_(sheet) {
  try {
    return getEnabledRows_(sheet, {
      blockType: 'input',
      requiredHeaders: ['Регион', 'Город', 'Включить в анализ']
    }).map(function(item) {
      return [
        cleanRegionCityName_(getByHeader_(item, ['Регион'])),
        cleanRegionCityName_(getByHeader_(item, ['Город'])),
        getByHeader_(item, ['Включить в анализ']) || 'Да'
      ];
    }).filter(function(row) {
      return row[0] || row[1];
    });
  } catch (e) {
    return [];
  }
}

function getRegionsCitiesOutputHeaders_() {
  return [
    'Месяц',
    'Регион',
    'Город',
    'Визиты',
    'Пользователи',
    'Показы',
    'Клики',
    'CTR',
    'Расход',
    'CPC',
    'Цели/заявки',
    'CR',
    'CPA',
    'Отказы',
    'Глубина',
    'Время на сайте',
    'Качество трафика',
    'Вывод',
    'Рекомендация'
  ];
}

function getRegionsCitiesInput_(ss, sheet) {
  const settingsSheet = ss.getSheetByName(TDM_REPORTS_CONFIG.sheets.regionsSettings);
  const inputSheet = settingsSheet || sheet;
  const manual = getEnabledRows_(inputSheet, {
    blockType: 'input',
    requiredHeaders: ['Регион', 'Город', 'Включить в анализ']
  }).map(function(item) {
    return {
      region: cleanRegionCityName_(getByHeader_(item, ['Регион'])),
      city: cleanRegionCityName_(getByHeader_(item, ['Город']))
    };
  }).filter(function(item) {
    return item.region || item.city;
  });

  if (manual.length) return uniqueRegionCityItems_(manual);

  return uniqueRegionCityItems_(getRegionsCitiesInputFromSource_(ss));
}

function getRegionsCitiesInputFromSource_(ss) {
  ss = ss || SpreadsheetApp.openById(TDM_REPORTS_CONFIG.spreadsheetId);
  const sheet = ss.getSheetByName(TDM_REPORTS_CONFIG.sheets.regionsSource);
  if (!sheet) return [];

  const values = sheet.getDataRange().getValues();
  if (!values.length) return [];

  const headers = values[0].map(normalizeHeader_);
  const regionColumns = [];
  headers.forEach(function(header, index) {
    if (header === 'наши регионы' || header === 'регионы клиента') {
      regionColumns.push(index);
    }
  });

  const skipValues = {
    'рк 1 - оплата за конверсии?????': true,
    'рк 2': true,
    'рк 3': true,
    'смена регионов 23.04': true,
    'есть заявки': true
  };

  const result = [];
  for (let r = 1; r < values.length; r++) {
    regionColumns.forEach(function(col) {
      const name = cleanRegionCityName_(values[r][col]);
      const key = normalize_(name);
      if (!name || skipValues[key]) return;
      result.push({ region: name, city: '' });
    });
  }
  return uniqueRegionCityItems_(result);
}

function uniqueRegionCityItems_(items) {
  const seen = {};
  return items.filter(function(item) {
    const region = cleanRegionCityName_(item.region);
    const city = cleanRegionCityName_(item.city);
    const key = normalize_(region + '|' + city);
    if (!key || key === '|' || seen[key]) return false;
    seen[key] = true;
    item.region = region;
    item.city = city;
    return true;
  });
}

function buildRegionCityFilter_(item) {
  const parts = [];
  if (item.region && item.city) {
    parts.push("ym:s:regionAreaName==" + quoteMetrikaValue_(item.region));
    parts.push("ym:s:regionCityName==" + quoteMetrikaValue_(item.city));
    return parts.join(' AND ');
  }

  if (item.region) {
    return [
      "ym:s:regionAreaName==" + quoteMetrikaValue_(item.region),
      "ym:s:regionCityName==" + quoteMetrikaValue_(item.region)
    ].join(' OR ');
  }

  if (item.city) {
    return [
      "ym:s:regionAreaName==" + quoteMetrikaValue_(item.city),
      "ym:s:regionCityName==" + quoteMetrikaValue_(item.city)
    ].join(' OR ');
  }

  return '';
}

function buildRegionsCitiesMetrics_(goals) {
  return [
    'ym:s:visits',
    'ym:s:users',
    sumGoalMetrics_(getSelectedRegionsGoalIds_(goals)),
    'ym:s:bounceRate',
    'ym:s:pageDepth',
    'ym:s:avgVisitDurationSeconds'
  ];
}

function writeRegionsCitiesSummary_(sheet, rows) {
  const summary = buildRegionsCitiesRegionSummary_(rows);
  const totals = buildRegionsCitiesTotals_(summary);
  const layout = getRegionsCitiesLayout_();

  sheet.getRange(layout.mainSummaryStartRow, 2, 5, 1).setValues([
    [totals.visits],
    [totals.cr],
    [totals.depth],
    [totals.topLeadsRegion],
    [totals.noLeadsCount]
  ]);
  sheet.getRange(layout.mainSummaryStartRow, 4, 5, 1).setValues([
    [totals.leads],
    [totals.bounces],
    [secondsToTime_(totals.duration)],
    [totals.topVisitsRegion],
    [totals.lowDataCount]
  ]);
  sheet.getRange(layout.mainSummaryStartRow, 2).setNumberFormat('0');
  sheet.getRange(layout.mainSummaryStartRow, 4).setNumberFormat('0');
  sheet.getRange(layout.mainSummaryStartRow + 1, 2).setNumberFormat('0.00%');
  sheet.getRange(layout.mainSummaryStartRow + 1, 4).setNumberFormat('0.00');
  sheet.getRange(layout.mainSummaryStartRow + 2, 2).setNumberFormat('0.00');
  sheet.getRange(layout.mainSummaryStartRow + 4, 2).setNumberFormat('0');
  sheet.getRange(layout.mainSummaryStartRow + 4, 4).setNumberFormat('0');

  const attentionRows = buildRegionsCitiesAttentionRows_(summary);
  sheet.getRange(layout.attentionStartRow, 1, 10, 4).clearContent();
  if (attentionRows.length) {
    sheet.getRange(layout.attentionStartRow, 1, Math.min(attentionRows.length, 10), 4)
      .setValues(attentionRows.slice(0, 10));
  }

  sheet.getRange('AA3:AI32').clearContent();
  if (summary.length) {
    sheet.getRange(3, 27, Math.min(summary.length, 30), 9)
      .setValues(summary.slice(0, 30));
  }
  sheet.getRange(3, 30, 30, 1).setNumberFormat('0.00%');

  const noLeadsRows = summary.filter(function(row) {
    return num_(row[1]) >= 50 && num_(row[2]) === 0;
  }).slice(0, 20).map(function(row) {
    return [row[0], row[1]];
  });
  sheet.getRange('AJ2:AK25').clearContent();
  sheet.getRange('AJ2:AK2').setValues([['Регион без заявок', 'Визиты']]);
  if (noLeadsRows.length) {
    sheet.getRange(3, 36, noLeadsRows.length, 2).setValues(noLeadsRows);
  }

  const monthRows = buildRegionsCitiesMonthSummary_(rows);
  sheet.getRange('AJ35:AL60').clearContent();
  sheet.getRange('AJ35:AL35').setValues([['Месяц', 'Визиты', 'Цели/заявки']]);
  if (monthRows.length) {
    sheet.getRange(36, 36, monthRows.length, 3).setValues(monthRows);
  }
}

function buildRegionsCitiesTotals_(summary) {
  const totals = {
    visits: 0,
    leads: 0,
    bouncesWeighted: 0,
    depthWeighted: 0,
    durationWeighted: 0,
    topLeadsRegion: 'нет данных',
    topVisitsRegion: 'нет данных',
    noLeadsCount: 0,
    lowDataCount: 0
  };

  summary.forEach(function(row) {
    const visits = num_(row[1]);
    const leads = num_(row[2]);
    totals.visits += visits;
    totals.leads += leads;
    totals.bouncesWeighted += num_(row[4]) * visits;
    totals.depthWeighted += num_(row[5]) * visits;
    totals.durationWeighted += timeToSeconds_(row[6]) * visits;
    if (visits >= 50 && leads === 0) totals.noLeadsCount += 1;
    if (visits > 0 && visits < 30) totals.lowDataCount += 1;
  });

  const byLeads = summary.slice().sort(function(a, b) { return num_(b[2]) - num_(a[2]); })[0];
  const byVisits = summary.slice().sort(function(a, b) { return num_(b[1]) - num_(a[1]); })[0];
  totals.cr = totals.visits ? totals.leads / totals.visits : 0;
  totals.bounces = totals.visits ? totals.bouncesWeighted / totals.visits : 0;
  totals.depth = totals.visits ? totals.depthWeighted / totals.visits : 0;
  totals.duration = totals.visits ? totals.durationWeighted / totals.visits : 0;
  if (byLeads) totals.topLeadsRegion = byLeads[0] + ' (' + byLeads[2] + ')';
  if (byVisits) totals.topVisitsRegion = byVisits[0] + ' (' + byVisits[1] + ')';
  return totals;
}

function buildRegionsCitiesAttentionRows_(summary) {
  return summary.map(function(row) {
    const decision = classifyRegionSummary_(row);
    if (decision.group === 'Усилить / оставить' || decision.group === 'Мало данных') return null;
    return [
      decision.group === 'Снизить / не усиливать' ? 'Высокий' : 'Средний',
      row[0],
      decision.reason,
      decision.recommendation
    ];
  }).filter(Boolean).slice(0, 10);
}

function buildRegionsCitiesRegionSummary_(rows) {
  const summary = {};
  rows.forEach(function(row) {
    const region = row[1] || 'Не определён';
    if (!summary[region]) {
      summary[region] = {
        region: region,
        visits: 0,
        leads: 0,
        bounceSum: 0,
        depthSum: 0,
        durationSum: 0,
        count: 0
      };
    }
    summary[region].visits += num_(row[3]);
    summary[region].leads += num_(row[10]);
    summary[region].bounceSum += num_(row[13]);
    summary[region].depthSum += num_(row[14]);
    summary[region].durationSum += timeToSeconds_(row[15]);
    summary[region].count += 1;
  });

  return Object.keys(summary).map(function(region) {
    const item = summary[region];
    const cr = item.visits ? item.leads / item.visits : 0;
    const bounces = item.count ? item.bounceSum / item.count : 0;
    const depth = item.count ? item.depthSum / item.count : 0;
    const duration = item.count ? item.durationSum / item.count : 0;
    const conclusion = buildRegionsCitiesConclusion_({
      visits: item.visits,
      leads: item.leads,
      cr: cr,
      bounces: bounces,
      depth: depth,
      duration: duration
    });
    return [
      region,
      item.visits,
      item.leads,
      cr,
      bounces,
      depth,
      secondsToTime_(duration),
      conclusion.text,
      conclusion.recommendation
    ];
  }).sort(function(a, b) {
    return b[1] - a[1];
  });
}

function buildRegionsCitiesMonthSummary_(rows) {
  const monthSummary = {};
  rows.forEach(function(row) {
    const month = row[0] || 'Не определён';
    if (!monthSummary[month]) {
      monthSummary[month] = { visits: 0, leads: 0 };
    }
    monthSummary[month].visits += num_(row[3]);
    monthSummary[month].leads += num_(row[10]);
  });
  return Object.keys(monthSummary).sort().map(function(month) {
    return [month, monthSummary[month].visits, monthSummary[month].leads];
  });
}

function writeRegionsCitiesInsights_(sheet, rows, errors) {
  const layout = getRegionsCitiesLayout_();
  const summary = buildRegionsCitiesRegionSummary_(rows);

  const insightRows = summary.map(function(row) {
    const decision = classifyRegionSummary_(row);
    return [decision.group, row[0], decision.reason, decision.recommendation];
  }).sort(function(a, b) {
    const order = {
      'Усилить / оставить': 1,
      'Проверить': 2,
      'Снизить / не усиливать': 3,
      'Мало данных': 4
    };
    return (order[a[0]] || 9) - (order[b[0]] || 9);
  }).slice(0, 25);

  if (errors.length) {
    insightRows.push(['Ошибки обновления', 'См. Logger', errors.join(' | '), 'Проверить регионы из списка ошибок.']);
  }

  sheet.getRange(layout.insightsStartRow, 1, 25, 4).clearContent();
  if (insightRows.length) {
    sheet.getRange(layout.insightsStartRow, 1, insightRows.length, 4).setValues(insightRows);
  }
}

function classifyRegionSummary_(row) {
  const visits = num_(row[1]);
  const leads = num_(row[2]);
  const cr = num_(row[3]);
  const bounces = num_(row[4]);
  const depth = num_(row[5]);
  const duration = timeToSeconds_(row[6]);

  if (visits > 0 && visits < 30) {
    return {
      group: 'Мало данных',
      reason: 'Мало данных: ' + visits + ' визитов, заявок ' + leads + '.',
      recommendation: 'Не принимать резких решений, дождаться накопления статистики.'
    };
  }

  if (visits >= 50 && leads === 0) {
    return {
      group: 'Снизить / не усиливать',
      reason: 'Высокий трафик без заявок: ' + visits + ' визитов, заявок нет.',
      recommendation: 'Не усиливать; проверить цели, звонки, CRM и посадочные страницы.'
    };
  }

  if (leads > 0 && cr >= 0.01 && bounces <= 35 && duration >= 45) {
    return {
      group: 'Усилить / оставить',
      reason: 'Есть заявки: ' + leads + ', CR ' + Utilities.formatString('%.2f%%', cr * 100) + ', качество приемлемое.',
      recommendation: 'Оставить в работе или аккуратно усилить после проверки качества лидов.'
    };
  }

  if (leads > 0) {
    const reasons = ['есть заявки: ' + leads];
    if (cr < 0.01) reasons.push('слабый CR ' + Utilities.formatString('%.2f%%', cr * 100));
    if (bounces > 35) reasons.push('высокий отказ ' + Utilities.formatString('%.1f%%', bounces));
    if (duration < 45) reasons.push('слабое время на сайте ' + secondsToTime_(duration));
    if (depth < 1.5) reasons.push('слабая глубина ' + Utilities.formatString('%.2f', depth));
    return {
      group: 'Проверить',
      reason: reasons.join('; ') + '.',
      recommendation: 'Проверить CRM/звонки/цели, качество заявок и посадочные страницы.'
    };
  }

  if (bounces > 35 || duration < 45 || depth < 1.5) {
    return {
      group: 'Проверить',
      reason: 'Слабое поведение: отказы ' + Utilities.formatString('%.1f%%', bounces) + ', глубина ' + Utilities.formatString('%.2f', depth) + ', время ' + secondsToTime_(duration) + '.',
      recommendation: 'Проверить поисковые запросы, посадочную страницу и корректность целей.'
    };
  }

  return {
    group: 'Проверить',
    reason: 'Есть трафик без уверенного вывода: ' + visits + ' визитов, заявок ' + leads + '.',
    recommendation: 'Нужна проверка CRM/звонков/целей и накопление статистики.'
  };
}

function buildRegionsCitiesCharts_(sheet) {
  const existingCharts = sheet.getCharts();
  existingCharts.forEach(function(chart) {
    sheet.removeChart(chart);
  });

  const layout = getRegionsCitiesLayout_();
  const summaryHeaderRow = layout.helperSummaryHeaderRow;
  const summaryStartRow = layout.helperSummaryStartRow;
  const summaryLastRow = Math.min(summaryStartRow + 29, Math.max(summaryStartRow, sheet.getLastRow()));

  if (summaryLastRow >= summaryStartRow) {
    sheet.insertChart(sheet.newChart()
      .setChartType(Charts.ChartType.BAR)
      .addRange(sheet.getRange(summaryHeaderRow, 27, summaryLastRow - summaryHeaderRow + 1, 2))
      .setPosition(18, 1, 0, 0)
      .setOption('title', 'Топ регионов по визитам')
      .build());

    sheet.insertChart(sheet.newChart()
      .setChartType(Charts.ChartType.BAR)
      .addRange(sheet.getRange(summaryHeaderRow, 27, summaryLastRow - summaryHeaderRow + 1, 1))
      .addRange(sheet.getRange(summaryHeaderRow, 29, summaryLastRow - summaryHeaderRow + 1, 1))
      .setPosition(18, 8, 0, 0)
      .setOption('title', 'Топ регионов по целям/заявкам')
      .build());

    sheet.insertChart(sheet.newChart()
      .setChartType(Charts.ChartType.COLUMN)
      .addRange(sheet.getRange(summaryHeaderRow, 27, summaryLastRow - summaryHeaderRow + 1, 1))
      .addRange(sheet.getRange(summaryHeaderRow, 30, summaryLastRow - summaryHeaderRow + 1, 1))
      .setPosition(33, 1, 0, 0)
      .setOption('title', 'CR по регионам')
      .build());

    sheet.insertChart(sheet.newChart()
      .setChartType(Charts.ChartType.BAR)
      .addRange(sheet.getRange('AJ2:AK25'))
      .setPosition(33, 8, 0, 0)
      .setOption('title', 'Регионы с трафиком, но без заявок')
      .build());
  }

  sheet.insertChart(sheet.newChart()
    .setChartType(Charts.ChartType.LINE)
    .addRange(sheet.getRange('AJ35:AK60'))
    .setPosition(18, 15, 0, 0)
    .setOption('title', 'Динамика визитов по месяцам')
    .build());

  sheet.insertChart(sheet.newChart()
    .setChartType(Charts.ChartType.LINE)
    .addRange(sheet.getRange('AJ35:AJ60'))
    .addRange(sheet.getRange('AL35:AL60'))
    .setPosition(33, 15, 0, 0)
    .setOption('title', 'Динамика целей/заявок по месяцам')
    .build());
}

function buildTrafficQuality_(data) {
  if (data.visits < 30) return 'Мало данных';
  if (data.leads > 0 && data.cr >= 0.03 && data.bounces <= 25) return 'Хорошее качество';
  if (data.leads > 0 && data.cr >= 0.01) return 'Среднее качество';
  if (data.visits >= 100 && data.leads === 0) return 'Трафик без заявок';
  if (data.bounces >= 35 || data.depth < 1.5) return 'Слабое качество';
  return 'Нужно перепроверить';
}

function buildRegionsCitiesConclusion_(data) {
  if (data.visits === 0) {
    return {
      text: 'По региону/городу нет трафика.',
      recommendation: 'Не принимать решение без данных.'
    };
  }

  if (data.visits < 30) {
    return {
      text: 'По региону/городу мало данных.',
      recommendation: 'Продолжить сбор статистики, не снижать резко.'
    };
  }

  if (data.leads > 0 && data.cr >= 0.03 && data.bounces <= 25) {
    return {
      text: 'Регион даёт заявки и нормальное качество трафика.',
      recommendation: 'Регион стоит усилить или оставить в приоритете.'
    };
  }

  if (data.leads > 0) {
    return {
      text: 'Регион даёт заявки, но качество или CR требуют проверки.',
      recommendation: 'Оставить, проверить CRM/звонки/цели и посадочные страницы.'
    };
  }

  if (data.visits >= 100 && data.leads === 0) {
    return {
      text: 'Регион даёт трафик, но не даёт заявки.',
      recommendation: 'Снизить, исключить или перепроверить цели и качество трафика.'
    };
  }

  if (data.bounces >= 35 || data.depth < 1.5) {
    return {
      text: 'Регион даёт слабое качество трафика.',
      recommendation: 'Снизить ставки/приоритет и проверить поисковые запросы.'
    };
  }

  return {
    text: 'По региону есть трафик, но вывод неоднозначный.',
    recommendation: 'Нужна проверка CRM/звонков/целей и накопление статистики.'
  };
}

function getGoalMappingReport_(goals) {
  const found = [];
  const missing = [];
  Object.keys(goals).forEach(function(key) {
    if (goals[key]) {
      found.push(key + '=' + goals[key]);
    } else {
      missing.push(key);
    }
  });
  return {
    found: found,
    missing: missing
  };
}

function getSelectedRegionsGoalIds_(goals) {
  return TDM_REPORTS_CONFIG.regionsGoalKeys.map(function(key) {
    return goals[key];
  });
}

function getSelectedRegionsGoalReport_(goals) {
  const found = [];
  const missing = [];
  TDM_REPORTS_CONFIG.regionsGoalKeys.forEach(function(key) {
    const label = TDM_REPORTS_CONFIG.goalLabels[key] || key;
    if (goals[key]) {
      found.push(label + '=' + goals[key]);
    } else {
      missing.push(label);
    }
  });
  return {
    found: found,
    missing: missing
  };
}

function getCurrentAvailableDate_() {
  const date = new Date();
  date.setDate(date.getDate() - 1);
  return formatDate_(date);
}

function timeToSeconds_(value) {
  if (typeof value === 'number') return value;
  const parts = String(value || '').split(':').map(Number);
  if (parts.length !== 3 || parts.some(isNaN)) return 0;
  return parts[0] * 3600 + parts[1] * 60 + parts[2];
}

function cleanRegionCityName_(value) {
  return String(value || '')
    .replace(/\s+/g, ' ')
    .replace(/\.\s*Предлагаю.*$/i, '')
    .replace(/\s+с областью$/i, '')
    .trim();
}

function updateProductsReport() {
  const ss = SpreadsheetApp.openById(TDM_REPORTS_CONFIG.spreadsheetId);
  const sheet = ss.getSheetByName(TDM_REPORTS_CONFIG.sheets.products);
  if (!sheet) throw new Error('Не найден лист: ' + TDM_REPORTS_CONFIG.sheets.products);

  logAttributionModel_();

  const counterId = getCounterIdFromSheet_(sheet);
  const period = getReportPeriod_(sheet, 'week');
  const goals = mapGoalsByName(getMetrikaGoals(counterId));
  const products = getProductsForWeeklyReport_(ss);

  const rows = [];
  products.forEach(function(base) {
    const url = base.url;
    if (!url) return;

    const allTraffic = fetchProductRows_(counterId, period, goals, url, '');
    const adTraffic = fetchProductRows_(counterId, period, goals, url, buildPaidTrafficFilter_());
    const otherTraffic = fetchProductRows_(counterId, period, goals, url, 'NOT(' + buildPaidTrafficFilter_() + ')');

    appendProductOutputRows_(rows, base, 'Рекламный трафик', adTraffic, allTraffic);
    appendProductOutputRows_(rows, base, 'Остальные источники', otherTraffic, allTraffic);
  });

  writeRowsToSheet(sheet, rows, [
    'Неделя',
    'ID позиции',
    'Группа',
    'Название',
    'URL',
    'Тип источника',
    'Визиты',
    'Пользователи',
    'Просмотры страницы',
    'Цели/заявки',
    'CR',
    'Отказы',
    'Глубина',
    'Время на сайте',
    'Доля от всего трафика страницы',
    'Вывод по спросу',
    'Рекомендация',
    'Дата обновления',
    'Комментарий'
  ]);
}

function fetchProductRows_(counterId, period, goals, url, extraFilter) {
  const filters = [
    "EXISTS(ym:pv:URL=@" + quoteMetrikaValue_(url) + ")"
  ];
  if (extraFilter) filters.push(extraFilter);

  const data = fetchMetrikaData({
    counterId: counterId,
    date1: period.date1,
    date2: period.date2,
    dimensions: ['ym:s:datePeriod<group>'],
    metrics: buildProductMetrics_(goals),
    params: { group: 'week' },
    filters: filters.join(' AND ')
  });

  const byWeek = {};
  (data.data || []).forEach(function(row) {
    const week = dimensionName_(row, 0);
    byWeek[week] = row.metrics || [];
  });
  return byWeek;
}

function getProductsForWeeklyReport_(ss) {
  const promoted = getPromotedPositionProducts_(ss);
  const electro = TDM_REPORTS_CONFIG.electroEquipmentPages.map(function(item) {
    return Object.assign({}, item);
  });

  return uniqueProductsByUrl_(promoted.concat(electro));
}

function getPromotedPositionProducts_(ss) {
  const sheet = ss.getSheetByName(TDM_REPORTS_CONFIG.sheets.promotedPositions);
  if (!sheet) {
    Logger.log('Предупреждение: не найден лист "' + TDM_REPORTS_CONFIG.sheets.promotedPositions + '". Новинки не будут добавлены в отчет.');
    return [];
  }

  const headerRow = findAnyHeaderRow_(sheet, ['URL']);
  if (!headerRow) {
    Logger.log('Предупреждение: на листе "' + TDM_REPORTS_CONFIG.sheets.promotedPositions + '" не найдена колонка URL.');
    return [];
  }

  const values = sheet.getRange(headerRow, 1, sheet.getLastRow() - headerRow + 1, sheet.getLastColumn()).getValues();
  const headers = values[0].map(String);
  const urlCol = findHeaderIndex_(headers, ['URL', 'Адрес', 'Страница', 'Ссылка']);
  const includeCol = findHeaderIndex_(headers, ['Включить в анализ', 'Включить']);
  const idCol = findHeaderIndex_(headers, ['ID позиции', 'ID', 'Артикул']);
  const groupCol = findHeaderIndex_(headers, ['Группа', 'Категория', 'Раздел']);
  const nameCol = findHeaderIndex_(headers, ['Название', 'Товар', 'Направление', 'Наименование']);

  if (urlCol === -1) return [];

  const result = [];
  for (let i = 1; i < values.length; i++) {
    const row = values[i];
    const url = String(row[urlCol] || '').trim();
    if (!url) continue;

    // Если колонка "Включить в анализ" есть, берем только строки со значением "Да".
    if (includeCol !== -1 && normalize_(row[includeCol]) !== 'да') continue;

    result.push({
      id: idCol !== -1 ? row[idCol] : 'new-' + i,
      group: groupCol !== -1 && row[groupCol] ? row[groupCol] : 'Новинки',
      name: nameCol !== -1 && row[nameCol] ? row[nameCol] : 'Новинка',
      url: url
    });
  }

  return result;
}

function uniqueProductsByUrl_(products) {
  const seen = {};
  return products.filter(function(product) {
    const key = normalizeUrl_(product.url);
    if (!key || seen[key]) return false;
    seen[key] = true;
    return true;
  });
}

function appendProductOutputRows_(rows, base, sourceType, sourceRows, allRows) {
  Object.keys(sourceRows).sort().forEach(function(week) {
    const m = sourceRows[week] || [];
    const allM = allRows[week] || [];
    const visits = num_(m[0]);
    const users = num_(m[1]);
    const pageviews = num_(m[2]);
    const leads = num_(m[3]);
    const bounces = num_(m[4]);
    const depth = num_(m[5]);
    const duration = num_(m[6]);
    const allVisits = num_(allM[0]);
    const cr = visits ? leads / visits : 0;
    const share = allVisits ? visits / allVisits : 0;
    const conclusion = buildConclusion({
      type: 'product',
      visits: visits,
      cr: cr,
      bounces: bounces,
      conversions: leads,
      sourceType: sourceType
    });

    rows.push([
      week,
      base.id,
      base.group,
      base.name,
      base.url,
      sourceType,
      visits,
      users,
      pageviews,
      leads,
      cr,
      bounces,
      depth,
      secondsToTime_(duration),
      share,
      conclusion.text,
      conclusion.recommendation,
      new Date(),
      ''
    ]);
  });
}

function fetchMetrikaData(options) {
  const token = getRequiredScriptProperty_(TDM_REPORTS_CONFIG.tokenPropertyName);
  const params = Object.assign({}, options.params || {}, {
    ids: options.counterId,
    date1: options.date1,
    date2: options.date2,
    dimensions: options.dimensions.join(','),
    metrics: options.metrics.join(','),
    filters: options.filters || '',
    attribution: TDM_REPORTS_CONFIG.attribution,
    accuracy: TDM_REPORTS_CONFIG.accuracy,
    lang: TDM_REPORTS_CONFIG.lang,
    limit: 100000,
    timezone: TDM_REPORTS_CONFIG.timezone
  });

  const url = TDM_REPORTS_CONFIG.metrikaApiUrl + '?' + tdmReportsToQueryString_(params);
  const response = UrlFetchApp.fetch(url, {
    method: 'get',
    headers: { Authorization: 'OAuth ' + token },
    muteHttpExceptions: true
  });

  const code = response.getResponseCode();
  const text = response.getContentText();
  if (code < 200 || code >= 300) {
    throw new Error('Ошибка API Метрики ' + code + ': ' + text);
  }

  const parsed = JSON.parse(text);
  Logger.log('Metrika attribution in request: ' + TDM_REPORTS_CONFIG.attribution);
  return parsed;
}

function getMetrikaGoals(counterId) {
  const token = getRequiredScriptProperty_(TDM_REPORTS_CONFIG.tokenPropertyName);
  const url = Utilities.formatString(TDM_REPORTS_CONFIG.metrikaGoalsUrl, counterId);
  const response = UrlFetchApp.fetch(url, {
    method: 'get',
    headers: { Authorization: 'OAuth ' + token },
    muteHttpExceptions: true
  });

  const code = response.getResponseCode();
  const text = response.getContentText();
  if (code < 200 || code >= 300) {
    throw new Error('Не удалось получить цели Метрики для счётчика ' + counterId + ' — ' + code + ': ' + text);
  }

  return (JSON.parse(text).goals || []);
}

function mapGoalsByName(goals) {
  const result = {};
  Object.keys(TDM_REPORTS_CONFIG.goalNameParts).forEach(function(key) {
    const parts = TDM_REPORTS_CONFIG.goalNameParts[key];
    const goal = goals.find(function(g) {
      const name = normalize_(g.name);
      return parts.every(function(part) {
        return name.indexOf(normalize_(part)) !== -1;
      });
    });

    if (goal) {
      result[key] = goal.id;
    } else {
      result[key] = null;
      Logger.log('Предупреждение: цель не найдена: ' + key + '. Будет записан 0.');
    }
  });
  return result;
}

function writeRowsToSheet(sheet, rows, expectedHeaders) {
  const output = findBlock_(sheet, 'output');
  const headerRow = findHeaderRowInRange_(sheet, output.startRow, output.endRow, expectedHeaders.slice(0, 3));
  if (!headerRow) {
    throw new Error('Не найдена строка заголовков блока вывода на листе: ' + sheet.getName());
  }

  const startCol = findFirstHeaderColumn_(sheet, headerRow, expectedHeaders[0]);
  const width = expectedHeaders.length;
  const firstDataRow = headerRow + 1;
  const lastRow = Math.max(sheet.getLastRow(), firstDataRow);
  const rowsToClear = Math.max(0, lastRow - firstDataRow + 1);

  if (rowsToClear > 0) {
    sheet.getRange(firstDataRow, startCol, rowsToClear, width).clearContent();
  }

  if (rows.length > 0) {
    sheet.getRange(firstDataRow, startCol, rows.length, width).setValues(rows);
  }
}

function buildConclusion(data) {
  if (data.visits === 0) {
    return {
      text: 'Трафика нет или он слишком мал для вывода.',
      recommendation: data.type === 'region' ? 'Проверить, нужен ли регион в продвижении.' : 'Не принимать решение без трафика.'
    };
  }

  if (data.conversions > 0 && data.cr >= 0.03 && data.bounces <= 25) {
    return {
      text: 'Есть спрос и нормальная вовлеченность.',
      recommendation: data.type === 'region' ? 'Оставить или усилить регион.' : 'Усилить направление и проверить рекламные связки.'
    };
  }

  if (data.conversions > 0 && data.cr < 0.03) {
    return {
      text: 'Конверсии есть, но CR слабый.',
      recommendation: 'Проверить посадочную страницу, оффер и качество трафика.'
    };
  }

  if (data.visits >= 100 && data.conversions === 0) {
    return {
      text: 'Трафик есть, заявок нет.',
      recommendation: data.type === 'region' ? 'Снизить приоритет или проверить релевантность региона.' : 'Не усиливать до проверки спроса и страницы.'
    };
  }

  return {
    text: 'Данных пока недостаточно для уверенного вывода.',
    recommendation: 'Продолжить сбор статистики и вернуться к оценке позже.'
  };
}

function buildProductMetrics_(goals) {
  return [
    'ym:s:visits',
    'ym:s:users',
    'ym:s:pageviews',
    sumGoalMetrics_([
      goals.phoneClick,
      goals.emailClick,
      goals.talkMe,
      goals.ecommercePurchase,
      goals.roistatLead,
      goals.roistatCall
    ]),
    'ym:s:bounceRate',
    'ym:s:pageDepth',
    'ym:s:avgVisitDurationSeconds'
  ];
}

function goalMetric_(goalId) {
  if (!goalId) return 'ym:s:visits*0';
  return 'ym:s:goal' + goalId + 'visits';
}

function sumGoalMetrics_(goalIds) {
  const metrics = goalIds.filter(Boolean).map(function(goalId) {
    return 'ym:s:goal' + goalId + 'visits';
  });
  if (metrics.length === 0) return 'ym:s:visits*0';
  return metrics.join('+');
}

function buildPaidTrafficFilter_() {
  return [
    "ym:s:<attribution>TrafficSource=='ad'",
    "ym:s:<attribution>UTMMedium=@'cpc'",
    "ym:s:<attribution>UTMMedium=@'paid'",
    "ym:s:<attribution>UTMMedium=@'ppc'"
  ].join(' OR ');
}

function getEnabledRows_(sheet, options) {
  const blocks = findBlocks_(sheet, options.blockType);
  let result = [];
  blocks.forEach(function(block) {
    result = result.concat(getEnabledRowsFromBlock_(sheet, block, options.requiredHeaders));
  });
  return result;
}

function getEnabledRowsFromBlock_(sheet, block, requiredHeaders) {
  const headerRow = findHeaderRowInRange_(sheet, block.startRow, block.endRow, requiredHeaders);
  if (!headerRow) {
    Logger.log('Предупреждение: пропущен блок без нужных заголовков на листе ' + sheet.getName());
    return [];
  }

  const values = sheet.getRange(headerRow, 1, block.endRow - headerRow + 1, sheet.getLastColumn()).getValues();
  const headers = values[0].map(String);
  const includeCol = findHeaderIndex_(headers, ['Включить в анализ', 'Включить']);
  if (includeCol === -1) throw new Error('Не найдена колонка "Включить в анализ"');

  const result = [];
  for (let i = 1; i < values.length; i++) {
    const row = values[i];
    const include = normalize_(row[includeCol]);
    if (include !== 'да') continue;

    const item = { _headers: headers, _row: row };
    headers.forEach(function(header, index) {
      item[normalizeHeader_(header)] = row[index];
    });
    result.push(item);
  }
  return result;
}

function getByHeader_(item, names) {
  for (let i = 0; i < names.length; i++) {
    const value = item[normalizeHeader_(names[i])];
    if (value !== '' && value !== null && value !== undefined) return value;
  }
  return '';
}

function getCounterIdFromSheet_(sheet) {
  const propertyCounterId = PropertiesService.getScriptProperties().getProperty('METRIKA_COUNTER_ID');
  if (propertyCounterId) return propertyCounterId;
  if (TDM_REPORTS_CONFIG.defaultMetrikaCounterId) return TDM_REPORTS_CONFIG.defaultMetrikaCounterId;

  const finder = sheet.createTextFinder('Счётчик Метрики').matchCase(false).findNext()
    || sheet.createTextFinder('Счетчик Метрики').matchCase(false).findNext();
  if (!finder) {
    throw new Error('В листе не найдена строка "Счётчик Метрики" и не задано Script Property METRIKA_COUNTER_ID');
  }

  const row = finder.getRow();
  const col = finder.getColumn();
  const values = sheet.getRange(row, col, 1, Math.max(1, sheet.getLastColumn() - col)).getValues()[0];
  for (let i = 1; i < values.length; i++) {
    const value = String(values[i] || '').trim();
    if (/^\d+$/.test(value)) return value;
  }

  throw new Error('Рядом с "Счётчик Метрики" не найден ID счетчика и не задано Script Property METRIKA_COUNTER_ID');
}

function getReportPeriod_(sheet, group) {
  const explicit = readDateRangeFromSheet_(sheet);
  if (explicit.date1 && explicit.date2) return explicit;

  const now = new Date();
  const year = now.getFullYear();
  const date1 = new Date(year, 0, 1);
  const date2 = new Date(now.getFullYear(), now.getMonth(), now.getDate() - 1);
  return {
    date1: formatDate_(date1),
    date2: formatDate_(date2),
    group: group
  };
}

function readDateRangeFromSheet_(sheet) {
  const result = { date1: '', date2: '' };
  const data = sheet.getDataRange().getValues();
  for (let r = 0; r < data.length; r++) {
    for (let c = 0; c < data[r].length; c++) {
      const label = normalize_(data[r][c]);
      if (label.indexOf('дата начала') !== -1 || label.indexOf('начало периода') !== -1) {
        result.date1 = findDateRight_(data[r], c);
      }
      if (label.indexOf('дата окончания') !== -1 || label.indexOf('конец периода') !== -1) {
        result.date2 = findDateRight_(data[r], c);
      }
    }
  }
  return result;
}

function findDateRight_(row, startIndex) {
  for (let i = startIndex + 1; i < row.length; i++) {
    if (row[i] instanceof Date) return formatDate_(row[i]);
    const value = String(row[i] || '').trim();
    if (/^\d{4}-\d{2}-\d{2}$/.test(value)) return value;
  }
  return '';
}

function findBlock_(sheet, type) {
  const blocks = findBlocks_(sheet, type);
  if (!blocks.length) {
    const marker = type === 'output' ? TDM_REPORTS_CONFIG.markers.outputBlock : TDM_REPORTS_CONFIG.markers.inputBlock;
    throw new Error('Не найден блок: ' + marker + ' на листе ' + sheet.getName());
  }
  return blocks[0];
}

function findBlocks_(sheet, type) {
  const marker = type === 'output' ? TDM_REPORTS_CONFIG.markers.outputBlock : TDM_REPORTS_CONFIG.markers.inputBlock;
  const found = sheet.createTextFinder(marker).matchCase(false).findAll();
  if (!found.length) return [];

  const markerCells = found.map(function(r) {
    return { row: r.getRow(), col: r.getColumn() };
  }).sort(function(a, b) { return a.row - b.row; });
  const outputCells = sheet.createTextFinder(TDM_REPORTS_CONFIG.markers.outputBlock).matchCase(false).findAll()
    .map(function(r) { return { row: r.getRow(), col: r.getColumn() }; })
    .sort(function(a, b) { return a.row - b.row; });
  const allKnownBlockRows = sheet.createTextFinder('Блок').matchCase(false).findAll()
    .map(function(r) { return { row: r.getRow(), col: r.getColumn() }; })
    .sort(function(a, b) { return a.row - b.row; });

  return markerCells.map(function(markerCell) {
    const markerRow = markerCell.row;
    const markerCol = markerCell.col;
    const startRow = markerRow + 1;
    const nextBlock = allKnownBlockRows.filter(function(cell) {
      const sameArea = (markerCol <= 20 && cell.col <= 20) || (markerCol > 20 && cell.col > 20);
      return sameArea && cell.row > markerRow;
    })[0];
    let endRow = nextBlock ? nextBlock.row - 1 : sheet.getLastRow();

    if (type === 'input' && outputCells.length) {
      const nextOutput = outputCells.filter(function(cell) {
        return cell.col === markerCol && cell.row > markerRow;
      })[0];
      if (nextOutput) endRow = Math.min(endRow, nextOutput.row - 1);
    }

    return { startRow: startRow, endRow: endRow };
  });
}

function findHeaderRowInRange_(sheet, startRow, endRow, requiredHeaders) {
  const values = sheet.getRange(startRow, 1, endRow - startRow + 1, sheet.getLastColumn()).getValues();
  for (let i = 0; i < values.length; i++) {
    const headers = values[i].map(normalizeHeader_);
    const ok = requiredHeaders.every(function(header) {
      return headers.indexOf(normalizeHeader_(header)) !== -1;
    });
    if (ok) return startRow + i;
  }
  return null;
}

function findAnyHeaderRow_(sheet, requiredHeaders) {
  return findHeaderRowInRange_(sheet, 1, sheet.getLastRow(), requiredHeaders);
}

function findFirstHeaderColumn_(sheet, headerRow, headerName) {
  const headers = sheet.getRange(headerRow, 1, 1, sheet.getLastColumn()).getValues()[0].map(normalizeHeader_);
  const index = headers.indexOf(normalizeHeader_(headerName));
  if (index === -1) throw new Error('Не найден заголовок: ' + headerName);
  return index + 1;
}

function findHeaderIndex_(headers, names) {
  const normalized = headers.map(normalizeHeader_);
  for (let i = 0; i < names.length; i++) {
    const index = normalized.indexOf(normalizeHeader_(names[i]));
    if (index !== -1) return index;
  }
  return -1;
}

function dimensionName_(row, index) {
  const dim = row.dimensions && row.dimensions[index];
  return dim ? (dim.name || dim.id || '') : '';
}

function quoteMetrikaValue_(value) {
  return "'" + String(value).replace(/\\/g, '\\\\').replace(/'/g, "\\'") + "'";
}

function tdmReportsToQueryString_(params) {
  return Object.keys(params)
    .filter(function(key) { return params[key] !== '' && params[key] !== null && params[key] !== undefined; })
    .map(function(key) {
      return encodeURIComponent(key) + '=' + encodeURIComponent(params[key]);
    })
    .join('&');
}

function getRequiredScriptProperty_(name) {
  const value = PropertiesService.getScriptProperties().getProperty(name);
  if (!value) {
    throw new Error('Не найдено Script Property: ' + name);
  }
  return value;
}

function logAttributionModel_() {
  Logger.log('Все запросы к Метрике используют attribution=' + TDM_REPORTS_CONFIG.attribution);
}

function normalize_(value) {
  return String(value || '').trim().toLowerCase().replace(/ё/g, 'е');
}

function normalizeHeader_(value) {
  return normalize_(value).replace(/\s+/g, ' ');
}

function normalizeUrl_(value) {
  return String(value || '').trim().replace(/[?#].*$/, '').replace(/\/?$/, '/').toLowerCase();
}

function num_(value) {
  const n = Number(value);
  return isNaN(n) ? 0 : n;
}

function secondsToTime_(seconds) {
  seconds = Math.round(num_(seconds));
  const h = Math.floor(seconds / 3600);
  const m = Math.floor((seconds % 3600) / 60);
  const s = seconds % 60;
  return [h, m, s].map(function(part) {
    return String(part).padStart(2, '0');
  }).join(':');
}

function formatDate_(date) {
  return Utilities.formatDate(date, Session.getScriptTimeZone(), 'yyyy-MM-dd');
}



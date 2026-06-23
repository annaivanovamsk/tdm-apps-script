// VAG-House ENO weekly report — one ready file v26
// Вставь этот код целиком в Apps Script вместо текущего файла ЕНО.
// Рабочая логика: отчёт всегда строится за предыдущую полную неделю
// с понедельника по воскресенье и может обновляться каждый понедельник в 10:00.
//
// Что запускать:
// 1) fillEnoPreviousFullWeekReport_v26() — вручную сформировать отчёт за предыдущую полную неделю ПН–ВС.
// 2) fillEnoWeek_2026_06_01_2026_06_07_v26() — разово пересобрать нужный сейчас период 01.06–07.06.
// 3) createEnoMonday10Trigger_v26() — один раз создать автозапуск каждый понедельник в 10:00.
// 4) deleteEnoMonday10Triggers_v26() — удалить автозапуск ЕНО, если понадобится.
// Обновлено v19: добавлены все цели M:U, включая маршрут 547214611.
// Исправлено оформление комментария: старый хвост шаблона очищается, комментарий пишется отдельным блоком.
// v20: исправлена ошибка объединения закрепленных и незакрепленных столбцов.
// v21: комментарий пишется в колонку A; заголовки выделены жирным и цветной заливкой.
// v22: цели M:U берутся из Метрики как "Достижения цели" (reaches), чтобы совпадать с отчетом Метрики.
// v23: исправлен формат колонки K "Отказы (%)" — значение приводится к доле для процентного формата.
// v24: Метрика берёт отдельный METRIKA_TOKEN, при 403 отчёт не падает, а оставляет цели из Директа и пишет предупреждение в комментарий.
// v25: цели Метрики берутся через обычный namespace ym:s, без ym:ad/direct_client_logins. Достаточно доступа к счётчику.
// v26: все внутренние имена и функции запуска имеют суффикс _v26, чтобы не конфликтовать со старым кодом в проекте.
// KPI / Факт сумма конверсий считается только по T + U.

const VAG_ENO_V26_CONFIG = {
  SHEET_NAME: 'ЕНО',
  CLIENT_LOGIN: 'vag-house98',
  TIMEZONE: 'Europe/Moscow',
  ATTRIBUTION_MODEL: 'AUTO',
  METRIKA_COUNTER_ID: '99052519',
  METRIKA_ATTRIBUTION: 'automatic',

  TEMPLATE_TOP_ROW: 624,
  TEMPLATE_COPY_ROWS: 70,
  MONTH_BUDGET: 120000,

  GOOD_CPA_LIMIT: 2500,
  MIDDLE_CPA_LIMIT: 5000,
  RED_COST_LIMIT: 3000,
  YELLOW_COST_LIMIT: 1500,
  YELLOW_CLICKS_LIMIT: 10,

  GOALS: {
    bb60sec2pages: '555861179',    // M — B-B_60сек+2стр
    phone8835: '511787501',        // N — Клик на 79110008835
    phone8865: '511788019',        // O — Клик на 79110008865
    max: '511788150',              // P — Клик на Макс
    telegramSlider: '511791937',   // Q — Успешно пройден слайдер для перехода в TG
    route: '547214611',            // R — Клик по кнопке "Маршрут"
    uisCalls: '517626340',         // S — Звонки Юис
    uisUniqueCalls: '562354501',   // T — Звонки_Уник - ЮИС
    uisLeads: '517626592'          // U — Заявки Юис
  }
};

const VAG_ENO_V26_RUNTIME = {
  metrikaWarning: ''
};

const VAG_ENO_V26_TRIGGER_HANDLERS = [
  'fillEnoPreviousFullWeekReport_v26',
  'fillEnoPreviousFullWeek',
  'createEnoWeeklyTriggerMonday9am'
];

// =====================
// ЗАПУСК
// =====================

function fillEnoPreviousFullWeekReport_v26() {
  const period = VAG_ENO_V26_getPreviousFullWeekPeriod_();
  VAG_ENO_V26_fillWeeklyReportForPeriod_(period.dateFrom, period.dateTo);
}

function fillEnoWeek_2026_06_01_2026_06_07_v26() {
  VAG_ENO_V26_fillWeeklyReportForPeriod_('2026-06-01', '2026-06-07');
}

function createEnoMonday10Trigger_v26() {
  deleteEnoMonday10Triggers_v26();

  ScriptApp.newTrigger('fillEnoPreviousFullWeekReport_v26')
    .timeBased()
    .onWeekDay(ScriptApp.WeekDay.MONDAY)
    .atHour(10)
    .create();
}

function deleteEnoMonday10Triggers_v26() {
  ScriptApp.getProjectTriggers()
    .filter(trigger => VAG_ENO_V26_TRIGGER_HANDLERS.indexOf(trigger.getHandlerFunction()) !== -1)
    .forEach(trigger => ScriptApp.deleteTrigger(trigger));
}

function checkEnoMetrikaAccess_v26() {
  const period = VAG_ENO_V26_getPreviousFullWeekPeriod_();
  const data = VAG_ENO_V26_getMetrikaGoalReachesByCampaign_(period.dateFrom, period.dateTo);
  Logger.log('Доступ к Метрике есть. Кампаний с целями найдено: ' + Object.keys(data).length);
}

function VAG_ENO_V26_fillWeeklyReportForPeriod_(dateFrom, dateTo) {
  VAG_ENO_V26_RUNTIME.metrikaWarning = '';

  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(VAG_ENO_V26_CONFIG.SHEET_NAME);

  if (!sheet) {
    throw new Error('Не найдена вкладка: ' + VAG_ENO_V26_CONFIG.SHEET_NAME);
  }

  const rows = VAG_ENO_V26_getCampaignReport_(dateFrom, dateTo);
  VAG_ENO_V26_overlayMetrikaGoalReaches_(rows, dateFrom, dateTo);
  rows.sort((a, b) => Number(b.cost || 0) - Number(a.cost || 0));

  const title = VAG_ENO_V26_buildTitle_(dateFrom, dateTo);
  const block = VAG_ENO_V26_getOrCreateBlock_(sheet, title);

  VAG_ENO_V26_setTitle_(sheet, block.topRow, title);
  VAG_ENO_V26_applyGoalHeaders_(sheet, block.headerRow);
  VAG_ENO_V26_fitRowsCount_(sheet, block, rows.length);

  const headerMap = VAG_ENO_V26_getHeaderMap_(sheet, block.headerRow);
  const dataStartRow = block.headerRow + 1;
  const rowsCount = Math.max(rows.length, 1);

  VAG_ENO_V26_clearRawCells_(sheet, headerMap, dataStartRow, rowsCount);
  VAG_ENO_V26_fillCampaignRows_(sheet, headerMap, dataStartRow, rows);
  VAG_ENO_V26_setFactFormulas_(sheet, headerMap, dataStartRow, block.totalRow);
  VAG_ENO_V26_cleanFormattingAfterFill_(sheet, headerMap, dataStartRow, rowsCount, rows);

  const comment = VAG_ENO_V26_buildComment_(rows, dateFrom, dateTo);
  const commentStartRow = VAG_ENO_V26_prepareCommentArea_(sheet, block);
  VAG_ENO_V26_writeComment_(sheet, commentStartRow, comment);

  SpreadsheetApp.flush();
}

// =====================
// ДИРЕКТ: базовые метрики + МЕТРИКА: достижения целей
// =====================

function VAG_ENO_V26_getCampaignReport_(dateFrom, dateTo) {
  const goalIds = Object.values(VAG_ENO_V26_CONFIG.GOALS).map(id => Number(id));

  const reportBody = {
    params: {
      SelectionCriteria: {
        DateFrom: dateFrom,
        DateTo: dateTo
      },
      Goals: goalIds,
      AttributionModels: [
        VAG_ENO_V26_CONFIG.ATTRIBUTION_MODEL
      ],
      FieldNames: [
        'CampaignName',
        'Impressions',
        'Clicks',
        'Cost',
        'AvgCpc',
        'AvgImpressionPosition',
        'AvgTrafficVolume',
        'BounceRate',
        'AvgPageviews',
        'Conversions'
      ],
      OrderBy: [
        { Field: 'Cost', SortOrder: 'DESCENDING' }
      ],
      ReportName: 'eno_weekly_' + VAG_ENO_V26_CONFIG.CLIENT_LOGIN + '_' + dateFrom + '_' + dateTo + '_' + Utilities.getUuid(),
      ReportType: 'CAMPAIGN_PERFORMANCE_REPORT',
      DateRangeType: 'CUSTOM_DATE',
      Format: 'TSV',
      IncludeVAT: 'YES',
      IncludeDiscount: 'NO'
    }
  };

  const text = VAG_ENO_V26_requestDirectReport_(reportBody);
  return VAG_ENO_V26_parseCampaignTsv_(text);
}

function VAG_ENO_V26_requestDirectReport_(reportBody) {
  const token = PropertiesService.getScriptProperties().getProperty('YANDEX_TOKEN');

  if (!token) {
    throw new Error('Не найден YANDEX_TOKEN в свойствах скрипта');
  }

  const url = 'https://api.direct.yandex.com/json/v5/reports';

  const options = {
    method: 'post',
    contentType: 'application/json',
    headers: {
      Authorization: 'Bearer ' + token,
      'Client-Login': VAG_ENO_V26_CONFIG.CLIENT_LOGIN,
      'Accept-Language': 'ru',
      processingMode: 'auto',
      returnMoneyInMicros: 'false',
      skipReportHeader: 'true',
      skipColumnHeader: 'false',
      skipReportSummary: 'true'
    },
    payload: JSON.stringify(reportBody),
    muteHttpExceptions: true
  };

  for (let attempt = 1; attempt <= 10; attempt++) {
    const response = UrlFetchApp.fetch(url, options);
    const code = response.getResponseCode();
    const text = response.getContentText();

    if (code === 200) return text;

    if (code === 201 || code === 202) {
      Utilities.sleep(5000);
      continue;
    }

    throw new Error('Ошибка Директа: ' + code + '\n' + text);
  }

  throw new Error('Отчёт долго формируется. Запусти позже.');
}

function VAG_ENO_V26_parseCampaignTsv_(tsvText) {
  const lines = String(tsvText || '').trim().split(/\r?\n/);
  if (lines.length < 2) return [];

  const headers = lines[0]
    .split('\t')
    .map(h => String(h).replace(/^\uFEFF/, '').trim());

  const idx = {};
  headers.forEach((h, i) => idx[h] = i);

  return lines.slice(1)
    .filter(line => String(line).trim() !== '')
    .map(line => {
      const row = line.split('\t');
      const campaignName = String(row[idx['CampaignName']] || '').trim();

      if (!campaignName) return null;

      const goals = {
        bb60sec2pages: VAG_ENO_V26_goalValue_(row, idx, VAG_ENO_V26_CONFIG.GOALS.bb60sec2pages),
        phone8835: VAG_ENO_V26_goalValue_(row, idx, VAG_ENO_V26_CONFIG.GOALS.phone8835),
        phone8865: VAG_ENO_V26_goalValue_(row, idx, VAG_ENO_V26_CONFIG.GOALS.phone8865),
        max: VAG_ENO_V26_goalValue_(row, idx, VAG_ENO_V26_CONFIG.GOALS.max),
        telegramSlider: VAG_ENO_V26_goalValue_(row, idx, VAG_ENO_V26_CONFIG.GOALS.telegramSlider),
        route: VAG_ENO_V26_goalValue_(row, idx, VAG_ENO_V26_CONFIG.GOALS.route),
        uisCalls: VAG_ENO_V26_goalValue_(row, idx, VAG_ENO_V26_CONFIG.GOALS.uisCalls),
        uisUniqueCalls: VAG_ENO_V26_goalValue_(row, idx, VAG_ENO_V26_CONFIG.GOALS.uisUniqueCalls),
        uisLeads: VAG_ENO_V26_goalValue_(row, idx, VAG_ENO_V26_CONFIG.GOALS.uisLeads)
      };

      const item = {
        campaignName: campaignName,
        campaignType: VAG_ENO_V26_getCampaignType_(campaignName),
        impressions: VAG_ENO_V26_toNumber_(row[idx['Impressions']]),
        clicks: VAG_ENO_V26_toNumber_(row[idx['Clicks']]),
        cost: VAG_ENO_V26_toNumber_(row[idx['Cost']]),
        avgEffectiveBid: idx['AvgCpc'] !== undefined ? VAG_ENO_V26_toNumber_(row[idx['AvgCpc']]) : 0,
        avgPosition: idx['AvgImpressionPosition'] !== undefined ? VAG_ENO_V26_toNumber_(row[idx['AvgImpressionPosition']]) : 0,
        avgTraffic: idx['AvgTrafficVolume'] !== undefined ? VAG_ENO_V26_toNumber_(row[idx['AvgTrafficVolume']]) : 0,
        bounceRate: idx['BounceRate'] !== undefined ? VAG_ENO_V26_toRate_(row[idx['BounceRate']]) : 0,
        depth: idx['AvgPageviews'] !== undefined ? VAG_ENO_V26_toNumber_(row[idx['AvgPageviews']]) : 0,
        goals: goals
      };

      return item;
    })
    .filter(item => item && (
      item.impressions > 0 ||
      item.clicks > 0 ||
      item.cost > 0 ||
      item.goals.bb60sec2pages > 0 ||
      item.goals.phone8835 > 0 ||
      item.goals.phone8865 > 0 ||
      item.goals.max > 0 ||
      item.goals.telegramSlider > 0 ||
      item.goals.route > 0 ||
      item.goals.uisCalls > 0 ||
      item.goals.uisUniqueCalls > 0 ||
      item.goals.uisLeads > 0
    ));
}

function VAG_ENO_V26_goalValue_(row, idx, goalId) {
  if (!goalId) return 0;

  const exact = 'Conversions_' + goalId + '_' + VAG_ENO_V26_CONFIG.ATTRIBUTION_MODEL;
  if (idx[exact] !== undefined) return VAG_ENO_V26_toNumber_(row[idx[exact]]);

  const prefix = 'Conversions_' + goalId + '_';
  for (const header in idx) {
    if (String(header).indexOf(prefix) === 0) {
      return VAG_ENO_V26_toNumber_(row[idx[header]]);
    }
  }

  return 0;
}


// =====================
// МЕТРИКА: ДОСТИЖЕНИЯ ЦЕЛЕЙ ПО КАМПАНИЯМ ДИРЕКТА
// =====================

function VAG_ENO_V26_overlayMetrikaGoalReaches_(rows, dateFrom, dateTo) {
  let metrikaMap = {};

  try {
    metrikaMap = VAG_ENO_V26_getMetrikaGoalReachesByCampaign_(dateFrom, dateTo);
  } catch (e) {
    const message = e && e.message ? e.message : String(e);
    VAG_ENO_V26_RUNTIME.metrikaWarning = VAG_ENO_V26_shortenWarning_(message);
    Logger.log('Метрика API недоступна, цели оставлены из Директа: ' + message);
    return;
  }

  const goalEntries = VAG_ENO_V26_getGoalEntries_();
  const existingKeys = {};

  rows.forEach(item => {
    const key = VAG_ENO_V26_campaignMatchKey_(item.campaignName);
    existingKeys[key] = true;

    if (!metrikaMap[key]) {
      goalEntries.forEach(entry => item.goals[entry.key] = 0);
      return;
    }

    goalEntries.forEach(entry => {
      item.goals[entry.key] = Number(metrikaMap[key].goals[entry.key] || 0);
    });
  });

  Object.keys(metrikaMap).forEach(key => {
    if (existingKeys[key]) return;

    const goals = VAG_ENO_V26_emptyGoals_();
    goalEntries.forEach(entry => goals[entry.key] = Number(metrikaMap[key].goals[entry.key] || 0));

    const hasGoals = goalEntries.some(entry => Number(goals[entry.key] || 0) > 0);
    if (!hasGoals) return;

    const campaignName = metrikaMap[key].campaignName;

    rows.push({
      campaignName: campaignName,
      campaignType: VAG_ENO_V26_getCampaignType_(campaignName),
      impressions: 0,
      clicks: 0,
      cost: 0,
      avgEffectiveBid: 0,
      avgPosition: 0,
      avgTraffic: 0,
      bounceRate: 0,
      depth: 0,
      goals: goals
    });
  });
}

function VAG_ENO_V26_getMetrikaGoalReachesByCampaign_(dateFrom, dateTo) {
  const token = VAG_ENO_V26_getMetrikaToken_();

  if (!token) {
    throw new Error('Не найден METRIKA_TOKEN или YANDEX_TOKEN в свойствах скрипта');
  }

  const goalEntries = VAG_ENO_V26_getGoalEntries_();
  const metrics = goalEntries
    .map(entry => 'ym:s:goal' + entry.id + 'reaches')
    .join(',');

  const params = {
    ids: VAG_ENO_V26_CONFIG.METRIKA_COUNTER_ID,
    date1: dateFrom,
    date2: dateTo,
    metrics: metrics,
    dimensions: 'ym:s:' + VAG_ENO_V26_CONFIG.METRIKA_ATTRIBUTION + 'DirectClickOrderName',
    accuracy: 'full',
    limit: '100000',
    lang: 'ru'
  };

  const url = 'https://api-metrika.yandex.net/stat/v1/data?' + VAG_ENO_V26_toQueryString_(params);

  const response = UrlFetchApp.fetch(url, {
    method: 'get',
    headers: {
      Authorization: 'OAuth ' + token
    },
    muteHttpExceptions: true
  });

  const code = response.getResponseCode();
  const text = response.getContentText();

  if (code !== 200) {
    if (code === 403) {
      throw new Error('Ошибка Метрики 403. Токен есть, но Метрика не отдала данные по счётчику ' + VAG_ENO_V26_CONFIG.METRIKA_COUNTER_ID + '. Проверь, что METRIKA_TOKEN от аккаунта с доступом к счётчику. Ответ API: ' + text);
    }
    throw new Error('Ошибка Метрики: ' + code + '\n' + text);
  }

  const json = JSON.parse(text);
  const result = {};

  (json.data || []).forEach(item => {
    const rawName = item.dimensions && item.dimensions[0] ? item.dimensions[0].name : '';
    const campaignName = VAG_ENO_V26_cleanMetrikaCampaignName_(rawName);
    const key = VAG_ENO_V26_campaignMatchKey_(campaignName);

    if (!key || key === 'none' || key === 'not set') return;

    if (!result[key]) {
      result[key] = {
        campaignName: campaignName,
        goals: VAG_ENO_V26_emptyGoals_()
      };
    }

    goalEntries.forEach((entry, index) => {
      result[key].goals[entry.key] += VAG_ENO_V26_toNumber_((item.metrics || [])[index]);
    });
  });

  return result;
}

function VAG_ENO_V26_getMetrikaToken_() {
  const props = PropertiesService.getScriptProperties();
  return props.getProperty('METRIKA_TOKEN') || props.getProperty('YANDEX_TOKEN');
}

function VAG_ENO_V26_shortenWarning_(message) {
  const text = String(message || '').replace(/\s+/g, ' ').trim();
  if (text.length <= 260) return text;
  return text.slice(0, 260) + '...';
}

function VAG_ENO_V26_getGoalEntries_() {
  return [
    { key: 'bb60sec2pages', id: VAG_ENO_V26_CONFIG.GOALS.bb60sec2pages },
    { key: 'phone8835', id: VAG_ENO_V26_CONFIG.GOALS.phone8835 },
    { key: 'phone8865', id: VAG_ENO_V26_CONFIG.GOALS.phone8865 },
    { key: 'max', id: VAG_ENO_V26_CONFIG.GOALS.max },
    { key: 'telegramSlider', id: VAG_ENO_V26_CONFIG.GOALS.telegramSlider },
    { key: 'route', id: VAG_ENO_V26_CONFIG.GOALS.route },
    { key: 'uisCalls', id: VAG_ENO_V26_CONFIG.GOALS.uisCalls },
    { key: 'uisUniqueCalls', id: VAG_ENO_V26_CONFIG.GOALS.uisUniqueCalls },
    { key: 'uisLeads', id: VAG_ENO_V26_CONFIG.GOALS.uisLeads }
  ];
}

function VAG_ENO_V26_emptyGoals_() {
  return {
    bb60sec2pages: 0,
    phone8835: 0,
    phone8865: 0,
    max: 0,
    telegramSlider: 0,
    route: 0,
    uisCalls: 0,
    uisUniqueCalls: 0,
    uisLeads: 0
  };
}

function VAG_ENO_V26_cleanMetrikaCampaignName_(name) {
  return String(name || '')
    .replace(/\s*\(N-\d+\)\s*$/i, '')
    .trim();
}

function VAG_ENO_V26_campaignMatchKey_(name) {
  return VAG_ENO_V26_cleanMetrikaCampaignName_(name)
    .replace(/\s+/g, ' ')
    .trim()
    .toLowerCase();
}

function VAG_ENO_V26_toQueryString_(params) {
  return Object.keys(params)
    .map(key => encodeURIComponent(key) + '=' + encodeURIComponent(params[key]))
    .join('&');
}

// =====================
// БЛОК ЕНО В ТАБЛИЦЕ
// =====================

function VAG_ENO_V26_getOrCreateBlock_(sheet, title) {
  const existingBlock = VAG_ENO_V26_findBlockByTitle_(sheet, title);
  if (existingBlock) return existingBlock;

  const template = VAG_ENO_V26_getTemplateBlock_(sheet);
  const newTopRow = VAG_ENO_V26_findLastFilledRowInColumnA_(sheet) + 4;
  const lastColumn = sheet.getLastColumn();
  const rowsToCopy = VAG_ENO_V26_CONFIG.TEMPLATE_COPY_ROWS;

  if (sheet.getMaxRows() < newTopRow + rowsToCopy + 20) {
    sheet.insertRowsAfter(sheet.getMaxRows(), newTopRow + rowsToCopy + 20 - sheet.getMaxRows());
  }

  sheet
    .getRange(template.topRow, 1, rowsToCopy, lastColumn)
    .copyTo(
      sheet.getRange(newTopRow, 1, rowsToCopy, lastColumn),
      SpreadsheetApp.CopyPasteType.PASTE_NORMAL,
      false
    );

  const headerRow = VAG_ENO_V26_findHeaderRow_(sheet, newTopRow);
  const totalRow = VAG_ENO_V26_findTotalRow_(sheet, headerRow);

  return { topRow: newTopRow, headerRow: headerRow, totalRow: totalRow };
}

function VAG_ENO_V26_getTemplateBlock_(sheet) {
  const topRow = VAG_ENO_V26_CONFIG.TEMPLATE_TOP_ROW;
  const headerRow = VAG_ENO_V26_findHeaderRow_(sheet, topRow);
  const totalRow = VAG_ENO_V26_findTotalRow_(sheet, headerRow);

  if (!headerRow || !totalRow) {
    throw new Error('Не нашла шаблон ЕНО. Проверь строку ' + topRow + ', шапку с РК и строку ИТОГО.');
  }

  return { topRow: topRow, headerRow: headerRow, totalRow: totalRow };
}

function VAG_ENO_V26_findBlockByTitle_(sheet, title) {
  const lastRow = sheet.getLastRow();
  const values = sheet.getRange(1, 1, lastRow, 1).getDisplayValues();

  for (let i = 0; i < values.length; i++) {
    if (String(values[i][0]).trim() === title) {
      const topRow = i + 1;
      const headerRow = VAG_ENO_V26_findHeaderRow_(sheet, topRow);
      const totalRow = VAG_ENO_V26_findTotalRow_(sheet, headerRow);
      return { topRow: topRow, headerRow: headerRow, totalRow: totalRow };
    }
  }

  return null;
}

function VAG_ENO_V26_findHeaderRow_(sheet, startRow) {
  const lastRow = sheet.getLastRow();
  const maxRow = Math.min(lastRow, startRow + 50);

  for (let row = startRow; row <= maxRow; row++) {
    const values = sheet.getRange(row, 1, 1, Math.min(sheet.getLastColumn(), 40)).getDisplayValues()[0];
    const normalized = values.map(value => VAG_ENO_V26_normalize_(value));

    const hasCampaign = normalized.indexOf('рк') !== -1;
    const hasImpressions = normalized.some(value => value.indexOf('показы') !== -1);

    if (hasCampaign && hasImpressions) return row;
  }

  return 0;
}

function VAG_ENO_V26_findTotalRow_(sheet, startRow) {
  const lastRow = sheet.getLastRow();
  const maxRow = Math.min(lastRow, startRow + 120);

  for (let row = startRow + 1; row <= maxRow; row++) {
    const values = sheet.getRange(row, 1, 1, Math.min(sheet.getLastColumn(), 10)).getDisplayValues()[0];
    if (values.some(value => VAG_ENO_V26_normalize_(value) === 'итого')) return row;
  }

  return 0;
}

function VAG_ENO_V26_findLastFilledRowInColumnA_(sheet) {
  const lastRow = sheet.getLastRow();
  const values = sheet.getRange(1, 1, lastRow, 1).getDisplayValues();

  for (let i = values.length - 1; i >= 0; i--) {
    if (String(values[i][0]).trim() !== '') return i + 1;
  }

  return VAG_ENO_V26_CONFIG.TEMPLATE_TOP_ROW + VAG_ENO_V26_CONFIG.TEMPLATE_COPY_ROWS;
}

function VAG_ENO_V26_fitRowsCount_(sheet, block, rowsCount) {
  const dataStartRow = block.headerRow + 1;
  const neededRows = Math.max(rowsCount, 1);
  let existingRows = block.totalRow - dataStartRow;
  const lastColumn = sheet.getLastColumn();

  if (existingRows < neededRows) {
    const rowsToAdd = neededRows - existingRows;
    sheet.insertRowsBefore(block.totalRow, rowsToAdd);
    sheet
      .getRange(dataStartRow, 1, 1, lastColumn)
      .copyTo(
        sheet.getRange(block.totalRow, 1, rowsToAdd, lastColumn),
        SpreadsheetApp.CopyPasteType.PASTE_NORMAL,
        false
      );
    block.totalRow += rowsToAdd;
    existingRows = neededRows;
  }

  if (existingRows > neededRows) {
    const rowsToDelete = existingRows - neededRows;
    sheet.deleteRows(dataStartRow + neededRows, rowsToDelete);
    block.totalRow -= rowsToDelete;
  }
}

function VAG_ENO_V26_setTitle_(sheet, topRow, title) {
  sheet.getRange(topRow, 1).setValue(title);
}

function VAG_ENO_V26_applyGoalHeaders_(sheet, headerRow) {
  const idRow = headerRow - 1;

  sheet.getRange(idRow, 13).setValue(Number(VAG_ENO_V26_CONFIG.GOALS.bb60sec2pages));
  sheet.getRange(idRow, 14).setValue(Number(VAG_ENO_V26_CONFIG.GOALS.phone8835));
  sheet.getRange(idRow, 15).setValue(Number(VAG_ENO_V26_CONFIG.GOALS.phone8865));
  sheet.getRange(idRow, 16).setValue(Number(VAG_ENO_V26_CONFIG.GOALS.max));
  sheet.getRange(idRow, 17).setValue(Number(VAG_ENO_V26_CONFIG.GOALS.telegramSlider));
  sheet.getRange(idRow, 18).setValue(Number(VAG_ENO_V26_CONFIG.GOALS.route));
  sheet.getRange(idRow, 19).setValue(Number(VAG_ENO_V26_CONFIG.GOALS.uisCalls));
  sheet.getRange(idRow, 20).setValue(Number(VAG_ENO_V26_CONFIG.GOALS.uisUniqueCalls));
  sheet.getRange(idRow, 21).setValue(Number(VAG_ENO_V26_CONFIG.GOALS.uisLeads));

  sheet.getRange(headerRow, 13).setValue('B-B_60сек+2стр (' + VAG_ENO_V26_CONFIG.GOALS.bb60sec2pages + ')');
  sheet.getRange(headerRow, 14).setValue('Клик на 79110008835 (' + VAG_ENO_V26_CONFIG.GOALS.phone8835 + ')');
  sheet.getRange(headerRow, 15).setValue('Клик на 79110008865 (' + VAG_ENO_V26_CONFIG.GOALS.phone8865 + ')');
  sheet.getRange(headerRow, 16).setValue('Клик на Макс (' + VAG_ENO_V26_CONFIG.GOALS.max + ')');
  sheet.getRange(headerRow, 17).setValue('Успешно пройден слайдер для перехода в TG (' + VAG_ENO_V26_CONFIG.GOALS.telegramSlider + ')');
  sheet.getRange(headerRow, 18).setValue('Клик по кнопке "Маршрут" (' + VAG_ENO_V26_CONFIG.GOALS.route + ')');
  sheet.getRange(headerRow, 19).setValue('Звонки Юис (' + VAG_ENO_V26_CONFIG.GOALS.uisCalls + ')');
  sheet.getRange(headerRow, 20).setValue('Звонки_Уник - ЮИС (' + VAG_ENO_V26_CONFIG.GOALS.uisUniqueCalls + ')');
  sheet.getRange(headerRow, 21).setValue('Заявки Юис (' + VAG_ENO_V26_CONFIG.GOALS.uisLeads + ')');
  sheet.getRange(headerRow, 22).setValue('Факт сумма конверсий');
}

// =====================
// КОЛОНКИ И ЗАПОЛНЕНИЕ
// =====================

function VAG_ENO_V26_getHeaderMap_(sheet, headerRow) {
  const headers = sheet.getRange(headerRow, 1, 1, sheet.getLastColumn()).getDisplayValues()[0];
  const map = {};

  headers.forEach((header, index) => {
    const key = VAG_ENO_V26_normalize_(header);
    if (key) map[key] = index + 1;
  });

  return {
    rk: VAG_ENO_V26_findExactCol_(map, 'рк'),
    type: VAG_ENO_V26_findContainsCol_(map, ['тип рк']),
    impressions: VAG_ENO_V26_findContainsCol_(map, ['показы']),
    clicks: VAG_ENO_V26_findContainsCol_(map, ['клики']),
    cost: VAG_ENO_V26_findCostCol_(map),
    avgClickCost: VAG_ENO_V26_findContainsCol_(map, ['ср. ставка за клик', 'cpc']),
    avgPosition: VAG_ENO_V26_findContainsCol_(map, ['ср. позиция показов']),
    avgTraffic: VAG_ENO_V26_findContainsCol_(map, ['ср. объем трафика', 'ср. объём трафика']),
    bounce: VAG_ENO_V26_findContainsCol_(map, ['отказы']),
    depth: VAG_ENO_V26_findContainsCol_(map, ['глубина']),
    bb60sec2pages: 13,
    phone8835: 14,
    phone8865: 15,
    max: 16,
    telegramSlider: 17,
    route: 18,
    uisCalls: 19,
    uisUniqueCalls: 20,
    uisLeads: 21,
    factConversions: 22,
    cr: 23,
    cpa: 24
  };
}

function VAG_ENO_V26_findExactCol_(map, header) {
  return map[VAG_ENO_V26_normalize_(header)] || 0;
}

function VAG_ENO_V26_findContainsCol_(map, parts) {
  const normalizedParts = parts.map(part => VAG_ENO_V26_normalize_(part));

  for (const key in map) {
    for (let i = 0; i < normalizedParts.length; i++) {
      if (key.indexOf(normalizedParts[i]) !== -1) return map[key];
    }
  }

  return 0;
}

function VAG_ENO_V26_findCostCol_(map) {
  for (const key in map) {
    if (key.indexOf('расход') !== -1 && key.indexOf('ндс') !== -1) return map[key];
  }
  return 0;
}

function VAG_ENO_V26_clearRawCells_(sheet, headerMap, startRow, rowsCount) {
  const columns = [
    headerMap.rk,
    headerMap.type,
    headerMap.impressions,
    headerMap.clicks,
    headerMap.cost,
    headerMap.avgClickCost,
    headerMap.avgPosition,
    headerMap.avgTraffic,
    headerMap.bounce,
    headerMap.depth,
    13, 14, 15, 16, 17, 18, 19, 20, 21
  ].filter(col => col > 0);

  columns.forEach(col => sheet.getRange(startRow, col, rowsCount, 1).clearContent());

  if (headerMap.bounce > 0) sheet.getRange(startRow, headerMap.bounce, rowsCount, 1).setNumberFormat('0.00%');
  if (headerMap.depth > 0) sheet.getRange(startRow, headerMap.depth, rowsCount, 1).setNumberFormat('0.00');
  if (headerMap.rk > 0) sheet.getRange(startRow, headerMap.rk, rowsCount, 1).setHorizontalAlignment('left');
}

function VAG_ENO_V26_fillCampaignRows_(sheet, headerMap, startRow, rows) {
  rows.forEach((item, index) => {
    const row = startRow + index;

    VAG_ENO_V26_setCell_(sheet, row, headerMap.rk, item.campaignName);
    VAG_ENO_V26_setCell_(sheet, row, headerMap.type, item.campaignType);
    VAG_ENO_V26_setCell_(sheet, row, headerMap.impressions, item.impressions);
    VAG_ENO_V26_setCell_(sheet, row, headerMap.clicks, item.clicks);
    VAG_ENO_V26_setCell_(sheet, row, headerMap.cost, item.cost);
    VAG_ENO_V26_setCell_(sheet, row, headerMap.avgClickCost, item.avgEffectiveBid);
    VAG_ENO_V26_setCell_(sheet, row, headerMap.avgPosition, item.avgPosition);
    VAG_ENO_V26_setCell_(sheet, row, headerMap.avgTraffic, item.avgTraffic);
    VAG_ENO_V26_setCell_(sheet, row, headerMap.bounce, item.bounceRate);
    VAG_ENO_V26_setCell_(sheet, row, headerMap.depth, item.depth);

    VAG_ENO_V26_setGoalCell_(sheet, row, 13, item.goals.bb60sec2pages);
    VAG_ENO_V26_setGoalCell_(sheet, row, 14, item.goals.phone8835);
    VAG_ENO_V26_setGoalCell_(sheet, row, 15, item.goals.phone8865);
    VAG_ENO_V26_setGoalCell_(sheet, row, 16, item.goals.max);
    VAG_ENO_V26_setGoalCell_(sheet, row, 17, item.goals.telegramSlider);
    VAG_ENO_V26_setGoalCell_(sheet, row, 18, item.goals.route);
    VAG_ENO_V26_setGoalCell_(sheet, row, 19, item.goals.uisCalls);
    VAG_ENO_V26_setGoalCell_(sheet, row, 20, item.goals.uisUniqueCalls);
    VAG_ENO_V26_setGoalCell_(sheet, row, 21, item.goals.uisLeads);
  });
}

function VAG_ENO_V26_setFactFormulas_(sheet, headerMap, dataStartRow, totalRow) {
  const dataEndRow = totalRow - 1;
  if (dataEndRow < dataStartRow) return;

  for (let row = dataStartRow; row <= dataEndRow; row++) {
    sheet.getRange(row, 22).setFormula('=SUM(T' + row + ':U' + row + ')');
    sheet.getRange(row, 23).setFormula('=IFERROR(V' + row + '/D' + row + ';0)');
    sheet.getRange(row, 24).setFormula('=IFERROR(G' + row + '/V' + row + ';0)');
  }

  const sumCols = [3, 4, 7, 13, 18, 19, 20, 21, 22];
  sumCols.forEach(col => {
    const letter = VAG_ENO_V26_columnToLetter_(col);
    sheet.getRange(totalRow, col).setFormula('=SUBTOTAL(9;' + letter + dataStartRow + ':' + letter + dataEndRow + ')');
  });

  sheet.getRange(totalRow, 5).setFormula('=IFERROR(D' + totalRow + '/C' + totalRow + ';0)');
  sheet.getRange(totalRow, 6).setFormula('=IFERROR(G' + totalRow + '/D' + totalRow + ';0)');
  sheet.getRange(totalRow, 9).setFormula('=AVERAGE(I' + dataStartRow + ':I' + dataEndRow + ')');
  sheet.getRange(totalRow, 10).setFormula('=AVERAGE(J' + dataStartRow + ':J' + dataEndRow + ')');

  if (headerMap.bounce > 0) {
    const bounceLetter = VAG_ENO_V26_columnToLetter_(headerMap.bounce);
    sheet.getRange(totalRow, headerMap.bounce)
      .setFormula('=AVERAGE(' + bounceLetter + dataStartRow + ':' + bounceLetter + dataEndRow + ')')
      .setNumberFormat('0.00%');
  }

  sheet.getRange(totalRow, 23).setFormula('=IFERROR(V' + totalRow + '/D' + totalRow + ';0)');
  sheet.getRange(totalRow, 24).setFormula('=IFERROR(G' + totalRow + '/V' + totalRow + ';0)');
}

function VAG_ENO_V26_setCell_(sheet, row, col, value) {
  if (col > 0) sheet.getRange(row, col).setValue(value);
}

function VAG_ENO_V26_setGoalCell_(sheet, row, col, value) {
  if (col > 0) {
    const number = Number(value || 0);
    sheet.getRange(row, col).setValue(number > 0 ? number : '-');
  }
}

function VAG_ENO_V26_cleanFormattingAfterFill_(sheet, headerMap, startRow, rowsCount, rows) {
  const lastCol = 24;

  sheet.getRange(startRow, 1, rowsCount, lastCol)
    .setBackground('#ffffff')
    .setFontWeight('normal');

  if (headerMap.bounce > 0) sheet.getRange(startRow, headerMap.bounce, rowsCount, 1).setNumberFormat('0.00%');
  if (headerMap.depth > 0) sheet.getRange(startRow, headerMap.depth, rowsCount, 1).setNumberFormat('0.00');
  if (headerMap.rk > 0) sheet.getRange(startRow, headerMap.rk, rowsCount, 1).setHorizontalAlignment('left');

  rows.forEach((item, index) => VAG_ENO_V26_highlightCampaignName_(sheet, headerMap, startRow + index, item));
}

function VAG_ENO_V26_highlightCampaignName_(sheet, headerMap, row, item) {
  if (headerMap.rk <= 0) return;

  const kpiConversions = Number(item.goals.uisUniqueCalls || 0) + Number(item.goals.uisLeads || 0);
  const cost = Number(item.cost || 0);
  const clicks = Number(item.clicks || 0);
  let color = '#ffffff';

  if (kpiConversions >= 1) {
    const cpa = cost / kpiConversions;
    if (cpa <= VAG_ENO_V26_CONFIG.GOOD_CPA_LIMIT) color = '#d9ead3';
    else if (cpa <= VAG_ENO_V26_CONFIG.MIDDLE_CPA_LIMIT) color = '#fff2cc';
    else color = '#f4cccc';
  } else {
    if (cost >= VAG_ENO_V26_CONFIG.RED_COST_LIMIT) color = '#f4cccc';
    else if (cost >= VAG_ENO_V26_CONFIG.YELLOW_COST_LIMIT || clicks >= VAG_ENO_V26_CONFIG.YELLOW_CLICKS_LIMIT) color = '#fff2cc';
  }

  sheet.getRange(row, headerMap.rk)
    .setBackground(color)
    .setHorizontalAlignment('left')
    .setFontWeight('normal');
}

// =====================
// КОММЕНТАРИЙ
// =====================

function VAG_ENO_V26_buildComment_(rows, dateFrom, dateTo) {
  const total = VAG_ENO_V26_getTotals_(rows);
  const kpiConversions = total.goals.uisUniqueCalls + total.goals.uisLeads;
  const cpa = kpiConversions > 0 ? total.cost / kpiConversions : 0;
  const cr = total.clicks > 0 ? kpiConversions / total.clicks * 100 : 0;
  const plan = VAG_ENO_V26_getPlan_(dateFrom, dateTo);
  const budgetPercent = plan > 0 ? total.cost / plan * 100 : 0;
  const byType = VAG_ENO_V26_groupByType_(rows);

  const lines = [];
  lines.push('Комментарии к отчёту за период с ' + VAG_ENO_V26_formatDateRu_(dateFrom) + ' по ' + VAG_ENO_V26_formatDateRu_(dateTo));
  lines.push('');
  lines.push('Бюджет:');
  lines.push('Всего потрачено за неделю ' + VAG_ENO_V26_rub_(total.cost) + '. Идём по бюджету на ' + VAG_ENO_V26_percent_(budgetPercent) + ' от плана на период.');
  lines.push('');
  lines.push('Трафик:');
  lines.push('Показы — ' + VAG_ENO_V26_num_(total.impressions));
  lines.push('Клики — ' + VAG_ENO_V26_num_(total.clicks));
  lines.push('Средний CPC — ' + VAG_ENO_V26_rub_(total.clicks > 0 ? total.cost / total.clicks : 0));
  lines.push('');
  lines.push('Достижения целей в Метрике:');
  lines.push('KPI-конверсий получено — ' + kpiConversions + ', CPA — ' + VAG_ENO_V26_rub_(cpa) + ', CR — ' + VAG_ENO_V26_percent_(cr));
  lines.push('B-B_60сек+2стр (555861179) — ' + total.goals.bb60sec2pages);
  lines.push('Клик на 79110008835 (511787501) — ' + total.goals.phone8835);
  lines.push('Клик на 79110008865 (511788019) — ' + total.goals.phone8865);
  lines.push('Клик на Макс (511788150) — ' + total.goals.max);
  lines.push('Успешно пройден слайдер для перехода в TG (511791937) — ' + total.goals.telegramSlider);
  lines.push('Клик по кнопке "Маршрут" (547214611) — ' + total.goals.route);
  lines.push('Звонки Юис (517626340) — ' + total.goals.uisCalls);
  lines.push('Звонки_Уник - ЮИС (562354501) — ' + total.goals.uisUniqueCalls);
  lines.push('Заявки Юис (517626592) — ' + total.goals.uisLeads);
  lines.push('');
  lines.push('По типам РК:');

  ['Поиск', 'Сети', 'Мастер', 'Товарная', 'Карты', 'Услуги', 'Поиск/Сети', 'н/а'].forEach(type => {
    const item = byType[type];
    if (!item) return;

    const ctr = item.impressions > 0 ? item.clicks / item.impressions * 100 : 0;
    const cpc = item.clicks > 0 ? item.cost / item.clicks : 0;
    const typeCr = item.clicks > 0 ? item.kpiConversions / item.clicks * 100 : 0;
    const typeCpa = item.kpiConversions > 0 ? item.cost / item.kpiConversions : 0;

    lines.push(
      type + ' — ' +
      VAG_ENO_V26_num_(item.clicks) + ' кликов, расход ' +
      VAG_ENO_V26_rub_(item.cost) + ', CTR — ' +
      VAG_ENO_V26_percent_(ctr) + ', CPC — ' +
      VAG_ENO_V26_rub_(cpc) + ', KPI-конверсий — ' +
      item.kpiConversions + ', CR — ' +
      VAG_ENO_V26_percent_(typeCr) + ', CPA — ' +
      VAG_ENO_V26_rub_(typeCpa)
    );
  });

  lines.push('');
  lines.push('Зона внимания:');
  VAG_ENO_V26_buildAttentionLines_(rows, cpa).forEach(line => lines.push(line));
  lines.push('');
  lines.push('Вывод:');
  lines.push('Усиливаем кампании, которые дают уникальные звонки и заявки. Кампании с расходом без KPI-конверсий проверяем по запросам, объявлениям, площадкам, отказам и посадочным страницам.');

  return lines.join('\n');
}

function VAG_ENO_V26_prepareCommentArea_(sheet, block) {
  const lastColumn = Math.min(sheet.getLastColumn(), 24); // A:X
  const clearStartRow = block.totalRow + 2;
  const copiedBlockEndRow = block.topRow + VAG_ENO_V26_CONFIG.TEMPLATE_COPY_ROWS - 1;
  const clearEndRow = Math.max(copiedBlockEndRow, clearStartRow + 45);

  if (sheet.getMaxRows() < clearEndRow + 5) {
    sheet.insertRowsAfter(sheet.getMaxRows(), clearEndRow + 5 - sheet.getMaxRows());
  }

  sheet.getRange(clearStartRow, 1, clearEndRow - clearStartRow + 1, lastColumn)
    .clearContent()
    .breakApart()
    .setBackground('#ffffff')
    .setFontWeight('normal')
    .setFontSize(10)
    .setFontColor('#000000')
    .setWrap(false)
    .setBorder(false, false, false, false, false, false);

  return clearStartRow;
}

function VAG_ENO_V26_writeComment_(sheet, startRow, comment) {
  const lastColumn = Math.min(sheet.getLastColumn(), 24); // A:X
  const rawLines = comment.split('\n').map(line => String(line).trim());

  // Оставляем пустые строки между смысловыми блоками, но убираем пустой хвост.
  while (rawLines.length && rawLines[rawLines.length - 1] === '') {
    rawLines.pop();
  }

  if (!rawLines.length) return;

  const neededRows = rawLines.length;
  if (sheet.getMaxRows() < startRow + neededRows + 2) {
    sheet.insertRowsAfter(sheet.getMaxRows(), startRow + neededRows + 2 - sheet.getMaxRows());
  }

  // Комментарий теперь пишется строго в колонку A.
  // Объединений нет, поэтому ошибки с закрепленными/незакрепленными колонками не будет.
  const fullRange = sheet.getRange(startRow, 1, neededRows, lastColumn);
  fullRange
    .clearContent()
    .breakApart()
    .setBackground('#ffffff')
    .setFontWeight('normal')
    .setFontSize(10)
    .setFontColor('#000000')
    .setWrap(true)
    .setVerticalAlignment('middle')
    .setHorizontalAlignment('left')
    .setBorder(false, false, false, false, false, false);

  const commentColumnRange = sheet.getRange(startRow, 1, neededRows, 1);
  commentColumnRange
    .setValues(rawLines.map(line => [line]))
    .setWrap(true)
    .setVerticalAlignment('middle')
    .setHorizontalAlignment('left')
    .setBorder(true, true, true, true, true, true, '#d9d9d9', SpreadsheetApp.BorderStyle.SOLID);

  // Делаем колонку A достаточно широкой для читаемого комментария, но не уменьшаем текущую ширину.
  const currentWidth = sheet.getColumnWidth(1);
  if (currentWidth < 520) {
    sheet.setColumnWidth(1, 520);
  }

  rawLines.forEach((line, index) => {
    const row = startRow + index;
    const cell = sheet.getRange(row, 1);
    const normalizedLine = String(line || '').trim();

    if (normalizedLine === '') {
      cell
        .setBackground('#ffffff')
        .setBorder(false, false, false, false, false, false);
      sheet.setRowHeight(row, 8);
      return;
    }

    const isTitle = index === 0 || normalizedLine.indexOf('Комментарии к отчёту') === 0;
    const isSectionHeader = normalizedLine.endsWith(':') || normalizedLine.indexOf('Вывод:') === 0;

    if (isTitle) {
      cell
        .setBackground('#d9ead3')
        .setFontWeight('bold')
        .setFontSize(11);
      return;
    }

    if (isSectionHeader) {
      cell
        .setBackground('#cfe2f3')
        .setFontWeight('bold')
        .setFontSize(10);
      return;
    }

    cell
      .setBackground('#ffffff')
      .setFontWeight('normal')
      .setFontSize(10);
  });

  sheet.autoResizeRows(startRow, neededRows);
}

function VAG_ENO_V26_buildAttentionLines_(rows, avgCpa) {
  const noConversionCostLimit = Math.max(1000, Number(avgCpa || 0) * 0.5);

  const risks = rows
    .map(row => {
      const conversions = Number(row.goals.uisUniqueCalls || 0) + Number(row.goals.uisLeads || 0);
      const cpa = conversions > 0 ? Number(row.cost || 0) / conversions : 0;
      return {
        name: row.campaignName,
        cost: Number(row.cost || 0),
        clicks: Number(row.clicks || 0),
        conversions: conversions,
        cpa: cpa
      };
    })
    .filter(item => {
      const hasSpendNoConversions = item.cost >= noConversionCostLimit && item.conversions === 0;
      const hasHighCpa = item.conversions > 0 && avgCpa > 0 && item.cpa > avgCpa * 1.5;
      return hasSpendNoConversions || hasHighCpa;
    })
    .sort((a, b) => {
      if (a.conversions === 0 && b.conversions > 0) return -1;
      if (a.conversions > 0 && b.conversions === 0) return 1;
      return b.cost - a.cost;
    })
    .slice(0, 5);

  if (!risks.length) return ['Критичных кампаний с заметным расходом без KPI-конверсий не выявлено.'];

  return risks.map(item => {
    if (item.conversions === 0) {
      return item.name + ' — расход ' + VAG_ENO_V26_rub_(item.cost) + ', KPI-конверсий нет. Проверить запросы, объявления, площадки и посадочную страницу.';
    }
    return item.name + ' — CPA ' + VAG_ENO_V26_rub_(item.cpa) + ', выше среднего по отчёту.';
  });
}

function VAG_ENO_V26_getTotals_(rows) {
  const total = {
    impressions: 0,
    clicks: 0,
    cost: 0,
    goals: {
      bb60sec2pages: 0,
      phone8835: 0,
      phone8865: 0,
      max: 0,
      telegramSlider: 0,
      route: 0,
      uisCalls: 0,
      uisUniqueCalls: 0,
      uisLeads: 0
    }
  };

  rows.forEach(row => {
    total.impressions += Number(row.impressions || 0);
    total.clicks += Number(row.clicks || 0);
    total.cost += Number(row.cost || 0);
    total.goals.bb60sec2pages += Number(row.goals.bb60sec2pages || 0);
    total.goals.phone8835 += Number(row.goals.phone8835 || 0);
    total.goals.phone8865 += Number(row.goals.phone8865 || 0);
    total.goals.max += Number(row.goals.max || 0);
    total.goals.telegramSlider += Number(row.goals.telegramSlider || 0);
    total.goals.route += Number(row.goals.route || 0);
    total.goals.uisCalls += Number(row.goals.uisCalls || 0);
    total.goals.uisUniqueCalls += Number(row.goals.uisUniqueCalls || 0);
    total.goals.uisLeads += Number(row.goals.uisLeads || 0);
  });

  return total;
}

function VAG_ENO_V26_groupByType_(rows) {
  const result = {};

  rows.forEach(row => {
    const type = row.campaignType || 'н/а';
    if (!result[type]) {
      result[type] = { impressions: 0, clicks: 0, cost: 0, kpiConversions: 0 };
    }
    result[type].impressions += Number(row.impressions || 0);
    result[type].clicks += Number(row.clicks || 0);
    result[type].cost += Number(row.cost || 0);
    result[type].kpiConversions += Number(row.goals.uisUniqueCalls || 0) + Number(row.goals.uisLeads || 0);
  });

  return result;
}

// =====================
// ДАТЫ И ФОРМАТЫ
// =====================

function VAG_ENO_V26_getCurrentWeekPeriod_() {
  const today = VAG_ENO_V26_parseApiDate_(Utilities.formatDate(new Date(), VAG_ENO_V26_CONFIG.TIMEZONE, 'yyyy-MM-dd'));
  const day = today.getDay();
  const diffToMonday = day === 0 ? 6 : day - 1;
  const monday = new Date(today);
  monday.setDate(today.getDate() - diffToMonday);

  return {
    dateFrom: Utilities.formatDate(monday, VAG_ENO_V26_CONFIG.TIMEZONE, 'yyyy-MM-dd'),
    dateTo: Utilities.formatDate(today, VAG_ENO_V26_CONFIG.TIMEZONE, 'yyyy-MM-dd')
  };
}

function VAG_ENO_V26_getPreviousFullWeekPeriod_() {
  const today = VAG_ENO_V26_parseApiDate_(Utilities.formatDate(new Date(), VAG_ENO_V26_CONFIG.TIMEZONE, 'yyyy-MM-dd'));
  const day = today.getDay();
  const diffToMonday = day === 0 ? 6 : day - 1;

  const currentMonday = new Date(today);
  currentMonday.setDate(today.getDate() - diffToMonday);

  const previousMonday = new Date(currentMonday);
  previousMonday.setDate(currentMonday.getDate() - 7);

  const previousSunday = new Date(currentMonday);
  previousSunday.setDate(currentMonday.getDate() - 1);

  return {
    dateFrom: Utilities.formatDate(previousMonday, VAG_ENO_V26_CONFIG.TIMEZONE, 'yyyy-MM-dd'),
    dateTo: Utilities.formatDate(previousSunday, VAG_ENO_V26_CONFIG.TIMEZONE, 'yyyy-MM-dd')
  };
}

function VAG_ENO_V26_buildTitle_(dateFrom, dateTo) {
  return 'Отчет за неделю ' + VAG_ENO_V26_shortDate_(dateFrom) + ' - ' + VAG_ENO_V26_shortDate_(dateTo);
}

function VAG_ENO_V26_shortDate_(apiDate) {
  const parts = apiDate.split('-');
  return parts[2] + '.' + parts[1];
}

function VAG_ENO_V26_formatDateRu_(apiDate) {
  const parts = apiDate.split('-');
  return parts[2] + '.' + parts[1] + '.';
}

function VAG_ENO_V26_parseApiDate_(apiDate) {
  const parts = String(apiDate).split('-').map(Number);
  return new Date(parts[0], parts[1] - 1, parts[2]);
}

function VAG_ENO_V26_getPlan_(dateFrom, dateTo) {
  const start = VAG_ENO_V26_parseApiDate_(dateFrom);
  const end = VAG_ENO_V26_parseApiDate_(dateTo);
  const days = Math.round((end - start) / (24 * 60 * 60 * 1000)) + 1;
  const daysInMonth = new Date(start.getFullYear(), start.getMonth() + 1, 0).getDate();
  return VAG_ENO_V26_CONFIG.MONTH_BUDGET / daysInMonth * days;
}

function VAG_ENO_V26_getCampaignType_(campaignName) {
  const name = String(campaignName || '').toLowerCase();

  if (name.includes('m-k') || name.includes('мастер')) return 'Мастер';
  if (name.includes('карт')) return 'Карты';
  if (name.includes('услуг') || name.includes('galereya')) return 'Услуги';
  if (name.includes('сет') || name.includes('_s_') || name.includes('rsya') || name.includes('рся')) return 'Сети';
  if (name.includes('товар')) return 'Товарная';
  if (name.includes('поиск') || name.includes('_p_') || name.includes('brand')) return 'Поиск';

  return 'н/а';
}

function VAG_ENO_V26_toNumber_(value) {
  if (value === null || value === undefined || value === '' || value === '-' || value === '--') return 0;
  const text = String(value).replace(/\s/g, '').replace(',', '.').replace(/[^\d.-]/g, '');
  const number = Number(text);
  return isNaN(number) ? 0 : number;
}

function VAG_ENO_V26_toRate_(value) {
  const number = VAG_ENO_V26_toNumber_(value);
  if (!number) return 0;

  // В отчёте Директа BounceRate приходит как процентное число: 26.92.
  // В Google Sheets при формате 0.00% нужно записывать долю: 0.2692.
  return number > 1 ? number / 100 : number;
}

function VAG_ENO_V26_rub_(value) {
  return Number(value || 0).toLocaleString('ru-RU', { minimumFractionDigits: 2, maximumFractionDigits: 2 }) + ' руб.';
}

function VAG_ENO_V26_num_(value) {
  return Number(value || 0).toLocaleString('ru-RU');
}

function VAG_ENO_V26_percent_(value) {
  return Number(value || 0).toLocaleString('ru-RU', { minimumFractionDigits: 2, maximumFractionDigits: 2 }) + '%';
}

function VAG_ENO_V26_normalize_(value) {
  return String(value || '')
    .toLowerCase()
    .replace(/ё/g, 'е')
    .replace(/[«»"']/g, '')
    .replace(/\s+/g, ' ')
    .trim();
}

function VAG_ENO_V26_columnToLetter_(column) {
  let temp = '';
  let letter = '';

  while (column > 0) {
    temp = (column - 1) % 26;
    letter = String.fromCharCode(temp + 65) + letter;
    column = (column - temp - 1) / 26;
  }

  return letter;
}

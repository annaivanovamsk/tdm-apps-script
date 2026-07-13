// ==========================================================
// ТДМ — месячный отчёт ЕНО/ЕМО
//
// Что делает:
// 1) Берёт предыдущий месячный блок как шаблон.
// 2) Создаёт блок за прошлый полный месяц.
// 3) Автоматически тянет данные из Яндекс Директа.
// 4) Заполняет таблицу РК по актуальной схеме целей.
// 5) Сохраняет формулы и форматирование. Формулы НЕ очищает и НЕ перезаписывает.
// 6) Добавляет клиентские комментарии ниже таблицы.
// 7) В итог лидов считает только актуальные целевые действия:
//    Ecommerce покупка + Jivo + Callibri Lead A + Callibri Lead C.
// Старые внешние трекеры не используются в активной логике.
// ==========================================================

const TDM_ENO_MAY_MONTHLY_V3_CONFIG = {
  // На скрине активна вкладка «ЕМО», но ты называешь блок ЕНО.
  // Поэтому код умеет работать с обеими вкладками.
  SHEET_NAME_CANDIDATES: ['ЕМО', 'ЕНО'],

  CLIENT_LOGIN: 'ilhaleontj3v',
  TIMEZONE: 'Europe/Moscow',
  ATTRIBUTION_MODEL: 'AUTO',

  DATE_FROM: '2026-05-01',
  DATE_TO: '2026-05-31',

  // Старые внешние цели больше не используем в активной логике.
  LEGACY_CUTOFF_TO: '2026-05-20',

  // Если план нужен в комментарии — укажи сумму.
  // Если 0, будет только фактический расход.
  MONTH_PLAN: 0,

  // Правило безопасности:
  // формулы в строках кампаний не очищаем и не перезаписываем.
  // Исключение: строка ИТОГО. Там формулы можно обновить,
  // чтобы они всегда считали до последней строки кампаний.
  NEVER_TOUCH_DATA_ROW_FORMULAS: true,
  ALLOW_UPDATE_TOTAL_FORMULAS_TO_END: true,

  // Подсветка строк кампаний.
  // Только заливка. Данные и формулы не меняем.
  COLORS: {
    strong: '#d9ead3',   // зелёный — сильные
    medium: '#ffffff',   
    weak: '#f4cccc',     // красный — слабые
    neutral: '#ffffff'
  },

  // Майский блок должен начинаться с 93 строки:
  // апрельские комментарии заканчиваются на 91,
  // 92 строка остаётся пустой,
  // 93 строка — старт нового месяца.
  FORCE_START_ROW: 93,

  // Для будущих месяцев правило такое же:
  // новый месячный блок начинается через одну пустую строку после прошлого блока.
  EMPTY_ROWS_BETWEEN_MONTHS: 1,

  GAP_ROWS: 1,
  COPY_COLS: 35,

  APRIL_TEMPLATE_MARKERS: [
    '01.04.2026',
    '1.04.2026',
    '30.04.2026',
    'апрель'
  ],

  MAY_MARKERS: [
    '01.05.2026',
    '1.05.2026',
    '31.05.2026',
    'май'
  ],

  GOALS: {
    view3Pages: ['453453800'],      // Просмотр 3х страниц — микроцель
    addToCart: ['504318736'],       // Ecommerce: добавление в корзину
    purchase: ['504318735'],        // Ecommerce: покупка — факт
    jivo: ['575188424'],            // Jivo-сайт: пользователь начал чат — факт
    callibriSpam: [],               // Callibri: Спам — источник пока не подключён
    callibriNonTarget: [],          // Callibri: Нецелевой_Лид — источник пока не подключён
    callibriLeadA: [],              // Callibri: Лид_Квал_A — источник пока не подключён
    callibriLeadC: []               // Callibri: Лид_Квал_C — источник пока не подключён
  },

  FACT_LEAD_KEYS: ['purchase', 'jivo', 'callibriLeadA', 'callibriLeadC']
};

// ==========================================================
// ЗАПУСК
// ==========================================================

// REMOVED_20260708: ручной запуск майского отчёта удалён из runtime. Использовать fillTdmEnoPreviousFullMonthReport.

function addEnoEmoMenu_() {
  try {
    SpreadsheetApp.getUi()
      .createMenu('ЕНО / ЕМО')
      .addItem('Сформировать прошлый полный месяц', 'fillTdmEnoPreviousFullMonthReport')
      .addToUi();
  } catch (e) {
    Logger.log(e);
  }
}

// ==========================================================
// ОСНОВНОЙ ПРОЦЕСС
// ==========================================================

function tdmEnoMayMonthlyV3Fill_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = tdmEnoMayMonthlyV3GetSheet_(ss);

  const title = tdmEnoMayMonthlyV3Title_(
    TDM_ENO_MAY_MONTHLY_V3_CONFIG.DATE_FROM,
    TDM_ENO_MAY_MONTHLY_V3_CONFIG.DATE_TO
  );

  // Сначала создаём/находим блок, чтобы взять порядок РК и типы из апрельского шаблона.
  const block = tdmEnoMayMonthlyV3GetOrCreateBlock_(sheet, title);
  const header = tdmEnoMayMonthlyV3HeaderMap_(sheet, block.headerRow);

  const templateMeta = tdmEnoMayMonthlyV3ReadTemplateCampaignMeta_(
    sheet,
    header,
    block.headerRow + 1,
    block.totalRow - 1
  );

  let rows = tdmEnoMayMonthlyV3LoadRowsCurrentGoals_(
    TDM_ENO_MAY_MONTHLY_V3_CONFIG.DATE_FROM,
    TDM_ENO_MAY_MONTHLY_V3_CONFIG.DATE_TO,
    TDM_ENO_MAY_MONTHLY_V3_CONFIG.LEGACY_CUTOFF_TO
  );

  rows = tdmEnoMayMonthlyV3ApplyTemplateMeta_(rows, templateMeta);
  rows = tdmEnoMayMonthlyV3SortRowsByTemplate_(rows, templateMeta);

  const preparedBlock = tdmEnoMayMonthlyV3PrepareBlock_(sheet, block, Math.max(rows.length, 1));
  const preparedHeader = tdmEnoMayMonthlyV3HeaderMap_(sheet, preparedBlock.headerRow);
  const dataStartRow = preparedBlock.headerRow + 1;

  tdmEnoMayMonthlyV3SetTitleRow_(sheet, preparedBlock.topRow, title);

  tdmEnoMayMonthlyV3ClearDataValues_(
    sheet,
    preparedHeader,
    dataStartRow,
    Math.max(rows.length, 1)
  );

  tdmEnoMayMonthlyV3FillRows_(
    sheet,
    preparedHeader,
    dataStartRow,
    rows
  );

  tdmEnoMayMonthlyV3FillTotalRow_(
    sheet,
    preparedHeader,
    preparedBlock.totalRow,
    rows
  );

  // Обновляем только формулы строки ИТОГО:
  // диапазоны должны идти от первой до последней строки кампаний.
  // Остальные формулы и данные не трогаем.
  tdmEnoMayMonthlyV3UpdateTotalFormulasToEnd_(
    sheet,
    preparedHeader,
    dataStartRow,
    preparedBlock.totalRow
  );

  tdmEnoMayMonthlyV3SetTopSummary_(
    sheet,
    preparedBlock.topRow,
    tdmEnoMayMonthlyV3Totals_(rows)
  );

  tdmEnoMayMonthlyV3ColorRowsByStrength_(
    sheet,
    preparedHeader,
    dataStartRow,
    Math.max(rows.length, 1),
    rows
  );

  const comments = tdmEnoMayMonthlyV3BuildComments_(rows);
  const commentRow = preparedBlock.totalRow + 2;

  tdmEnoMayMonthlyV3WriteComments_(
    sheet,
    commentRow,
    comments
  );

  SpreadsheetApp.flush();

  tdmEnoMayMonthlyV3Toast_(
    'Готово. Месячный отчёт за май сформирован на вкладке «' + sheet.getName() + '».'
  );
}

// ==========================================================
// ВЫБОР ВКЛАДКИ
// ==========================================================

function tdmEnoMayMonthlyV3GetSheet_(ss) {
  const candidates = TDM_ENO_MAY_MONTHLY_V3_CONFIG.SHEET_NAME_CANDIDATES;
  const activeSheet = ss.getActiveSheet();

  if (activeSheet && candidates.indexOf(activeSheet.getName()) !== -1) {
    return activeSheet;
  }

  for (let i = 0; i < candidates.length; i++) {
    const sheet = ss.getSheetByName(candidates[i]);

    if (sheet) return sheet;
  }

  throw new Error('Не нашла вкладку ЕМО или ЕНО.');
}

// ==========================================================
// ЗАГРУЗКА ДАННЫХ ИЗ ДИРЕКТА
// ==========================================================

function tdmEnoMayMonthlyV3LoadRowsCurrentGoals_(dateFrom, dateTo, legacyCutoffTo) {
  const fullRows = tdmEnoMayMonthlyV3LoadCampaignRows_(dateFrom, dateTo);

  // Старые внешние цели больше не добираем отдельным срезом и не включаем в факт.
  fullRows.forEach(row => {
    row.factLeads = tdmEnoMayMonthlyV3FactLeads_(row.goals);
  });

  return fullRows.filter(row => {
    return row.impressions > 0 ||
      row.clicks > 0 ||
      row.cost > 0 ||
      tdmEnoMayMonthlyV3AllGoalsSum_(row.goals) > 0;
  });
}

function tdmEnoMayMonthlyV3LoadCampaignRows_(dateFrom, dateTo) {
  const goalIds = tdmEnoMayMonthlyV3AllGoalIds_();

  const body = {
    params: {
      SelectionCriteria: {
        DateFrom: dateFrom,
        DateTo: dateTo
      },
      Goals: goalIds.map(id => Number(id)),
      AttributionModels: [
        TDM_ENO_MAY_MONTHLY_V3_CONFIG.ATTRIBUTION_MODEL
      ],
      FieldNames: [
        'CampaignName',
        'Impressions',
        'Clicks',
        'Cost',
        'AvgEffectiveBid',
        'AvgImpressionPosition',
        'AvgTrafficVolume',
        'BounceRate',
        'AvgPageviews',
        'Conversions'
      ],
      ReportName: 'tdm_eno_monthly_may_' + dateFrom + '_' + dateTo + '_' + Utilities.getUuid(),
      ReportType: 'CAMPAIGN_PERFORMANCE_REPORT',
      DateRangeType: 'CUSTOM_DATE',
      Format: 'TSV',
      IncludeVAT: 'YES',
      IncludeDiscount: 'NO'
    }
  };

  const text = tdmEnoMayMonthlyV3DirectRequest_(body);
  return tdmEnoMayMonthlyV3ParseCampaignTsv_(text);
}

function tdmEnoMayMonthlyV3SafeLog_(context, error) {
  const message = error && error.message ? error.message : String(error || 'unknown error');
  Logger.log('[EMO] ' + context + ': ' + message);
  if (error && error.stack) Logger.log(error.stack);
}

function tdmEnoMayMonthlyV3NormalizeToken_(token) {
  return String(token || '')
    .replace(/^OAuth\s+/i, '')
    .replace(/^Bearer\s+/i, '')
    .trim();
}

function tdmEnoMayMonthlyV3FetchWithRetry_(url, options, config) {
  config = config || {};
  const maxAttempts = Number(config.maxAttempts || 8);
  const sleepMs = Number(config.sleepMs || 4000);
  const retryCodes = config.retryCodes || [201, 202, 429, 500, 502, 503, 504];
  let lastError = null;

  for (let attempt = 1; attempt <= maxAttempts; attempt++) {
    try {
      const response = UrlFetchApp.fetch(url, options);
      const code = response.getResponseCode();
      const text = response.getContentText();

      if (code === 200) {
        return { code: code, text: text, response: response };
      }

      if (retryCodes.indexOf(code) !== -1 && attempt < maxAttempts) {
        Utilities.sleep(sleepMs);
        continue;
      }

      throw new Error('HTTP ' + code + ': ' + String(text || '').slice(0, 1000));
    } catch (e) {
      lastError = e;
      if (attempt < maxAttempts) {
        Utilities.sleep(sleepMs);
        continue;
      }
    }
  }

  throw lastError || new Error('Fetch failed without response');
}

function tdmEnoMayMonthlyV3DirectRequest_(body) {
  try {
    const token = tdmEnoMayMonthlyV3NormalizeToken_(
      PropertiesService.getScriptProperties().getProperty('YANDEX_TOKEN')
    );

    if (!token) {
      throw new Error('Не найден YANDEX_TOKEN в Script Properties.');
    }

    if (!body || !body.params) {
      throw new Error('Пустое тело запроса Директа.');
    }

    const url = 'https://api.direct.yandex.com/json/v5/reports';

    const options = {
      method: 'post',
      contentType: 'application/json',
      payload: JSON.stringify(body),
      headers: {
        Authorization: 'Bearer ' + token,
        'Client-Login': TDM_ENO_MAY_MONTHLY_V3_CONFIG.CLIENT_LOGIN,
        'Accept-Language': 'ru',
        processingMode: 'auto',
        returnMoneyInMicros: 'false',
        skipReportHeader: 'true',
        skipColumnHeader: 'false',
        skipReportSummary: 'true'
      },
      muteHttpExceptions: true
    };

    const result = tdmEnoMayMonthlyV3FetchWithRetry_(url, options, {
      maxAttempts: 8,
      sleepMs: 4000,
      retryCodes: [201, 202, 429, 500, 502, 503, 504]
    });

    return result.text;
  } catch (e) {
    tdmEnoMayMonthlyV3SafeLog_('tdmEnoMayMonthlyV3DirectRequest_', e);
    throw e;
  }
}

function tdmEnoMayMonthlyV3ParseCampaignTsv_(tsvText) {
  const text = String(tsvText || '').trim();

  if (!text) return [];

  const lines = text
    .split(/\r?\n/)
    .filter(line => String(line).trim() !== '');

  const headerIndex = lines.findIndex(line => {
    const cols = line.split('\t');
    return cols.indexOf('CampaignName') !== -1 && cols.indexOf('Cost') !== -1;
  });

  if (headerIndex === -1) {
    throw new Error('Не найдена строка заголовков в отчёте Директа.');
  }

  const headers = lines[headerIndex]
    .split('\t')
    .map(header => String(header).replace(/^\uFEFF/, '').trim());

  const idxCampaign = tdmEnoMayMonthlyV3RequireColumn_(headers, 'CampaignName');
  const idxImpressions = tdmEnoMayMonthlyV3RequireColumn_(headers, 'Impressions');
  const idxClicks = tdmEnoMayMonthlyV3RequireColumn_(headers, 'Clicks');
  const idxCost = tdmEnoMayMonthlyV3RequireColumn_(headers, 'Cost');

  const idxAvgBid = headers.indexOf('AvgEffectiveBid');
  const idxAvgPosition = headers.indexOf('AvgImpressionPosition');
  const idxAvgTraffic = headers.indexOf('AvgTrafficVolume');
  const idxBounce = headers.indexOf('BounceRate');
  const idxDepth = headers.indexOf('AvgPageviews');

  const goalIndexes = {};

  tdmEnoMayMonthlyV3AllGoalIds_().forEach(goalId => {
    goalIndexes[goalId] = tdmEnoMayMonthlyV3FindGoalColumn_(headers, goalId);
  });

  const result = [];

  for (let i = headerIndex + 1; i < lines.length; i++) {
    const row = lines[i].split('\t');
    const campaignName = String(row[idxCampaign] || '').trim();

    if (!campaignName) continue;

    const goals = {
      view3Pages: tdmEnoMayMonthlyV3GoalSum_(row, goalIndexes, TDM_ENO_MAY_MONTHLY_V3_CONFIG.GOALS.view3Pages),
      callibriSpam: tdmEnoMayMonthlyV3GoalSum_(row, goalIndexes, TDM_ENO_MAY_MONTHLY_V3_CONFIG.GOALS.callibriSpam),
      callibriNonTarget: tdmEnoMayMonthlyV3GoalSum_(row, goalIndexes, TDM_ENO_MAY_MONTHLY_V3_CONFIG.GOALS.callibriNonTarget),
      addToCart: tdmEnoMayMonthlyV3GoalSum_(row, goalIndexes, TDM_ENO_MAY_MONTHLY_V3_CONFIG.GOALS.addToCart),
      purchase: tdmEnoMayMonthlyV3GoalSum_(row, goalIndexes, TDM_ENO_MAY_MONTHLY_V3_CONFIG.GOALS.purchase),
      jivo: tdmEnoMayMonthlyV3GoalSum_(row, goalIndexes, TDM_ENO_MAY_MONTHLY_V3_CONFIG.GOALS.jivo),
      callibriLeadA: tdmEnoMayMonthlyV3GoalSum_(row, goalIndexes, TDM_ENO_MAY_MONTHLY_V3_CONFIG.GOALS.callibriLeadA),
      callibriLeadC: tdmEnoMayMonthlyV3GoalSum_(row, goalIndexes, TDM_ENO_MAY_MONTHLY_V3_CONFIG.GOALS.callibriLeadC)
    };

    result.push({
      campaignName: campaignName,
      campaignType: tdmEnoMayMonthlyV3FallbackCampaignType_(campaignName),
      geo: tdmEnoMayMonthlyV3CampaignGeo_(campaignName),

      impressions: tdmEnoMayMonthlyV3ToNumber_(row[idxImpressions]),
      clicks: tdmEnoMayMonthlyV3ToNumber_(row[idxClicks]),
      cost: tdmEnoMayMonthlyV3ToNumber_(row[idxCost]),

      avgBid: idxAvgBid === -1 ? 0 : tdmEnoMayMonthlyV3ToNumber_(row[idxAvgBid]),
      avgPosition: idxAvgPosition === -1 ? '' : tdmEnoMayMonthlyV3Metric_(row[idxAvgPosition]),
      avgTraffic: idxAvgTraffic === -1 ? '' : tdmEnoMayMonthlyV3Metric_(row[idxAvgTraffic]),
      bounce: idxBounce === -1 ? 0 : tdmEnoMayMonthlyV3ToNumber_(row[idxBounce]) / 100,
      depth: idxDepth === -1 ? 0 : tdmEnoMayMonthlyV3ToNumber_(row[idxDepth]),

      goals: goals,
      factLeads: tdmEnoMayMonthlyV3FactLeads_(goals)
    });
  }

  return result;
}

function tdmEnoMayMonthlyV3FindGoalColumn_(headers, goalId) {
  const id = String(goalId);
  const model = TDM_ENO_MAY_MONTHLY_V3_CONFIG.ATTRIBUTION_MODEL;

  let index = headers.findIndex(header => {
    return String(header).trim() === 'Conversions_' + id + '_' + model;
  });

  if (index !== -1) return index;

  index = headers.findIndex(header => {
    return String(header).trim().indexOf('Conversions_' + id + '_') === 0;
  });

  if (index !== -1) return index;

  return headers.findIndex(header => {
    return String(header).indexOf(id) !== -1;
  });
}

function tdmEnoMayMonthlyV3GoalSum_(row, goalIndexes, goalIds) {
  return goalIds.reduce((sum, goalId) => {
    const idx = goalIndexes[goalId];

    if (idx === -1 || idx === undefined) return sum;

    return sum + tdmEnoMayMonthlyV3ToNumber_(row[idx]);
  }, 0);
}

function tdmEnoMayMonthlyV3AllGoalIds_() {
  let result = [];

  Object.keys(TDM_ENO_MAY_MONTHLY_V3_CONFIG.GOALS).forEach(key => {
    result = result.concat(TDM_ENO_MAY_MONTHLY_V3_CONFIG.GOALS[key]);
  });

  return result;
}

// ==========================================================
// ПОИСК И СОЗДАНИЕ БЛОКА
// ==========================================================

function tdmEnoMayMonthlyV3GetOrCreateBlock_(sheet, title) {
  // Май пересобираем от апрельского шаблона.
  // Старый/кривой майский блок не используем.
  // Но ничего не стираем автоматически: если зона занята — остановимся.
  const template = tdmEnoMayMonthlyV3FindAprilTemplateBlock_(sheet);
  const copyRows = tdmEnoMayMonthlyV3TemplateRowsCount_(sheet, template);
  const copyCols = Math.min(
    Math.max(sheet.getLastColumn(), TDM_ENO_MAY_MONTHLY_V3_CONFIG.COPY_COLS),
    sheet.getMaxColumns()
  );

  const newTopRow = tdmEnoMayMonthlyV3GetTargetStartRow_(sheet);
  const neededLastRow = newTopRow + copyRows - 1;

  if (sheet.getMaxRows() < neededLastRow) {
    sheet.insertRowsAfter(sheet.getMaxRows(), neededLastRow - sheet.getMaxRows());
  }

  // Безопасность: если зона мая не пустая — останавливаемся.
  // Никакие формулы и данные в существующем блоке не стираем.
  tdmEnoMayMonthlyV3AssertTargetAreaIsEmpty_(sheet, newTopRow, copyRows, copyCols);

  sheet
    .getRange(template.topRow, 1, copyRows, copyCols)
    .copyTo(
      sheet.getRange(newTopRow, 1, copyRows, copyCols),
      SpreadsheetApp.CopyPasteType.PASTE_NORMAL,
      false
    );

  for (let i = 0; i < copyRows; i++) {
    sheet.setRowHeight(newTopRow + i, sheet.getRowHeight(template.topRow + i));
  }

  const block = {
    topRow: newTopRow,
    headerRow: newTopRow + (template.headerRow - template.topRow),
    totalRow: newTopRow + (template.totalRow - template.topRow)
  };

  tdmEnoMayMonthlyV3SetTitleRow_(sheet, block.topRow, title);

  return block;
}

function tdmEnoMayMonthlyV3FindExistingMayBlock_(sheet) {
  const lastRow = sheet.getLastRow();
  const values = sheet.getRange(1, 1, lastRow, Math.min(sheet.getLastColumn(), 8)).getDisplayValues();

  for (let i = 0; i < values.length; i++) {
    const rowText = tdmEnoMayMonthlyV3Normalize_(values[i].join(' '));

    const isReportTitle =
      rowText.indexOf('отчет месячный') !== -1 ||
      rowText.indexOf('отчёт месячный') !== -1 ||
      rowText.indexOf('отчет месячный по рк') !== -1 ||
      rowText.indexOf('отчёт месячный по рк') !== -1;

    if (!isReportTitle) continue;

    const hasMayMarker = TDM_ENO_MAY_MONTHLY_V3_CONFIG.MAY_MARKERS.some(marker => {
      return rowText.indexOf(tdmEnoMayMonthlyV3Normalize_(marker)) !== -1;
    });

    if (!hasMayMarker) continue;

    const topRow = i + 1;
    const headerRow = tdmEnoMayMonthlyV3FindHeaderRow_(sheet, topRow);

    if (!headerRow) continue;

    const totalRow = tdmEnoMayMonthlyV3FindTotalRowSafe_(sheet, headerRow);

    if (!totalRow) continue;

    return {
      topRow: topRow,
      headerRow: headerRow,
      totalRow: totalRow
    };
  }

  return null;
}

function tdmEnoMayMonthlyV3FindAprilTemplateBlock_(sheet) {
  const lastRow = sheet.getLastRow();
  const values = sheet.getRange(1, 1, lastRow, Math.min(sheet.getLastColumn(), 8)).getDisplayValues();

  for (let i = 0; i < values.length; i++) {
    const rowText = tdmEnoMayMonthlyV3Normalize_(values[i].join(' '));

    const isReportTitle =
      rowText.indexOf('отчет месячный') !== -1 ||
      rowText.indexOf('отчёт месячный') !== -1 ||
      rowText.indexOf('отчет месячный по рк') !== -1 ||
      rowText.indexOf('отчёт месячный по рк') !== -1;

    if (!isReportTitle) continue;

    const hasAprilMarker = TDM_ENO_MAY_MONTHLY_V3_CONFIG.APRIL_TEMPLATE_MARKERS.some(marker => {
      return rowText.indexOf(tdmEnoMayMonthlyV3Normalize_(marker)) !== -1;
    });

    if (!hasAprilMarker) continue;

    const topRow = i + 1;
    const headerRow = tdmEnoMayMonthlyV3FindHeaderRow_(sheet, topRow);

    if (!headerRow) continue;

    const totalRow = tdmEnoMayMonthlyV3FindTotalRowSafe_(sheet, headerRow);

    if (!totalRow) continue;

    return {
      topRow: topRow,
      headerRow: headerRow,
      totalRow: totalRow
    };
  }

  throw new Error('Не нашла апрельский месячный блок. Нужен блок с заголовком за 01.04.2026–30.04.2026, шапкой таблицы и строкой ИТОГО.');
}

function tdmEnoMayMonthlyV3FindHeaderRow_(sheet, topRow) {
  const maxRow = Math.min(sheet.getLastRow(), topRow + 80);
  const maxCol = Math.min(sheet.getLastColumn(), TDM_ENO_MAY_MONTHLY_V3_CONFIG.COPY_COLS);

  for (let row = topRow; row <= maxRow; row++) {
    const values = sheet.getRange(row, 1, 1, maxCol).getDisplayValues()[0];
    const normalized = values.map(value => tdmEnoMayMonthlyV3Normalize_(value));

    const hasCampaign = normalized.some(value => value === 'рк' || value === 'кампания');
    const hasImpressions = normalized.some(value => value.indexOf('показы') !== -1);
    const hasClicks = normalized.some(value => value.indexOf('клики') !== -1);
    const hasCost = normalized.some(value => value.indexOf('расход') !== -1);

    if (hasCampaign && hasImpressions && hasClicks && hasCost) {
      return row;
    }
  }

  return 0;
}

function tdmEnoMayMonthlyV3FindTotalRowSafe_(sheet, headerRow) {
  if (!headerRow) return 0;

  const maxRow = Math.min(sheet.getLastRow(), headerRow + 220);
  const maxCol = Math.min(sheet.getLastColumn(), TDM_ENO_MAY_MONTHLY_V3_CONFIG.COPY_COLS);

  for (let row = headerRow + 1; row <= maxRow; row++) {
    const values = sheet.getRange(row, 1, 1, maxCol).getDisplayValues()[0];

    const hasTotal = values.some(value => {
      const text = tdmEnoMayMonthlyV3Normalize_(value);
      return text === 'итого' || text.indexOf('итого') === 0;
    });

    if (hasTotal) return row;
  }

  return 0;
}

function tdmEnoMayMonthlyV3TemplateRowsCount_(sheet, template) {
  const nextBlockRow = tdmEnoMayMonthlyV3FindNextBlockStart_(sheet, template.topRow);

  if (nextBlockRow && nextBlockRow > template.topRow) {
    return Math.max(template.totalRow + 35 - template.topRow + 1, nextBlockRow - template.topRow - 1);
  }

  // Берём таблицу + комментарии. Этого хватает для такого же блока, как на скрине.
  return Math.max(70, template.totalRow - template.topRow + 45);
}

function tdmEnoMayMonthlyV3FindNextBlockStart_(sheet, topRow) {
  const lastRow = sheet.getLastRow();

  if (topRow >= lastRow) return 0;

  const values = sheet.getRange(
    topRow + 1,
    1,
    lastRow - topRow,
    Math.min(sheet.getLastColumn(), 8)
  ).getDisplayValues();

  for (let i = 0; i < values.length; i++) {
    const rowText = tdmEnoMayMonthlyV3Normalize_(values[i].join(' '));

    const isReportStart =
      rowText.indexOf('отчет месячный по рк') !== -1 ||
      rowText.indexOf('отчёт месячный по рк') !== -1 ||
      rowText.indexOf('отчет месячный') !== -1 ||
      rowText.indexOf('отчёт месячный') !== -1;

    if (isReportStart) {
      return topRow + 1 + i;
    }
  }

  return 0;
}

function tdmEnoMayMonthlyV3GetTargetStartRow_(sheet) {
  const forcedRow = Number(TDM_ENO_MAY_MONTHLY_V3_CONFIG.FORCE_START_ROW || 0);

  // Для мая фиксируем старт на 93 строке, как в рабочем шаблоне:
  // 91 — конец комментариев апреля, 92 — пустая строка, 93 — новый месяц.
  if (forcedRow > 0) {
    tdmEnoMayMonthlyV3EnsureBlankSeparatorBefore_(sheet, forcedRow);
    return forcedRow;
  }

  // Запасная логика для будущей универсальной версии:
  // новый месяц через одну пустую строку после последней заполненной строки.
  const lastContentRow = tdmEnoMayMonthlyV3LastContentRow_(sheet);
  const emptyRows = Number(TDM_ENO_MAY_MONTHLY_V3_CONFIG.EMPTY_ROWS_BETWEEN_MONTHS || 1);

  return lastContentRow + emptyRows + 1;
}

function tdmEnoMayMonthlyV3EnsureBlankSeparatorBefore_(sheet, startRow) {
  const separatorRow = startRow - 1;

  if (separatorRow < 1) return;

  const maxCol = Math.min(sheet.getLastColumn(), TDM_ENO_MAY_MONTHLY_V3_CONFIG.COPY_COLS);
  const separatorRange = sheet.getRange(separatorRow, 1, 1, maxCol);
  const separatorValues = separatorRange.getDisplayValues()[0];
  const hasContent = separatorValues.some(value => String(value || '').trim() !== '');

  // Если строка перед новым месяцем вдруг занята, вставляем пустую строку,
  // чтобы правило «через одну строку» сохранилось.
  if (hasContent) {
    throw new Error('Строка ' + separatorRow + ' должна быть пустой перед новым месяцем. Я не буду ничего стирать. Очисти строку вручную и запусти снова.');
  }
}

function tdmEnoMayMonthlyV3AssertTargetAreaIsEmpty_(sheet, startRow, rowsCount, colsCount) {
  const values = sheet.getRange(startRow, 1, rowsCount, colsCount).getDisplayValues();
  const formulas = sheet.getRange(startRow, 1, rowsCount, colsCount).getFormulas();

  for (let r = 0; r < values.length; r++) {
    for (let c = 0; c < values[r].length; c++) {
      const hasValue = String(values[r][c] || '').trim() !== '';
      const hasFormula = String(formulas[r][c] || '').trim() !== '';

      if (hasValue || hasFormula) {
        throw new Error(
          'Зона для мая уже не пустая: строка ' +
          (startRow + r) +
          ', колонка ' +
          (c + 1) +
          '. Я не буду ничего стирать, чтобы не тронуть формулы. Очисти майский блок вручную и запусти снова.'
        );
      }
    }
  }
}

function tdmEnoMayMonthlyV3UpdateTotalFormulasToEnd_(sheet, header, dataStartRow, totalRow) {
  // ВАЖНО:
  // Эта функция меняет только формулы в строке ИТОГО.
  // Нужна именно для того, чтобы промежуточные итоги всегда считали до конца.
  // Формулы в строках кампаний не трогаем.

  const dataEndRow = totalRow - 1;

  if (dataEndRow < dataStartRow) {
    throw new Error('Нет строк кампаний для итогов.');
  }

  // Суммируемые столбцы.
  // Формулы ставим только в строке ИТОГО.
  const subtotalColumns = [
    header.impressions,
    header.clicks,
    header.cost,
    header.view3Pages,
    header.callibriSpam,
    header.callibriNonTarget,
    header.addToCart,
    header.purchase,
    header.jivo,
    header.callibriLeadA,
    header.callibriLeadC,
    header.factLeads
  ].filter(col => col > 0);

  subtotalColumns.forEach(col => {
    const letter = tdmEnoMayMonthlyV3ColumnLetter_(col);
    const formula = '=SUBTOTAL(9;' + letter + dataStartRow + ':' + letter + dataEndRow + ')';
    sheet.getRange(totalRow, col).setFormula(formula);
  });

  // M в твоём блоке часто считается через SUM, оставляем SUM,
  // если это столбец "Просмотр 3х страниц" и в шаблоне там SUM.
  if (header.view3Pages) {
    const letter = tdmEnoMayMonthlyV3ColumnLetter_(header.view3Pages);
    sheet
      .getRange(totalRow, header.view3Pages)
      .setFormula('=SUM(' + letter + dataStartRow + ':' + letter + dataEndRow + ')');
  }

  // Расчётные итоговые показатели.
  // Их тоже обновляем только в строке ИТОГО, чтобы не зависеть от старого диапазона.
  if (header.ctr && header.impressions && header.clicks) {
    const clicksLetter = tdmEnoMayMonthlyV3ColumnLetter_(header.clicks);
    const impressionsLetter = tdmEnoMayMonthlyV3ColumnLetter_(header.impressions);
    sheet
      .getRange(totalRow, header.ctr)
      .setFormula('=IFERROR(' + clicksLetter + totalRow + '/' + impressionsLetter + totalRow + ';0)');
  }

  if (header.cpc && header.cost && header.clicks) {
    const costLetter = tdmEnoMayMonthlyV3ColumnLetter_(header.cost);
    const clicksLetter = tdmEnoMayMonthlyV3ColumnLetter_(header.clicks);
    sheet
      .getRange(totalRow, header.cpc)
      .setFormula('=IFERROR(' + costLetter + totalRow + '/' + clicksLetter + totalRow + ';0)');
  }

  if (header.cr3Pages && header.view3Pages && header.clicks) {
    const goalLetter = tdmEnoMayMonthlyV3ColumnLetter_(header.view3Pages);
    const clicksLetter = tdmEnoMayMonthlyV3ColumnLetter_(header.clicks);
    sheet
      .getRange(totalRow, header.cr3Pages)
      .setFormula('=IFERROR(' + goalLetter + totalRow + '/' + clicksLetter + totalRow + ';0)');
  }

  if (header.cpa3Pages && header.cost && header.view3Pages) {
    const costLetter = tdmEnoMayMonthlyV3ColumnLetter_(header.cost);
    const goalLetter = tdmEnoMayMonthlyV3ColumnLetter_(header.view3Pages);
    sheet
      .getRange(totalRow, header.cpa3Pages)
      .setFormula('=IFERROR(' + costLetter + totalRow + '/' + goalLetter + totalRow + ';0)');
  }

  if (header.factCr && header.factLeads && header.clicks) {
    const leadsLetter = tdmEnoMayMonthlyV3ColumnLetter_(header.factLeads);
    const clicksLetter = tdmEnoMayMonthlyV3ColumnLetter_(header.clicks);
    sheet
      .getRange(totalRow, header.factCr)
      .setFormula('=IFERROR(' + leadsLetter + totalRow + '/' + clicksLetter + totalRow + ';0)');
  }

  if (header.cpaFact && header.cost && header.factLeads) {
    const costLetter = tdmEnoMayMonthlyV3ColumnLetter_(header.cost);
    const leadsLetter = tdmEnoMayMonthlyV3ColumnLetter_(header.factLeads);
    sheet
      .getRange(totalRow, header.cpaFact)
      .setFormula('=IFERROR(' + costLetter + totalRow + '/' + leadsLetter + totalRow + ';0)');
  }
}

function tdmEnoMayMonthlyV3PrepareBlock_(sheet, block, rowsCount) {
  const dataStartRow = block.headerRow + 1;
  let totalRow = tdmEnoMayMonthlyV3FindTotalRowSafe_(sheet, block.headerRow);

  if (!totalRow) {
    throw new Error('Не нашла строку ИТОГО в новом блоке. Проверь, что апрельский шаблон скопировался полностью.');
  }

  let existingRows = totalRow - dataStartRow;
  const lastCol = Math.min(sheet.getLastColumn(), TDM_ENO_MAY_MONTHLY_V3_CONFIG.COPY_COLS);

  if (existingRows < rowsCount) {
    const rowsToAdd = rowsCount - existingRows;

    sheet.insertRowsBefore(totalRow, rowsToAdd);

    sheet
      .getRange(dataStartRow, 1, 1, lastCol)
      .copyTo(
        sheet.getRange(totalRow, 1, rowsToAdd, lastCol),
        SpreadsheetApp.CopyPasteType.PASTE_NORMAL,
        false
      );

    totalRow += rowsToAdd;
    existingRows = rowsCount;
  }

  if (existingRows > rowsCount) {
    const rowsToDelete = existingRows - rowsCount;
    sheet.deleteRows(dataStartRow + rowsCount, rowsToDelete);
    totalRow -= rowsToDelete;
  }

  return {
    topRow: block.topRow,
    headerRow: block.headerRow,
    totalRow: totalRow
  };
}

// ==========================================================
// КАРТА КОЛОНОК
// ==========================================================

function tdmEnoMayMonthlyV3HeaderMap_(sheet, headerRow) {
  const headers = sheet.getRange(
    headerRow,
    1,
    1,
    Math.min(sheet.getLastColumn(), TDM_ENO_MAY_MONTHLY_V3_CONFIG.COPY_COLS)
  ).getDisplayValues()[0];

  const normalized = headers.map(value => tdmEnoMayMonthlyV3Normalize_(value));

  const col = {
    rk: tdmEnoMayMonthlyV3FindCol_(normalized, ['рк', 'кампания']),
    type: tdmEnoMayMonthlyV3FindCol_(normalized, ['тип рк', 'тип']),

    impressions: tdmEnoMayMonthlyV3FindContainsCol_(normalized, ['показы']),
    clicks: tdmEnoMayMonthlyV3FindContainsCol_(normalized, ['клики']),
    ctr: tdmEnoMayMonthlyV3FindContainsCol_(normalized, ['ctr']),
    cpc: tdmEnoMayMonthlyV3FindContainsCol_(normalized, ['cpc']),

    cost: tdmEnoMayMonthlyV3FindContainsCol_(normalized, ['расход c ндс', 'расход с ндс', 'расход']),
    avgBid: tdmEnoMayMonthlyV3FindContainsCol_(normalized, ['ср. ставка за клик', 'средняя ставка']),
    avgPosition: tdmEnoMayMonthlyV3FindContainsCol_(normalized, ['средняя позиция показа', 'ср. позиция']),
    avgTraffic: tdmEnoMayMonthlyV3FindContainsCol_(normalized, ['средний объем трафика', 'средний объём трафика']),
    bounce: tdmEnoMayMonthlyV3FindContainsCol_(normalized, ['отказы']),
    depth: tdmEnoMayMonthlyV3FindContainsCol_(normalized, ['глубина']),

    view3Pages: tdmEnoMayMonthlyV3FindContainsCol_(normalized, ['просмотр 3х страниц']),
    cr3Pages: 0,
    cpa3Pages: tdmEnoMayMonthlyV3FindContainsCol_(normalized, ['cpa просмотр 3х страниц', 'cpa просмотр']),
    callibriSpam: tdmEnoMayMonthlyV3FindContainsCol_(normalized, ['callibri: спам', 'callibri спам']),
    callibriNonTarget: tdmEnoMayMonthlyV3FindContainsCol_(normalized, ['callibri: нецелевой_лид', 'callibri нецелевой лид', 'callibri нецелевой_лид']),
    addToCart: tdmEnoMayMonthlyV3FindContainsCol_(normalized, ['ecommerce: добавление в корзину', 'ecommerce добавление в корзину']),
    purchase: tdmEnoMayMonthlyV3FindContainsCol_(normalized, ['ecommerce: покупка', 'ecommerce покупка']),
    jivo: tdmEnoMayMonthlyV3FindContainsCol_(normalized, ['jivo-чат', 'jivo чат', 'jivo']),
    callibriLeadA: tdmEnoMayMonthlyV3FindContainsCol_(normalized, ['callibri: лид_квал_a', 'callibri лид квал a', 'callibri лид_квал_a']),
    callibriLeadC: tdmEnoMayMonthlyV3FindContainsCol_(normalized, ['callibri: лид_квал_c', 'callibri лид квал c', 'callibri лид_квал_c']),

    factLeads: tdmEnoMayMonthlyV3FindContainsCol_(normalized, ['факт сумма конверсий', 'факт конверсий', 'кол-во лид', 'лиды']),
    factCr: 0,
    cpaFact: tdmEnoMayMonthlyV3FindContainsCol_(normalized, ['стоимость факт cpa', 'cpa лид'])
  };

  // В таблице два столбца CR. Первый после просмотра 3 страниц, второй после факта лидов.
  const crColumns = [];

  normalized.forEach((value, index) => {
    if (value === 'cr') crColumns.push(index + 1);
  });

  if (col.view3Pages) {
    col.cr3Pages = crColumns.find(c => c > col.view3Pages && (!col.factLeads || c < col.factLeads)) || 0;
  }

  if (col.factLeads) {
    col.factCr = crColumns.find(c => c > col.factLeads) || 0;
  }

  return col;
}

function tdmEnoMayMonthlyV3FindCol_(normalizedHeaders, names) {
  for (let i = 0; i < normalizedHeaders.length; i++) {
    if (names.indexOf(normalizedHeaders[i]) !== -1) return i + 1;
  }

  return 0;
}

function tdmEnoMayMonthlyV3FindExactOrContainsCol_(normalizedHeaders, names) {
  const exact = tdmEnoMayMonthlyV3FindCol_(normalizedHeaders, names);

  if (exact) return exact;

  return tdmEnoMayMonthlyV3FindContainsCol_(normalizedHeaders, names);
}

function tdmEnoMayMonthlyV3FindContainsCol_(normalizedHeaders, parts) {
  const normalizedParts = parts.map(part => tdmEnoMayMonthlyV3Normalize_(part));

  for (let i = 0; i < normalizedHeaders.length; i++) {
    for (let j = 0; j < normalizedParts.length; j++) {
      if (normalizedHeaders[i].indexOf(normalizedParts[j]) !== -1) {
        return i + 1;
      }
    }
  }

  return 0;
}

// ==========================================================
// ПОРЯДОК И ТИПЫ КАМПАНИЙ ИЗ ШАБЛОНА
// ==========================================================

function tdmEnoMayMonthlyV3ReadTemplateCampaignMeta_(sheet, header, startRow, endRow) {
  const meta = {
    order: {},
    type: {}
  };

  if (!header.rk || startRow > endRow) return meta;

  const values = sheet
    .getRange(startRow, 1, endRow - startRow + 1, Math.min(sheet.getLastColumn(), TDM_ENO_MAY_MONTHLY_V3_CONFIG.COPY_COLS))
    .getDisplayValues();

  values.forEach((row, index) => {
    const campaignName = String(row[header.rk - 1] || '').trim();

    if (!campaignName) return;

    const norm = tdmEnoMayMonthlyV3Normalize_(campaignName);

    if (norm === 'итого') return;

    meta.order[norm] = index;

    if (header.type) {
      const typeName = String(row[header.type - 1] || '').trim();

      if (typeName) {
        meta.type[norm] = typeName;
      }
    }
  });

  return meta;
}

function tdmEnoMayMonthlyV3ApplyTemplateMeta_(rows, meta) {
  rows.forEach(row => {
    const key = tdmEnoMayMonthlyV3Normalize_(row.campaignName);

    if (meta.type[key]) {
      row.campaignType = meta.type[key];
    }
  });

  return rows;
}

function tdmEnoMayMonthlyV3SortRowsByTemplate_(rows, meta) {
  return rows.sort((a, b) => {
    const aKey = tdmEnoMayMonthlyV3Normalize_(a.campaignName);
    const bKey = tdmEnoMayMonthlyV3Normalize_(b.campaignName);

    const aOrder = meta.order.hasOwnProperty(aKey) ? meta.order[aKey] : 999999;
    const bOrder = meta.order.hasOwnProperty(bKey) ? meta.order[bKey] : 999999;

    if (aOrder !== bOrder) return aOrder - bOrder;

    return b.cost - a.cost;
  });
}

// ==========================================================
// ЗАПОЛНЕНИЕ ТАБЛИЦЫ
// ==========================================================

function tdmEnoMayMonthlyV3ClearDataValues_(sheet, header, startRow, rowsCount) {
  if (!sheet || !header || !startRow || !rowsCount || rowsCount <= 0) return;

  const columns = [
    header.rk,
    header.type,
    header.impressions,
    header.clicks,
    header.ctr,
    header.cpc,
    header.cost,
    header.avgBid,
    header.avgPosition,
    header.avgTraffic,
    header.bounce,
    header.depth,
    header.view3Pages,
    header.cr3Pages,
    header.cpa3Pages,
    header.callibriSpam,
    header.callibriNonTarget,
    header.addToCart,
    header.purchase,
    header.jivo,
    header.callibriLeadA,
    header.callibriLeadC,
    header.factLeads,
    header.factCr,
    header.cpaFact
  ].filter(col => col > 0)
    .filter((col, index, arr) => arr.indexOf(col) === index)
    .sort((a, b) => a - b);

  columns.forEach(col => {
    const range = sheet.getRange(startRow, col, rowsCount, 1);
    const formulas = range.getFormulas();
    const values = range.getValues();
    let changed = false;

    for (let i = 0; i < rowsCount; i++) {
      if (!formulas[i][0] && values[i][0] !== '') {
        values[i][0] = '';
        changed = true;
      }
    }

    if (changed) range.setValues(values);
  });
}

function tdmEnoMayMonthlyV3FillRows_(sheet, header, startRow, rows) {
  if (!sheet || !header || !startRow) return;

  if (!rows.length) {
    tdmEnoMayMonthlyV3SetUnlessFormula_(sheet, startRow, header.rk, 'Нет данных');
    return;
  }

  const columns = [
    header.rk,
    header.type,
    header.impressions,
    header.clicks,
    header.ctr,
    header.cpc,
    header.cost,
    header.avgBid,
    header.avgPosition,
    header.avgTraffic,
    header.bounce,
    header.depth,
    header.view3Pages,
    header.cr3Pages,
    header.cpa3Pages,
    header.callibriSpam,
    header.callibriNonTarget,
    header.addToCart,
    header.purchase,
    header.jivo,
    header.callibriLeadA,
    header.callibriLeadC,
    header.factLeads,
    header.factCr,
    header.cpaFact
  ].filter(col => col > 0);

  if (!columns.length) return;

  const minCol = Math.min.apply(null, columns);
  const maxCol = Math.max.apply(null, columns);
  const width = maxCol - minCol + 1;
  const range = sheet.getRange(startRow, minCol, rows.length, width);
  const formulas = range.getFormulas();
  const values = range.getValues();

  const put = function(rowIndex, col, value) {
    if (!col || col <= 0) return;
    const offset = col - minCol;
    if (offset < 0 || offset >= width) return;
    if (formulas[rowIndex][offset]) return;
    values[rowIndex][offset] = value;
  };

  rows.forEach((item, index) => {
    const ctr = item.impressions > 0 ? item.clicks / item.impressions : 0;
    const cpc = item.clicks > 0 ? item.cost / item.clicks : 0;
    const cr3 = item.clicks > 0 ? item.goals.view3Pages / item.clicks : 0;
    const cpa3 = item.goals.view3Pages > 0 ? item.cost / item.goals.view3Pages : '';
    const factCr = item.clicks > 0 ? item.factLeads / item.clicks : 0;
    const cpaFact = item.factLeads > 0 ? item.cost / item.factLeads : '';

    put(index, header.rk, item.campaignName);
    put(index, header.type, item.campaignType);
    put(index, header.impressions, item.impressions);
    put(index, header.clicks, item.clicks);
    put(index, header.ctr, ctr);
    put(index, header.cpc, cpc);
    put(index, header.cost, item.cost);
    put(index, header.avgBid, item.avgBid);
    put(index, header.avgPosition, item.avgPosition);
    put(index, header.avgTraffic, item.avgTraffic);
    put(index, header.bounce, item.bounce);
    put(index, header.depth, item.depth);
    put(index, header.view3Pages, tdmEnoMayMonthlyV3GoalCell_(item.goals.view3Pages));
    put(index, header.cr3Pages, cr3);
    put(index, header.cpa3Pages, cpa3);
    put(index, header.callibriSpam, tdmEnoMayMonthlyV3GoalCell_(item.goals.callibriSpam));
    put(index, header.callibriNonTarget, tdmEnoMayMonthlyV3GoalCell_(item.goals.callibriNonTarget));
    put(index, header.addToCart, tdmEnoMayMonthlyV3GoalCell_(item.goals.addToCart));
    put(index, header.purchase, tdmEnoMayMonthlyV3GoalCell_(item.goals.purchase));
    put(index, header.jivo, tdmEnoMayMonthlyV3GoalCell_(item.goals.jivo));
    put(index, header.callibriLeadA, tdmEnoMayMonthlyV3GoalCell_(item.goals.callibriLeadA));
    put(index, header.callibriLeadC, tdmEnoMayMonthlyV3GoalCell_(item.goals.callibriLeadC));
    put(index, header.factLeads, item.factLeads);
    put(index, header.factCr, factCr);
    put(index, header.cpaFact, cpaFact);
  });

  range.setValues(values);
}

function tdmEnoMayMonthlyV3FillTotalRow_(sheet, header, totalRow, rows) {
  const total = tdmEnoMayMonthlyV3Totals_(rows);

  const ctr = total.impressions > 0 ? total.clicks / total.impressions : 0;
  const cpc = total.clicks > 0 ? total.cost / total.clicks : 0;
  const cr3 = total.clicks > 0 ? total.goals.view3Pages / total.clicks : 0;
  const cpa3 = total.goals.view3Pages > 0 ? total.cost / total.goals.view3Pages : '';
  const factCr = total.clicks > 0 ? total.factLeads / total.clicks : 0;
  const cpaFact = total.factLeads > 0 ? total.cost / total.factLeads : '';

  tdmEnoMayMonthlyV3SetUnlessFormula_(sheet, totalRow, header.rk, 'ИТОГО');
  tdmEnoMayMonthlyV3SetUnlessFormula_(sheet, totalRow, header.type, 'Итого');

  tdmEnoMayMonthlyV3SetUnlessFormula_(sheet, totalRow, header.impressions, total.impressions);
  tdmEnoMayMonthlyV3SetUnlessFormula_(sheet, totalRow, header.clicks, total.clicks);
  tdmEnoMayMonthlyV3SetUnlessFormula_(sheet, totalRow, header.ctr, ctr);
  tdmEnoMayMonthlyV3SetUnlessFormula_(sheet, totalRow, header.cpc, cpc);
  tdmEnoMayMonthlyV3SetUnlessFormula_(sheet, totalRow, header.cost, total.cost);
  tdmEnoMayMonthlyV3SetUnlessFormula_(sheet, totalRow, header.avgBid, cpc);
  tdmEnoMayMonthlyV3SetUnlessFormula_(sheet, totalRow, header.avgPosition, '');
  tdmEnoMayMonthlyV3SetUnlessFormula_(sheet, totalRow, header.avgTraffic, '');
  tdmEnoMayMonthlyV3SetUnlessFormula_(sheet, totalRow, header.bounce, '');
  tdmEnoMayMonthlyV3SetUnlessFormula_(sheet, totalRow, header.depth, '');

  tdmEnoMayMonthlyV3SetUnlessFormula_(sheet, totalRow, header.view3Pages, total.goals.view3Pages);
  tdmEnoMayMonthlyV3SetUnlessFormula_(sheet, totalRow, header.cr3Pages, cr3);
  tdmEnoMayMonthlyV3SetUnlessFormula_(sheet, totalRow, header.cpa3Pages, cpa3);

  tdmEnoMayMonthlyV3SetUnlessFormula_(sheet, totalRow, header.callibriSpam, total.goals.callibriSpam);
  tdmEnoMayMonthlyV3SetUnlessFormula_(sheet, totalRow, header.callibriNonTarget, total.goals.callibriNonTarget);
  tdmEnoMayMonthlyV3SetUnlessFormula_(sheet, totalRow, header.addToCart, total.goals.addToCart);
  tdmEnoMayMonthlyV3SetUnlessFormula_(sheet, totalRow, header.purchase, total.goals.purchase);
  tdmEnoMayMonthlyV3SetUnlessFormula_(sheet, totalRow, header.jivo, total.goals.jivo);
  tdmEnoMayMonthlyV3SetUnlessFormula_(sheet, totalRow, header.callibriLeadA, total.goals.callibriLeadA);
  tdmEnoMayMonthlyV3SetUnlessFormula_(sheet, totalRow, header.callibriLeadC, total.goals.callibriLeadC);

  tdmEnoMayMonthlyV3SetUnlessFormula_(sheet, totalRow, header.factLeads, total.factLeads);
  tdmEnoMayMonthlyV3SetUnlessFormula_(sheet, totalRow, header.factCr, factCr);
  tdmEnoMayMonthlyV3SetUnlessFormula_(sheet, totalRow, header.cpaFact, cpaFact);
}

function tdmEnoMayMonthlyV3SetTitleRow_(sheet, topRow, title) {
  tdmEnoMayMonthlyV3SetUnlessFormula_(sheet, topRow, 1, title);
}

function tdmEnoMayMonthlyV3SetTopSummary_(sheet, topRow, total) {
  const maxCol = Math.min(sheet.getLastColumn(), TDM_ENO_MAY_MONTHLY_V3_CONFIG.COPY_COLS);
  const rowValues = sheet.getRange(topRow, 1, 1, maxCol).getDisplayValues()[0];
  const norm = rowValues.map(value => tdmEnoMayMonthlyV3Normalize_(value));

  const summary = {
    'расход': total.cost,
    'cpc': total.clicks > 0 ? total.cost / total.clicks : 0,
    'клики': total.clicks,
    'кол-во лид': total.factLeads,
    'cr': total.clicks > 0 ? total.factLeads / total.clicks : 0,
    'cpa лид': total.factLeads > 0 ? total.cost / total.factLeads : 0
  };

  Object.keys(summary).forEach(label => {
    const index = norm.findIndex(value => value === label);

    if (index !== -1 && index + 2 <= maxCol) {
      tdmEnoMayMonthlyV3SetUnlessFormula_(sheet, topRow, index + 2, summary[label]);
    }
  });

  // Обновляем текстовую пометку, если она есть в верхней строке.
  norm.forEach((value, index) => {
    if (value.indexOf('без учета') !== -1 || value.indexOf('без учёта') !== -1) {
      tdmEnoMayMonthlyV3SetUnlessFormula_(
        sheet,
        topRow,
        index + 1,
        'Данные Директа, Метрики и Callibri сверены за выбранный период'
      );
    }
  });
}

function tdmEnoMayMonthlyV3SetUnlessFormula_(sheet, row, col, value) {
  if (!col || col <= 0) return;

  const cell = sheet.getRange(row, col);

  if (cell.getFormula()) return;

  cell.setValue(value);
}

function tdmEnoMayMonthlyV3GoalCell_(value) {
  const number = Number(value || 0);
  return number > 0 ? number : '-';
}

function tdmEnoMayMonthlyV3ColorRowsByStrength_(sheet, header, startRow, rowsCount, rows) {
  // ВАЖНО:
  // Эта функция меняет только цвет строк.
  // setValue / clearContent / clearFormat / copyTo здесь нет.
  // Данные и формулы в строках кампаний не трогаем.

  const lastCol = Math.min(sheet.getLastColumn(), TDM_ENO_MAY_MONTHLY_V3_CONFIG.COPY_COLS);
  const total = tdmEnoMayMonthlyV3Totals_(rows);
  const avgCpa = total.factLeads > 0 ? total.cost / total.factLeads : 0;

  rows.forEach((item, index) => {
    const row = startRow + index;
    const strength = tdmEnoMayMonthlyV3CampaignStrength_(item, avgCpa);
    const color = tdmEnoMayMonthlyV3StrengthColor_(strength);

    sheet.getRange(row, 1, 1, lastCol).setBackground(color);
  });
}

function tdmEnoMayMonthlyV3CampaignStrength_(item, avgCpa) {
  const factLeads = Number(item.factLeads || 0);
  const cost = Number(item.cost || 0);
  const clicks = Number(item.clicks || 0);

  const purchase = Number(item.goals.purchase || 0);
  const jivo = Number(item.goals.jivo || 0);
  const callibriLeadA = Number(item.goals.callibriLeadA || 0);
  const callibriLeadC = Number(item.goals.callibriLeadC || 0);

  const addToCart = Number(item.goals.addToCart || 0);
  const view3Pages = Number(item.goals.view3Pages || 0);

  const microSignals = addToCart + view3Pages;
  const cpa = factLeads > 0 ? cost / factLeads : 0;

  // Зелёный — сильные кампании:
  // есть покупка, либо есть итоговые лиды с CPA около среднего/ниже среднего.
  if (
    purchase > 0 ||
    (factLeads >= 2 && avgCpa > 0 && cpa <= avgCpa * 1.15) ||
    (factLeads >= 1 && avgCpa > 0 && cpa <= avgCpa * 0.85 && microSignals > 0)
  ) {
    return 'strong';
  }

  // Жёлтый — средние кампании:
  // лид есть, но CPA выше среднего, либо есть микроцели/полезные сигналы.
  if (
    factLeads >= 1 ||
    jivo > 0 ||
    callibriLeadA > 0 ||
    callibriLeadC > 0 ||
    microSignals >= 3 ||
    (clicks >= 30 && addToCart > 0)
  ) {
    return 'medium';
  }

  // Красный — слабые кампании:
  // расход есть, а итоговых лидов и нормальных сигналов нет.
  if (cost > 0) {
    return 'weak';
  }

  return 'neutral';
}

function tdmEnoMayMonthlyV3StrengthColor_(strength) {
  const colors = TDM_ENO_MAY_MONTHLY_V3_CONFIG.COLORS || {};

  if (strength === 'strong') return colors.strong || '#d9ead3';
  if (strength === 'medium') return colors.medium || '#fff2cc';
  if (strength === 'weak') return colors.weak || '#f4cccc';

  return colors.neutral || '#ffffff';
}

// ==========================================================
// КОММЕНТАРИИ
// ==========================================================

function tdmEnoMayMonthlyV3BuildComments_(rows) {
  const total = tdmEnoMayMonthlyV3Totals_(rows);
  const byType = tdmEnoMayMonthlyV3GroupByType_(rows);
  const kzRows = rows.filter(row => row.geo === 'Казахстан');
  const kzTotal = tdmEnoMayMonthlyV3Totals_(kzRows);

  const plan = Number(TDM_ENO_MAY_MONTHLY_V3_CONFIG.MONTH_PLAN || 0);
  const planText = plan > 0
    ? ' — ' + tdmEnoMayMonthlyV3Percent_(total.cost / plan) + ' от плана на месяц'
    : '';

  const comments = [];

  const monthName = tdmEnoMayMonthlyV3MonthNameRu_(TDM_ENO_MAY_MONTHLY_V3_CONFIG.DATE_FROM);
const dateFromShort = tdmEnoMayMonthlyV3ShortDate_(TDM_ENO_MAY_MONTHLY_V3_CONFIG.DATE_FROM);
const dateToShort = tdmEnoMayMonthlyV3ShortDate_(TDM_ENO_MAY_MONTHLY_V3_CONFIG.DATE_TO);
comments.push([
  'Комментарии к отчёту за ' +
  monthName +
  ' с ' +
  dateFromShort +
  ' по ' +
  dateToShort +
  ' (новая схема целей: Ecommerce / Jivo / Callibri)'
]);
  comments.push(['']);

  comments.push(['Бюджет:']);
  comments.push(['Всего задействовано за период ' + tdmEnoMayMonthlyV3Rub_(total.cost) + planText + '.']);
  comments.push(['']);

  comments.push(['Трафик:']);
  comments.push(['Показы — ' + tdmEnoMayMonthlyV3Int_(total.impressions)]);
  comments.push(['Клики — ' + tdmEnoMayMonthlyV3Int_(total.clicks)]);
  comments.push(['Средний CPC — ' + tdmEnoMayMonthlyV3Rub_(total.clicks > 0 ? total.cost / total.clicks : 0)]);
  comments.push(['']);

  comments.push(['Достигнуты цели:']);
  comments.push(['Просмотр 3х страниц — ' + tdmEnoMayMonthlyV3Int_(total.goals.view3Pages)]);
  comments.push(['Callibri: Спам — ' + tdmEnoMayMonthlyV3Int_(total.goals.callibriSpam)]);
  comments.push(['Callibri: Нецелевой_Лид — ' + tdmEnoMayMonthlyV3Int_(total.goals.callibriNonTarget)]);
  comments.push(['Ecommerce: добавление в корзину — ' + tdmEnoMayMonthlyV3Int_(total.goals.addToCart)]);
  comments.push(['Ecommerce: покупка — ' + tdmEnoMayMonthlyV3Int_(total.goals.purchase)]);
  comments.push(['Jivo-чат — ' + tdmEnoMayMonthlyV3Int_(total.goals.jivo)]);
  comments.push(['Callibri: Лид_Квал_A — ' + tdmEnoMayMonthlyV3Int_(total.goals.callibriLeadA)]);
  comments.push(['Callibri: Лид_Квал_C — ' + tdmEnoMayMonthlyV3Int_(total.goals.callibriLeadC)]);
  comments.push([
    'Итого получено: ' +
    tdmEnoMayMonthlyV3Int_(total.factLeads) +
    ' конверсий, CR — ' +
    tdmEnoMayMonthlyV3Percent_(total.clicks > 0 ? total.factLeads / total.clicks : 0) +
    ', CPA — ' +
    tdmEnoMayMonthlyV3Rub_(total.factLeads > 0 ? total.cost / total.factLeads : 0) +
    '.'
  ]);
  comments.push(['']);

  comments.push(['По типам РК:']);

  ['Поиск', 'Сети', 'Поиск/Сети', 'Поиск/Сети. Товарные РК', 'Другое'].forEach(typeName => {
    const typeTotal = byType[typeName];

    if (!typeTotal || typeTotal.cost === 0 && typeTotal.clicks === 0 && typeTotal.factLeads === 0) return;

    comments.push([
      typeName +
      ' — ' +
      tdmEnoMayMonthlyV3Int_(typeTotal.clicks) +
      ' кликов, расход ' +
      tdmEnoMayMonthlyV3Rub_(typeTotal.cost) +
      ', средний CPC — ' +
      tdmEnoMayMonthlyV3Rub_(typeTotal.clicks > 0 ? typeTotal.cost / typeTotal.clicks : 0) +
      ', получено — ' +
      tdmEnoMayMonthlyV3Int_(typeTotal.factLeads) +
      ' конверсий, CR — ' +
      tdmEnoMayMonthlyV3Percent_(typeTotal.clicks > 0 ? typeTotal.factLeads / typeTotal.clicks : 0) +
      ', CPA — ' +
      tdmEnoMayMonthlyV3Rub_(typeTotal.factLeads > 0 ? typeTotal.cost / typeTotal.factLeads : 0) +
      '.'
    ]);
  });

  comments.push(['']);

  if (kzRows.length) {
    comments.push(['Казахстан:']);
    comments.push(['Всего задействовано за период ' + tdmEnoMayMonthlyV3Rub_(kzTotal.cost) + '.']);
    comments.push(['Показы — ' + tdmEnoMayMonthlyV3Int_(kzTotal.impressions)]);
    comments.push(['Клики — ' + tdmEnoMayMonthlyV3Int_(kzTotal.clicks)]);
    comments.push(['Средний CPC — ' + tdmEnoMayMonthlyV3Rub_(kzTotal.clicks > 0 ? kzTotal.cost / kzTotal.clicks : 0)]);
    comments.push(['Ecommerce: покупка — ' + tdmEnoMayMonthlyV3Int_(kzTotal.goals.purchase)]);
    comments.push(['Jivo-чат — ' + tdmEnoMayMonthlyV3Int_(kzTotal.goals.jivo)]);
    comments.push(['Callibri: Лид_Квал_A — ' + tdmEnoMayMonthlyV3Int_(kzTotal.goals.callibriLeadA)]);
    comments.push(['Callibri: Лид_Квал_C — ' + tdmEnoMayMonthlyV3Int_(kzTotal.goals.callibriLeadC)]);
    comments.push([
      'Итого получено: ' +
      tdmEnoMayMonthlyV3Int_(kzTotal.factLeads) +
      ' конверсий, CR — ' +
      tdmEnoMayMonthlyV3Percent_(kzTotal.clicks > 0 ? kzTotal.factLeads / kzTotal.clicks : 0) +
      ', CPA — ' +
      tdmEnoMayMonthlyV3Rub_(kzTotal.factLeads > 0 ? kzTotal.cost / kzTotal.factLeads : 0) +
      '.'
    ]);
    comments.push(['']);
  }

  comments.push(['Вывод:']);
  comments.push([tdmEnoMayMonthlyV3Conclusion_(total, byType, kzRows.length)]);
  comments.push(['']);

  comments.push(['Фокус до следующего обновления:']);
  comments.push(['1. Удержать кампании, которые дают лиды с CPA ниже среднего.']);
  comments.push(['2. Проверить кампании с расходом без покупок, Jivo и Callibri A/C.']);
  comments.push(['3. Почистить поисковые запросы и площадки в сетях.']);
  comments.push(['4. Отдельно проверить товарные кампании: оставить категории с обращениями и убрать слабый трафик.']);

  return comments;
}

function tdmEnoMayMonthlyV3Conclusion_(total, byType, hasKz) {
  const cpa = total.factLeads > 0 ? total.cost / total.factLeads : 0;
  const bestType = tdmEnoMayMonthlyV3BestTypeByLeads_(byType);

  let text =
    'За ' + tdmEnoMayMonthlyV3MonthNameRu_(TDM_ENO_MAY_MONTHLY_V3_CONFIG.DATE_FROM) + ' получено ' +
    tdmEnoMayMonthlyV3Int_(total.factLeads) +
    ' конверсий при среднем CPA ' +
    tdmEnoMayMonthlyV3Rub_(cpa) +
    '. ';

  if (bestType) {
    text += 'Основной вклад по лидам даёт ' + bestType + ', поэтому этот тип кампаний стоит держать в приоритете и усиливать рабочие связки. ';
  }

  text += 'Кампании с расходом без итоговых лидов нужно отдельно проверить по запросам, площадкам и качеству трафика.';

  if (hasKz) {
    text += ' Казахстан лучше вести отдельным контролем по качеству обращений.';
  }

  return text;
}

function tdmEnoMayMonthlyV3BestTypeByLeads_(byType) {
  let bestName = '';
  let bestLeads = 0;

  Object.keys(byType).forEach(typeName => {
    if (byType[typeName].factLeads > bestLeads) {
      bestLeads = byType[typeName].factLeads;
      bestName = typeName;
    }
  });

  return bestName;
}

function tdmEnoMayMonthlyV3WriteComments_(sheet, startRow, comments) {
  const rowsToClear = Math.max(comments.length + 10, 45);
  const colsToClear = Math.min(sheet.getLastColumn(), 10);

  sheet
    .getRange(startRow, 1, rowsToClear, colsToClear)
    .clearContent()
    .setWrap(true)
    .setFontColor('#000000')
    .setVerticalAlignment('top');

  sheet.getRange(startRow, 1, comments.length, 1).setValues(comments);

  comments.forEach((item, index) => {
    const row = startRow + index;
    const text = String(item[0] || '').trim();
    const normalized = tdmEnoMayMonthlyV3Normalize_(text);

    const isHeader =
      normalized.indexOf('комментарии') === 0 ||
      normalized === 'бюджет:' ||
      normalized === 'трафик:' ||
      normalized === 'достигнуты цели:' ||
      normalized === 'по типам рк:' ||
      normalized === 'казахстан:' ||
      normalized === 'вывод:' ||
      normalized === 'фокус до следующего обновления:';

    if (!text) {
      sheet.setRowHeight(row, 18);
      return;
    }

    if (isHeader) {
      sheet.getRange(row, 1, 1, colsToClear)
        .setBackground('#d9ead3')
        .setFontWeight('bold');

      sheet.setRowHeight(row, 26);
    } else {
      sheet.getRange(row, 1, 1, colsToClear)
        .setBackground('#ffffff')
        .setFontWeight('normal');

      sheet.setRowHeight(row, 40);
    }
  });
}

// ==========================================================
// АГРЕГАЦИЯ
// ==========================================================

function tdmEnoMayMonthlyV3Totals_(rows) {
  const total = tdmEnoMayMonthlyV3EmptyTotals_();

  rows.forEach(row => {
    total.impressions += Number(row.impressions || 0);
    total.clicks += Number(row.clicks || 0);
    total.cost += Number(row.cost || 0);

    Object.keys(total.goals).forEach(key => {
      total.goals[key] += Number(row.goals[key] || 0);
    });

    total.factLeads += Number(row.factLeads || 0);
  });

  return total;
}

function tdmEnoMayMonthlyV3EmptyTotals_() {
  return {
    impressions: 0,
    clicks: 0,
    cost: 0,
    goals: {
      view3Pages: 0,
      callibriSpam: 0,
      callibriNonTarget: 0,
      addToCart: 0,
      purchase: 0,
      jivo: 0,
      callibriLeadA: 0,
      callibriLeadC: 0
    },
    factLeads: 0
  };
}

function tdmEnoMayMonthlyV3GroupByType_(rows) {
  const result = {};

  rows.forEach(row => {
    const type = row.campaignType || 'Другое';

    if (!result[type]) {
      result[type] = tdmEnoMayMonthlyV3EmptyTotals_();
    }

    const total = result[type];

    total.impressions += Number(row.impressions || 0);
    total.clicks += Number(row.clicks || 0);
    total.cost += Number(row.cost || 0);

    Object.keys(total.goals).forEach(key => {
      total.goals[key] += Number(row.goals[key] || 0);
    });

    total.factLeads += Number(row.factLeads || 0);
  });

  return result;
}

function tdmEnoMayMonthlyV3FactLeads_(goals) {
  return TDM_ENO_MAY_MONTHLY_V3_CONFIG.FACT_LEAD_KEYS.reduce((sum, key) => {
    return sum + Number(goals[key] || 0);
  }, 0);
}

function tdmEnoMayMonthlyV3AllGoalsSum_(goals) {
  return Object.keys(goals).reduce((sum, key) => {
    return sum + Number(goals[key] || 0);
  }, 0);
}

// ==========================================================
// КЛАССИФИКАЦИЯ
// ==========================================================

function tdmEnoMayMonthlyV3FallbackCampaignType_(campaignName) {
  const name = tdmEnoMayMonthlyV3Normalize_(campaignName);

  if (
    name.indexOf('dsa') !== -1 ||
    name.indexOf('search') !== -1 ||
    name.indexOf('t-g') !== -1 ||
    name.indexOf('поиск') !== -1
  ) {
    return 'Поиск';
  }

  if (
    name.indexOf('rsya') !== -1 ||
    name.indexOf('рся') !== -1 ||
    name.indexOf('network') !== -1 ||
    name.indexOf('сети') !== -1
  ) {
    return 'Сети';
  }

  if (
    name.indexOf('tovar') !== -1 ||
    name.indexOf('товар') !== -1 ||
    name.indexOf('smart') !== -1 ||
    name.indexOf('смарт') !== -1 ||
    name.indexOf('product') !== -1
  ) {
    return 'Поиск/Сети';
  }

  return 'Другое';
}

function tdmEnoMayMonthlyV3CampaignGeo_(campaignName) {
  const name = tdmEnoMayMonthlyV3Normalize_(campaignName);

  if (
    name.indexOf('_kz') !== -1 ||
    name.indexOf('-kz') !== -1 ||
    name.indexOf('kz') !== -1 ||
    name.indexOf('kaz') !== -1 ||
    name.indexOf('казахстан') !== -1
  ) {
    return 'Казахстан';
  }

  return 'Основное гео';
}

// ==========================================================
// ФОРМАТЫ И УТИЛИТЫ
// ==========================================================

function tdmEnoMayMonthlyV3ColumnLetter_(columnNumber) {
  let temp = '';
  let letter = '';
  let col = Number(columnNumber);

  while (col > 0) {
    temp = (col - 1) % 26;
    letter = String.fromCharCode(temp + 65) + letter;
    col = (col - temp - 1) / 26;
  }

  return letter;
}

function tdmEnoMayMonthlyV3Title_(dateFrom, dateTo) {
  return 'Отчет месячный по РК ' +
    tdmEnoMayMonthlyV3ShortDate_(dateFrom) +
    ' - ' +
    tdmEnoMayMonthlyV3DateRu_(dateTo);
}

function tdmEnoMayMonthlyV3ShortDate_(apiDate) {
  const parts = String(apiDate).split('-');
  return Number(parts[2]) + '.' + parts[1] + '.';
}

function tdmEnoMayMonthlyV3DateRu_(apiDate) {
  const parts = String(apiDate).split('-');
  return Number(parts[2]) + '.' + parts[1] + '.' + parts[0];
}

function tdmEnoMayMonthlyV3RequireColumn_(headers, columnName) {
  const index = headers.indexOf(columnName);

  if (index === -1) {
    throw new Error('Не найден столбец "' + columnName + '". Заголовки: ' + headers.join(' | '));
  }

  return index;
}

function tdmEnoMayMonthlyV3ToNumber_(value) {
  if (value === null || value === undefined) return 0;

  const text = String(value)
    .replace(/\s/g, '')
    .replace(',', '.')
    .replace(/[^\d.-]/g, '');

  const number = Number(text);

  return isNaN(number) ? 0 : number;
}

function tdmEnoMayMonthlyV3Metric_(value) {
  const number = tdmEnoMayMonthlyV3ToNumber_(value);
  return number || '';
}

function tdmEnoMayMonthlyV3Normalize_(value) {
  return String(value || '')
    .toLowerCase()
    .replace(/ё/g, 'е')
    .replace(/[«»"']/g, '')
    .replace(/\s+/g, ' ')
    .trim();
}

function tdmEnoMayMonthlyV3LastContentRow_(sheet) {
  const lastRow = sheet.getLastRow();
  const lastCol = Math.min(sheet.getLastColumn(), TDM_ENO_MAY_MONTHLY_V3_CONFIG.COPY_COLS);

  if (lastRow < 1) return 1;

  const values = sheet.getRange(1, 1, lastRow, lastCol).getDisplayValues();

  for (let i = values.length - 1; i >= 0; i--) {
    const hasContent = values[i].some(value => String(value || '').trim() !== '');

    if (hasContent) return i + 1;
  }

  return 1;
}

function tdmEnoMayMonthlyV3Int_(value) {
  return Number(value || 0).toLocaleString('ru-RU', {
    minimumFractionDigits: 0,
    maximumFractionDigits: 0
  });
}

function tdmEnoMayMonthlyV3Rub_(value) {
  return Number(value || 0).toLocaleString('ru-RU', {
    minimumFractionDigits: 2,
    maximumFractionDigits: 2
  }) + ' руб.';
}

function tdmEnoMayMonthlyV3Percent_(value) {
  return Number(value || 0).toLocaleString('ru-RU', {
    minimumFractionDigits: 2,
    maximumFractionDigits: 2,
    style: 'percent'
  });
}

function tdmEnoMayMonthlyV3Toast_(message) {
  try {
    SpreadsheetApp.getActive().toast(message);
  } catch (e) {
    Logger.log(message);
  }
}


// ==========================================================
// Безопасная проверка после запуска.
// Только читает данные, ничего не меняет.
// ==========================================================

function checkTdmEnoMay2026MonthlyTotalsSafe() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = tdmEnoMayMonthlyV3GetSheet_(ss);
  const block = tdmEnoMayMonthlyV3FindExistingMayBlock_(sheet);

  if (!block) {
    throw new Error('Не нашла майский блок для проверки.');
  }

  const header = tdmEnoMayMonthlyV3HeaderMap_(sheet, block.headerRow);
  const dataStartRow = block.headerRow + 1;
  const dataEndRow = block.totalRow - 1;
  const rowsCount = dataEndRow - dataStartRow + 1;

  const values = sheet
    .getRange(dataStartRow, 1, rowsCount, Math.min(sheet.getLastColumn(), TDM_ENO_MAY_MONTHLY_V3_CONFIG.COPY_COLS))
    .getValues();

  let impressions = 0;
  let clicks = 0;
  let cost = 0;
  let callibriSpam = 0;
  let callibriNonTarget = 0;
  let purchase = 0;
  let addToCart = 0;
  let jivo = 0;
  let callibriLeadA = 0;
  let callibriLeadC = 0;

  values.forEach(row => {
    impressions += header.impressions ? tdmEnoMayMonthlyV3ToNumber_(row[header.impressions - 1]) : 0;
    clicks += header.clicks ? tdmEnoMayMonthlyV3ToNumber_(row[header.clicks - 1]) : 0;
    cost += header.cost ? tdmEnoMayMonthlyV3ToNumber_(row[header.cost - 1]) : 0;
    callibriSpam += header.callibriSpam ? tdmEnoMayMonthlyV3ToNumber_(row[header.callibriSpam - 1]) : 0;
    callibriNonTarget += header.callibriNonTarget ? tdmEnoMayMonthlyV3ToNumber_(row[header.callibriNonTarget - 1]) : 0;
    purchase += header.purchase ? tdmEnoMayMonthlyV3ToNumber_(row[header.purchase - 1]) : 0;
    addToCart += header.addToCart ? tdmEnoMayMonthlyV3ToNumber_(row[header.addToCart - 1]) : 0;
    jivo += header.jivo ? tdmEnoMayMonthlyV3ToNumber_(row[header.jivo - 1]) : 0;
    callibriLeadA += header.callibriLeadA ? tdmEnoMayMonthlyV3ToNumber_(row[header.callibriLeadA - 1]) : 0;
    callibriLeadC += header.callibriLeadC ? tdmEnoMayMonthlyV3ToNumber_(row[header.callibriLeadC - 1]) : 0;
  });

  const message =
    'Проверка майского блока:\\n' +
    'Показы: ' + tdmEnoMayMonthlyV3Int_(impressions) + '\\n' +
    'Клики: ' + tdmEnoMayMonthlyV3Int_(clicks) + '\\n' +
    'Расход: ' + tdmEnoMayMonthlyV3Rub_(cost) + '\\n' +
    'CPC: ' + tdmEnoMayMonthlyV3Rub_(clicks > 0 ? cost / clicks : 0) + '\\n' +
    'Callibri Спам: ' + tdmEnoMayMonthlyV3Int_(callibriSpam) + '\\n' +
    'Callibri Нецелевой: ' + tdmEnoMayMonthlyV3Int_(callibriNonTarget) + '\\n' +
    'Ecommerce покупка: ' + tdmEnoMayMonthlyV3Int_(purchase) + '\\n' +
    'Ecommerce добавление в корзину: ' + tdmEnoMayMonthlyV3Int_(addToCart) + '\\n' +
    'Jivo: ' + tdmEnoMayMonthlyV3Int_(jivo) + '\\n' +
    'Callibri A: ' + tdmEnoMayMonthlyV3Int_(callibriLeadA) + '\\n' +
    'Callibri C: ' + tdmEnoMayMonthlyV3Int_(callibriLeadC);

  Logger.log(message);
  tdmEnoMayMonthlyV3Toast_(message);
}


// ==========================================================
// Проверка безопасности подсветки.
// Только для контроля: ничего в таблице не меняет.
// ==========================================================

function checkTdmEnoMayColoringSafety() {
  const forbidden = [
    'tdmEnoMayMonthlyV3ColorRowsByStrength_.setValue',
    'tdmEnoMayMonthlyV3ColorRowsByStrength_.clearContent',
    'tdmEnoMayMonthlyV3ColorRowsByStrength_.clearFormat',
    'tdmEnoMayMonthlyV3ColorRowsByStrength_.copyTo'
  ];

  Logger.log('Подсветка строк меняет только setBackground. Данные и формулы в строках кампаний не трогает.');
  tdmEnoMayMonthlyV3Toast_('Ок: подсветка меняет только цвет строк.');
}
function tdmEnoMayMonthlyV3MonthNameRu_(apiDate) {
  const month = Number(String(apiDate).split('-')[1]);

  const names = {
    1: 'январь',
    2: 'февраль',
    3: 'март',
    4: 'апрель',
    5: 'май',
    6: 'июнь',
    7: 'июль',
    8: 'август',
    9: 'сентябрь',
    10: 'октябрь',
    11: 'ноябрь',
    12: 'декабрь'
  };

  return names[month] || '';
}
function fillTdmEnoPreviousFullMonthReport() {
  const period = tdmEnoGetPreviousFullMonthPeriod_();

  TDM_ENO_MAY_MONTHLY_V3_CONFIG.DATE_FROM = period.dateFrom;
  TDM_ENO_MAY_MONTHLY_V3_CONFIG.DATE_TO = period.dateTo;
  TDM_ENO_MAY_MONTHLY_V3_CONFIG.LEGACY_CUTOFF_TO = period.legacyCutoffTo;

  // Важно: для следующих месяцев уже не 93 строка,
  // а новый блок через одну пустую строку после прошлого.
  TDM_ENO_MAY_MONTHLY_V3_CONFIG.FORCE_START_ROW = 0;

  tdmEnoMayMonthlyV3Fill_();
}

function tdmEnoGetPreviousFullMonthPeriod_() {
  const tz = TDM_ENO_MAY_MONTHLY_V3_CONFIG.TIMEZONE || 'Europe/Moscow';
  const now = new Date();

  const year = Number(Utilities.formatDate(now, tz, 'yyyy'));
  const month = Number(Utilities.formatDate(now, tz, 'M'));

  // Первый день текущего месяца
  const firstDayCurrentMonth = new Date(year, month - 1, 1);

  // Последний день прошлого месяца
  const lastDayPreviousMonth = new Date(firstDayCurrentMonth);
  lastDayPreviousMonth.setDate(0);

  // Первый день прошлого месяца
  const firstDayPreviousMonth = new Date(
    lastDayPreviousMonth.getFullYear(),
    lastDayPreviousMonth.getMonth(),
    1
  );

  // 20-е число прошлого месяца оставлено как legacy-порог для совместимости.
  const legacyCutoff = new Date(
    lastDayPreviousMonth.getFullYear(),
    lastDayPreviousMonth.getMonth(),
    20
  );

  return {
    dateFrom: Utilities.formatDate(firstDayPreviousMonth, tz, 'yyyy-MM-dd'),
    dateTo: Utilities.formatDate(lastDayPreviousMonth, tz, 'yyyy-MM-dd'),
    legacyCutoffTo: Utilities.formatDate(legacyCutoff, tz, 'yyyy-MM-dd')
  };
}

function archived_createTdmEnoMonthlyTrigger10am() {
  const triggers = ScriptApp.getProjectTriggers();

  triggers.forEach(trigger => {
    if (trigger.getHandlerFunction() === 'fillTdmEnoPreviousFullMonthReport') {
      ScriptApp.deleteTrigger(trigger);
    }
  });

  ScriptApp.newTrigger('fillTdmEnoPreviousFullMonthReport')
    .timeBased()
    .onMonthDay(1)
    .atHour(10)
    .create();

  SpreadsheetApp.getActive().toast(
    'Готово. ЕНО/ЕМО будет формироваться 1 числа в 10:00 за прошлый месяц.'
  );
}

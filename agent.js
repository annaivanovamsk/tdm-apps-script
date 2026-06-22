/**
 * ТДМ — внутренний агент "План на неделю"
 * Версия 2.0 — реальный анализ по данным из вкладок ЕНО/отчёта.
 *
 * Что делает:
 * 1. Создаёт / обновляет вкладку "ТДМ план на неделю".
 * 2. Берёт данные из всех листов таблицы, кроме служебной вкладки.
 * 3. Ищет кампании, группы, ключи, поисковые запросы, площадки, позиции показа.
 * 4. Подсвечивает перерасход / недорасход, расход без целей, высокий CPC, слабые и сильные элементы.
 * 5. Формирует задачи со статусами.
 * 6. Отправляет письмо со списком задач.
 * 7. Ставит ежедневный триггер на 10:30 по Москве.
 */

const TDM_CONFIG = {
  PLAN_SHEET_NAME: 'ТДМ план на неделю',
  SPREADSHEET_ID: '1zXDDfiRYHJE34iOJg-QfAC0bvoTKMIAuVMfsJS0z5Dg',
  EMAIL_TO: 'kelli837000@gmail.com',
  TIMEZONE: 'Europe/Moscow',
  TRIGGER_FUNCTION: 'tdmDailyMorningAgent',

  // Пока false, чтобы ты видела вкладку при тесте.
  // Когда всё проверим, можно поставить true.
  HIDE_PLAN_SHEET_AFTER_BUILD: false,

  // Фильтр по датам.
  // auto — если в листе есть колонка "Дата", скрипт возьмёт только прошлую неделю.
  // Если даты нет, берёт все строки листа.
  DATE_FILTER_MODE: 'auto',

  // Пороги для рекомендаций.
  MIN_SPEND_NO_GOALS: 500,
  MIN_CLICKS_NO_GOALS: 10,
  HIGH_CPC_MULTIPLIER: 1.5,
  LOW_CTR_SEARCH: 1.0,
  LOW_CTR_NETWORK: 0.2,
  BUDGET_UNDER_PCT: 85,
  BUDGET_OVER_PCT: 105,
  SEARCH_LOW_POSITION: 3.0,
  SEARCH_GOOD_POSITION: 2.0,

  // Макроцели — используем для принятия решений.
  MACRO_GOALS: [
    'звонок',
    'звонки',
    'клик по номеру',
    'клик по номеру телефона',
    'email',
    'клик по email',
    'talkme',
    'talk me',
    'онлайн',
    'чат',
    'покупка',
    'ecommerce: покупка',
    'заказ'
  ],

  // Микроцели — полезны для гипотез, но не заменяют макроцели.
  MICRO_GOALS: [
    'просмотр 3',
    'просмотр 3х страниц',
    '3 страницы',
    'добавление в корзину',
    'корзина',
    'ecommerce: добавление в корзину'
  ],

  STOPWORDS: [
    'купить', 'цена', 'цены', 'стоимость', 'заказать', 'заказ', 'оптом',
    'москва', 'спб', 'санкт', 'петербург', 'россия', 'казахстан',
    'интернет', 'магазин', 'официальный', 'сайт', 'каталог',
    'для', 'или', 'как', 'что', 'это', 'где', 'при', 'под', 'без',
    'болт', 'болты', 'винт', 'винты', 'гайка', 'гайки', 'шайба', 'шайбы'
  ]
};

/**
 * Меню в Google Таблице.
 */
function addTdmWeeklyPlanMenu_() {
  SpreadsheetApp.getUi()
    .createMenu('ТДМ план на неделю')
    .addItem('Сформировать реальный план', 'tdmBuildWeeklyPlanReal')
    .addItem('Отправить письмо с задачами', 'tdmSendTasksEmail')
    .addItem('Сформировать план + отправить письмо', 'tdmDailyMorningAgent')
    .addSeparator()
    .addItem('Поставить триггер каждый день 10:30', 'tdmCreateDailyTrigger')
    .addItem('Удалить триггеры агента', 'tdmDeleteAgentTriggers')
    .addSeparator()
    .addItem('Показать вкладку плана', 'tdmShowPlanSheet')
    .addItem('Скрыть вкладку плана', 'tdmHidePlanSheet')
    .addSeparator()
    .addItem('Обновить отчёт Регионы_города', 'tdmUpdateRegionsCitiesReport')
    .addToUi();
}

/**
 * Главная функция: сформировать реальный план.
 */
function tdmBuildWeeklyPlanReal() {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);

  try {
    const ss = tdmGetSpreadsheet_();
    const periodObj = tdmGetPreviousWeekPeriodObj_();
    const periodText = periodObj.text;
    const today = Utilities.formatDate(new Date(), TDM_CONFIG.TIMEZONE, 'dd.MM.yyyy HH:mm');

    const data = tdmCollectData_(ss, periodObj);
    const analysis = tdmAnalyzeData_(data, periodObj);

    const sheet = tdmGetOrCreatePlanSheet_(ss);
    sheet.clear();
    sheet.showSheet();

    tdmWritePlanSheet_(sheet, periodText, today, data, analysis);
    tdmFormatPlanSheet_(sheet);

    if (TDM_CONFIG.HIDE_PLAN_SHEET_AFTER_BUILD) {
      sheet.hideSheet();
    }

    tdmShowMessage_('Реальный план ТДМ сформирован.');
  } finally {
    lock.releaseLock();
  }
}

/**
 * Для совместимости со старой тестовой функцией.
 */
function tdmBuildWeeklyPlanTest() {
  tdmBuildWeeklyPlanReal();
}

/**
 * Запуск отчёта "Регионы_города" из выпадающего списка Apps Script.
 */
function tdmUpdateRegionsCitiesReport() {
  updateRegionsCitiesReport();
}

function tdmDebugMetrikaConfig() {
  const props = PropertiesService.getScriptProperties();
  return {
    metrikaCounterId: props.getProperty('METRIKA_COUNTER_ID') || '',
    defaultMetrikaCounterId: TDM_REPORTS_CONFIG.defaultMetrikaCounterId || '',
    hasMetrikaToken: Boolean(props.getProperty('METRIKA_TOKEN')),
    hasYandexToken: Boolean(props.getProperty('YANDEX_TOKEN'))
  };
}

function tdmDebugMetrikaAccess() {
  const props = PropertiesService.getScriptProperties();
  const token = props.getProperty('METRIKA_TOKEN');
  const counterId = props.getProperty('METRIKA_COUNTER_ID') || TDM_REPORTS_CONFIG.defaultMetrikaCounterId;
  const statsUrl = TDM_REPORTS_CONFIG.metrikaApiUrl + '?' + tdmReportsToQueryString_({
    ids: counterId,
    date1: '2026-04-01',
    date2: getCurrentAvailableDate_(),
    metrics: 'ym:s:visits',
    accuracy: TDM_REPORTS_CONFIG.accuracy
  });
  const goalsUrl = Utilities.formatString(TDM_REPORTS_CONFIG.metrikaGoalsUrl, counterId);
  const stats = UrlFetchApp.fetch(statsUrl, {
    method: 'get',
    headers: { Authorization: 'OAuth ' + token },
    muteHttpExceptions: true
  });
  const goals = UrlFetchApp.fetch(goalsUrl, {
    method: 'get',
    headers: { Authorization: 'OAuth ' + token },
    muteHttpExceptions: true
  });
  return {
    counterId: counterId,
    statsCode: stats.getResponseCode(),
    statsText: stats.getContentText().slice(0, 300),
    goalsCode: goals.getResponseCode(),
    goalsText: goals.getContentText().slice(0, 300)
  };
}

/**
 * Ежедневный агент:
 * 1. Формирует план.
 * 2. Отправляет письмо.
 */
function tdmDailyMorningAgent() {
  tdmBuildWeeklyPlanReal();
  tdmSendTasksEmail();
}

/**
 * Отправка письма со списком актуальных задач.
 */
function tdmSendTasksEmail() {
  const ss = tdmGetSpreadsheet_();
  let sheet = ss.getSheetByName(TDM_CONFIG.PLAN_SHEET_NAME);

  if (!sheet) {
    tdmBuildWeeklyPlanReal();
    sheet = ss.getSheetByName(TDM_CONFIG.PLAN_SHEET_NAME);
  }

  const values = sheet.getDataRange().getValues();
  const period = String(sheet.getRange('B2').getValue() || '');

  const tasks = tdmReadTasksFromPlan_(values);

  if (!tasks.length) {
    MailApp.sendEmail({
      to: TDM_CONFIG.EMAIL_TO,
      subject: `ТДМ — задач нет за ${period}`,
      body: `ТДМ — план на неделю\n\nПериод анализа: ${period}\n\nАктуальных задач со статусом Новая / В работе / На проверке нет.\n\nСсылка на таблицу:\n${ss.getUrl()}`
    });
    return;
  }

  let body = `ТДМ — план на неделю\n`;
  body += `Период анализа: ${period}\n\n`;
  body += `Задачи:\n\n`;

  tasks.forEach((item, index) => {
    body += `${index + 1}. ${item.task}\n`;
    body += `Статус: ${item.status}\n`;
    body += `Приоритет: ${item.priority}\n`;
    body += `Что сделать: ${item.action}\n`;
    body += `Основание: ${item.reason}\n\n`;
  });

  body += `Ссылка на таблицу:\n${ss.getUrl()}`;

  MailApp.sendEmail({
    to: TDM_CONFIG.EMAIL_TO,
    subject: `ТДМ — план на неделю за ${period}`,
    body: body
  });
}

/**
 * Поставить ежедневный триггер на 10:30 МСК.
 */
function tdmCreateDailyTrigger() {
  tdmDeleteAgentTriggers();

  ScriptApp.newTrigger(TDM_CONFIG.TRIGGER_FUNCTION)
    .timeBased()
    .everyDays(1)
    .atHour(10)
    .nearMinute(30)
    .inTimezone(TDM_CONFIG.TIMEZONE)
    .create();

  tdmShowMessage_('Триггер поставлен: каждый день около 10:30 по Москве.');
}

/**
 * Удалить триггеры агента.
 */
function tdmDeleteAgentTriggers() {
  const triggers = ScriptApp.getProjectTriggers();

  triggers.forEach(trigger => {
    if (trigger.getHandlerFunction() === TDM_CONFIG.TRIGGER_FUNCTION) {
      ScriptApp.deleteTrigger(trigger);
    }
  });

  tdmShowMessage_('Триггеры агента удалены.');
}

function tdmShowPlanSheet() {
  const sheet = tdmGetSpreadsheet_().getSheetByName(TDM_CONFIG.PLAN_SHEET_NAME);
  if (sheet) sheet.showSheet();
}

function tdmHidePlanSheet() {
  const sheet = tdmGetSpreadsheet_().getSheetByName(TDM_CONFIG.PLAN_SHEET_NAME);
  if (sheet) sheet.hideSheet();
}

/**
 * Сбор данных из всех листов.
 */
function tdmCollectData_(ss, periodObj) {
  const result = {
    sources: [],
    campaigns: [],
    groups: [],
    keywords: [],
    queries: [],
    placements: [],
    geo: [],
    types: [],
    budgetPct: null,
    rawDetectedRows: 0
  };

  const sheets = ss.getSheets();

  sheets.forEach(sheet => {
    const name = sheet.getName();

    if (name === TDM_CONFIG.PLAN_SHEET_NAME) return;

    const values = sheet.getDataRange().getValues();
    if (!values || values.length < 2) return;

    const budgetPct = tdmFindBudgetPctInValues_(values);
    if (budgetPct !== null && result.budgetPct === null) {
      result.budgetPct = budgetPct;
    }

    const headerInfo = tdmFindHeaderRow_(values);
    if (!headerInfo) {
      result.sources.push([name, 'Не найден табличный заголовок', 0, 'Пропущен']);
      return;
    }

    const headerRow = headerInfo.rowIndex;
    const headers = headerInfo.headers;
    const map = tdmBuildColumnMap_(headers);
    const sheetType = tdmDetectSheetType_(name, map);

    let rowsCount = 0;

    for (let r = headerRow + 1; r < values.length; r++) {
      const row = values[r];
      if (tdmIsEmptyRow_(row)) continue;

      const parsed = tdmParseMetricRow_(row, map, headers, name, sheetType);

      if (!parsed.hasMetrics && !parsed.object) continue;

      if (!tdmIsRowInPeriod_(parsed, periodObj)) continue;

      rowsCount++;
      result.rawDetectedRows++;

      if (parsed.type === 'campaign') result.campaigns.push(parsed);
      if (parsed.type === 'group') result.groups.push(parsed);
      if (parsed.type === 'keyword') result.keywords.push(parsed);
      if (parsed.type === 'query') result.queries.push(parsed);
      if (parsed.type === 'placement') result.placements.push(parsed);
      if (parsed.type === 'geo') result.geo.push(parsed);
      if (parsed.type === 'campaign_type') result.types.push(parsed);
    }

    result.sources.push([name, sheetType, rowsCount, rowsCount ? 'ОК' : 'Нет строк за период / нет метрик']);
  });

  // Если отдельных кампаний нет, собираем кампании по детальным строкам.
  if (!result.campaigns.length) {
    result.campaigns = tdmBuildCampaignsFromDetails_(
      result.groups.concat(result.keywords, result.queries, result.placements)
    );
  }

  return result;
}

/**
 * Анализ данных и формирование задач.
 */
function tdmAnalyzeData_(data, periodObj) {
  const campaigns = tdmAggregateByObject_(data.campaigns, 'campaign');
  const groups = tdmAggregateByObject_(data.groups, 'group');
  const keywords = tdmAggregateByObject_(data.keywords, 'keyword');
  const queries = tdmAggregateByObject_(data.queries, 'query');
  const placements = tdmAggregateByObject_(data.placements, 'placement');
  const geo = tdmAggregateByObject_(data.geo, 'geo');
  const types = tdmAggregateByObject_(data.types, 'campaignType');

  const allRows = campaigns.concat(groups, keywords, queries, placements, geo, types);
  const totals = tdmCalcTotals_(campaigns.length ? campaigns : allRows);

  const avgCpc = totals.clicks ? totals.cost / totals.clicks : 0;
  const avgCpa = totals.macroGoals ? totals.cost / totals.macroGoals : 0;

  const campaignSignals = tdmAnalyzeLevel_(campaigns, 'Кампания', avgCpc, avgCpa);
  const groupSignals = tdmAnalyzeLevel_(groups, 'Группа', avgCpc, avgCpa);
  const keywordSignals = tdmAnalyzeLevel_(keywords, 'Ключ', avgCpc, avgCpa);
  const querySignals = tdmAnalyzeQueries_(queries, avgCpc);
  const placementSignals = tdmAnalyzePlacements_(placements, avgCpc);
  const positionSignals = tdmAnalyzePositions_(data.keywords.concat(data.groups, data.campaigns), avgCpc, avgCpa);
  const budgetSignals = tdmAnalyzeBudget_(data.budgetPct, totals);
  const growthIdeas = tdmBuildGrowthIdeas_(data, campaigns, geo, types, queries, totals);

  const tasks = tdmBuildTasks_({
    data,
    totals,
    avgCpc,
    avgCpa,
    campaignSignals,
    groupSignals,
    keywordSignals,
    querySignals,
    placementSignals,
    positionSignals,
    budgetSignals,
    growthIdeas
  });

  return {
    totals,
    avgCpc,
    avgCpa,
    tasks,
    budgetSignals,
    campaignSignals,
    groupSignals,
    keywordSignals,
    querySignals,
    placementSignals,
    positionSignals,
    growthIdeas,
    sources: data.sources
  };
}

/**
 * Запись плана на лист.
 */
function tdmWritePlanSheet_(sheet, periodText, today, data, analysis) {
  let row = 1;

  sheet.getRange(row, 1).setValue('ТДМ — план на неделю');
  row++;

  sheet.getRange(row, 1).setValue('Период анализа:');
  sheet.getRange(row, 2).setValue(periodText);
  row++;

  sheet.getRange(row, 1).setValue('Дата формирования:');
  sheet.getRange(row, 2).setValue(today);
  row++;

  sheet.getRange(row, 1).setValue('Строк найдено в данных:');
  sheet.getRange(row, 2).setValue(data.rawDetectedRows);
  row += 2;

  row = tdmWriteTable_(sheet, row, 'Задачи на неделю', [
    'Дата создания',
    'Период',
    'Задача',
    'Основание',
    'Что сделать',
    'Приоритет',
    'Статус',
    'Комментарий'
  ], analysis.tasks);

  row += 2;

  row = tdmWriteTable_(sheet, row, 'Бюджет / общий контроль', [
    'Сигнал',
    'Что заметили',
    'Рекомендация',
    'Приоритет',
    'Статус'
  ], analysis.budgetSignals);

  row += 2;

  row = tdmWriteTable_(sheet, row, 'Кампании: сильные и слабые зоны', [
    'Уровень',
    'Объект',
    'Расход',
    'Клики',
    'CPC',
    'CTR',
    'Макроцели',
    'CPA',
    'Сигнал',
    'Рекомендация',
    'Приоритет',
    'Статус'
  ], analysis.campaignSignals);

  row += 2;

  row = tdmWriteTable_(sheet, row, 'Группы: сильные и слабые зоны', [
    'Уровень',
    'Объект',
    'Расход',
    'Клики',
    'CPC',
    'CTR',
    'Макроцели',
    'CPA',
    'Сигнал',
    'Рекомендация',
    'Приоритет',
    'Статус'
  ], analysis.groupSignals);

  row += 2;

  row = tdmWriteTable_(sheet, row, 'Ключи: сильные и слабые зоны', [
    'Уровень',
    'Объект',
    'Расход',
    'Клики',
    'CPC',
    'CTR',
    'Макроцели',
    'CPA',
    'Сигнал',
    'Рекомендация',
    'Приоритет',
    'Статус'
  ], analysis.keywordSignals);

  row += 2;

  row = tdmWriteTable_(sheet, row, 'Минус-слова / поисковые запросы на проверку', [
    'Запрос',
    'Кампания',
    'Расход',
    'Клики',
    'CPC',
    'Макроцели',
    'Сигнал',
    'Рекомендация',
    'Приоритет',
    'Статус'
  ], analysis.querySignals);

  row += 2;

  row = tdmWriteTable_(sheet, row, 'Площадки на проверку', [
    'Площадка',
    'Кампания',
    'Расход',
    'Клики',
    'CPC',
    'Макроцели',
    'Сигнал',
    'Рекомендация',
    'Приоритет',
    'Статус'
  ], analysis.placementSignals);

  row += 2;

  row = tdmWriteTable_(sheet, row, 'Позиция показа на Поиске', [
    'Уровень',
    'Объект',
    'Кампания',
    'Позиция',
    'Расход',
    'Клики',
    'CPC',
    'Макроцели',
    'Сигнал',
    'Рекомендация',
    'Приоритет',
    'Статус'
  ], analysis.positionSignals);

  row += 2;

  row = tdmWriteTable_(sheet, row, 'Идеи новых кампаний / типов РК', [
    'Что заметили',
    'Рекомендация',
    'Тип РК',
    'Приоритет',
    'Статус'
  ], analysis.growthIdeas);

  row += 2;

  row = tdmWriteTable_(sheet, row, 'Диагностика источников данных', [
    'Лист',
    'Тип данных',
    'Строк обработано',
    'Статус'
  ], analysis.sources);
}

/**
 * Форматирование листа.
 */
function tdmFormatPlanSheet_(sheet) {
  const lastRow = Math.max(sheet.getLastRow(), 1);
  const lastCol = Math.max(sheet.getLastColumn(), 1);

  sheet.getRange(1, 1, lastRow, lastCol)
    .setFontFamily('Arial')
    .setFontSize(10)
    .setVerticalAlignment('top');

  sheet.getRange('A1:B4').setFontWeight('bold');

  sheet.autoResizeColumns(1, Math.min(lastCol, 12));

  for (let col = 1; col <= lastCol; col++) {
    const width = sheet.getColumnWidth(col);
    if (width > 360) sheet.setColumnWidth(col, 360);
    if (width < 90) sheet.setColumnWidth(col, 90);
  }

  sheet.setFrozenRows(5);

  // Перед новыми выпадающими списками убираем старые проверки данных.
  // Иначе Google Sheets может ругаться на старые значения в ячейках статусов.
  sheet.getRange(1, 1, lastRow, lastCol).clearDataValidations();

  // Статусы — выпадающие списки по всему листу, где есть колонка "Статус".
  tdmApplyStatusValidations_(sheet);

  // Условное форматирование.
  const rules = [];

  rules.push(
    SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo('Высокий')
      .setBackground('#f4cccc')
      .setRanges([sheet.getDataRange()])
      .build()
  );

  rules.push(
    SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo('Средний')
      .setBackground('#fff2cc')
      .setRanges([sheet.getDataRange()])
      .build()
  );

  rules.push(
    SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo('Выполнено')
      .setBackground('#d9ead3')
      .setRanges([sheet.getDataRange()])
      .build()
  );

  rules.push(
    SpreadsheetApp.newConditionalFormatRule()
      .whenTextContains('Расход без целей')
      .setBackground('#f4cccc')
      .setRanges([sheet.getDataRange()])
      .build()
  );

  rules.push(
    SpreadsheetApp.newConditionalFormatRule()
      .whenTextContains('Точка роста')
      .setBackground('#d9ead3')
      .setRanges([sheet.getDataRange()])
      .build()
  );

  sheet.setConditionalFormatRules(rules);
}

/**
 * Создание таблицы с заголовком секции.
 */
function tdmWriteTable_(sheet, startRow, title, headers, rows) {
  const safeRows = rows && rows.length ? rows : [['Нет данных для блока']];

  sheet.getRange(startRow, 1).setValue(title).setFontWeight('bold').setBackground('#cfe2f3');
  startRow++;

  sheet.getRange(startRow, 1, 1, headers.length)
    .setValues([headers])
    .setFontWeight('bold')
    .setBackground('#d9ead3');

  startRow++;

  if (safeRows.length && safeRows[0].length === 1) {
    sheet.getRange(startRow, 1).setValue(safeRows[0][0]);
    return startRow + 1;
  }

  sheet.getRange(startRow, 1, safeRows.length, headers.length).setValues(safeRows);

  return startRow + safeRows.length;
}

/**
 * Анализ уровня: кампании / группы / ключи.
 */
function tdmAnalyzeLevel_(rows, levelName, avgCpc, avgCpa) {
  const result = [];

  rows.forEach(item => {
    const spend = item.cost || 0;
    const clicks = item.clicks || 0;
    const cpc = item.cpc || (clicks ? spend / clicks : 0);
    const ctr = item.ctr || 0;
    const goals = item.macroGoals || 0;
    const cpa = goals ? spend / goals : 0;

    let signal = '';
    let recommendation = '';
    let priority = 'Средний';

    if (spend >= TDM_CONFIG.MIN_SPEND_NO_GOALS && goals === 0) {
      signal = 'Расход без целей';
      recommendation = 'Проверить трафик: запросы / площадки / объявления / посадочную. Пессимизировать после проверки.';
      priority = 'Высокий';
    } else if (goals > 0 && (!avgCpa || cpa <= avgCpa)) {
      signal = 'Точка роста';
      recommendation = 'Оставить в приоритете. Проверить возможность аккуратного усиления бюджета / ставок.';
      priority = 'Средний';
    } else if (goals > 0 && avgCpa && cpa > avgCpa * 1.5) {
      signal = 'Цели есть, но CPA выше среднего';
      recommendation = 'Проверить ставки, запросы, объявления и посадочную. Усиливать осторожно.';
      priority = 'Средний';
    } else if (cpc && avgCpc && cpc > avgCpc * TDM_CONFIG.HIGH_CPC_MULTIPLIER && goals === 0) {
      signal = 'Высокий CPC без целей';
      recommendation = 'Проверить максимальную цену клика, стратегию и целесообразность ставки.';
      priority = 'Высокий';
    } else if (clicks >= TDM_CONFIG.MIN_CLICKS_NO_GOALS && goals === 0) {
      signal = 'Клики есть, целей нет';
      recommendation = 'Оставить на проверку: качество трафика, релевантность, посадочная.';
      priority = 'Средний';
    } else {
      return;
    }

    result.push([
      levelName,
      item.object,
      tdmMoney_(spend),
      clicks,
      tdmMoney_(cpc),
      tdmPct_(ctr),
      goals,
      goals ? tdmMoney_(cpa) : '',
      signal,
      recommendation,
      priority,
      'Новая'
    ]);
  });

  return result.slice(0, 50);
}

/**
 * Анализ поисковых запросов.
 */
function tdmAnalyzeQueries_(queries, avgCpc) {
  const result = [];

  queries.forEach(item => {
    const spend = item.cost || 0;
    const clicks = item.clicks || 0;
    const cpc = item.cpc || (clicks ? spend / clicks : 0);
    const goals = item.macroGoals || 0;

    if ((spend >= TDM_CONFIG.MIN_SPEND_NO_GOALS || clicks >= TDM_CONFIG.MIN_CLICKS_NO_GOALS) && goals === 0) {
      result.push([
        item.object,
        item.campaign || '',
        tdmMoney_(spend),
        clicks,
        tdmMoney_(cpc),
        goals,
        'Запрос с расходом без целей',
        'Проверить на минус-слова / уточнение соответствия / перенос в отдельную группу.',
        'Высокий',
        'Новая'
      ]);
    } else if (goals > 0) {
      result.push([
        item.object,
        item.campaign || '',
        tdmMoney_(spend),
        clicks,
        tdmMoney_(cpc),
        goals,
        'Рабочий запрос',
        'Оставить. Проверить, можно ли усилить или вынести в отдельную группу.',
        'Средний',
        'Новая'
      ]);
    }
  });

  return result.slice(0, 80);
}

/**
 * Анализ площадок.
 */
function tdmAnalyzePlacements_(placements, avgCpc) {
  const result = [];

  placements.forEach(item => {
    const spend = item.cost || 0;
    const clicks = item.clicks || 0;
    const cpc = item.cpc || (clicks ? spend / clicks : 0);
    const goals = item.macroGoals || 0;

    if ((spend >= TDM_CONFIG.MIN_SPEND_NO_GOALS || clicks >= TDM_CONFIG.MIN_CLICKS_NO_GOALS) && goals === 0) {
      result.push([
        item.object,
        item.campaign || '',
        tdmMoney_(spend),
        clicks,
        tdmMoney_(cpc),
        goals,
        'Площадка с расходом без целей',
        'Проверить качество площадки. После ручной проверки — исключить / снизить долю показов.',
        'Высокий',
        'Новая'
      ]);
    } else if (goals > 0) {
      result.push([
        item.object,
        item.campaign || '',
        tdmMoney_(spend),
        clicks,
        tdmMoney_(cpc),
        goals,
        'Площадка дала цели',
        'Оставить. Проверить, можно ли сохранить трафик с этой площадки.',
        'Средний',
        'Новая'
      ]);
    }
  });

  return result.slice(0, 80);
}

/**
 * Анализ позиции показа.
 */
function tdmAnalyzePositions_(rows, avgCpc, avgCpa) {
  const result = [];

  rows.forEach(item => {
    const position = item.position;
    if (!position) return;

    const spend = item.cost || 0;
    const clicks = item.clicks || 0;
    const cpc = item.cpc || (clicks ? spend / clicks : 0);
    const goals = item.macroGoals || 0;

    let signal = '';
    let recommendation = '';
    let priority = 'Средний';

    if (position > TDM_CONFIG.SEARCH_LOW_POSITION && goals > 0) {
      signal = 'Низкая позиция, но цели есть';
      recommendation = 'Проверить возможность аккуратно поднять ставку / ценность конверсии.';
      priority = 'Высокий';
    } else if (position <= TDM_CONFIG.SEARCH_GOOD_POSITION && spend >= TDM_CONFIG.MIN_SPEND_NO_GOALS && goals === 0) {
      signal = 'Высокая позиция без целей';
      recommendation = 'Не масштабировать. Проверить запросы, объявление, посадочную и ставку.';
      priority = 'Высокий';
    } else if (position > TDM_CONFIG.SEARCH_LOW_POSITION && goals === 0) {
      signal = 'Низкая позиция без целей';
      recommendation = 'Сначала проверить качество запроса. Ставку поднимать только после проверки.';
      priority = 'Средний';
    } else if (position <= TDM_CONFIG.SEARCH_GOOD_POSITION && goals > 0) {
      signal = 'Хорошая позиция и есть цели';
      recommendation = 'Оставить в приоритете. Проверить возможность масштабирования.';
      priority = 'Средний';
    } else {
      return;
    }

    result.push([
      item.level || item.type || '',
      item.object || '',
      item.campaign || '',
      position,
      tdmMoney_(spend),
      clicks,
      tdmMoney_(cpc),
      goals,
      signal,
      recommendation,
      priority,
      'Новая'
    ]);
  });

  return result.slice(0, 80);
}

/**
 * Анализ бюджета.
 */
function tdmAnalyzeBudget_(budgetPct, totals) {
  const result = [];

  if (budgetPct !== null) {
    if (budgetPct < TDM_CONFIG.BUDGET_UNDER_PCT) {
      result.push([
        'Недорасход',
        `Идём по бюджету на ${tdmPct_(budgetPct)} от плана`,
        'Проверить охват, ограничения, ставки, дневные лимиты и доступный объём трафика.',
        'Высокий',
        'Новая'
      ]);
    } else if (budgetPct > TDM_CONFIG.BUDGET_OVER_PCT) {
      result.push([
        'Перерасход',
        `Идём по бюджету на ${tdmPct_(budgetPct)} от плана`,
        'Проверить дневные лимиты, перераспределение бюджета и кампании с расходом без целей.',
        'Высокий',
        'Новая'
      ]);
    } else {
      result.push([
        'Бюджет в норме',
        `Идём по бюджету на ${tdmPct_(budgetPct)} от плана`,
        'Сохранять контроль бюджета и перераспределять средства по эффективности кампаний.',
        'Средний',
        'Новая'
      ]);
    }
  } else {
    result.push([
      'План бюджета не найден',
      `Факт расхода по найденным данным: ${tdmMoney_(totals.cost)}`,
      'Если нужен контроль перерасхода / недорасхода, нужно добавить в отчёт строку с % выполнения бюджета или отдельную ячейку с планом.',
      'Средний',
      'Новая'
    ]);
  }

  if (totals.macroGoals === 0 && totals.cost > 0) {
    result.push([
      'Нет макроцелей',
      `Расход есть: ${tdmMoney_(totals.cost)}, макроцелей не найдено`,
      'Проверить цели, коллтрекинг, TalkMe, email, формы и корректность передачи данных.',
      'Высокий',
      'Новая'
    ]);
  }

  return result;
}

/**
 * Идеи новых кампаний / типов РК.
 */
function tdmBuildGrowthIdeas_(data, campaigns, geoRows, typeRows, queries, totals) {
  const result = [];

  // Гео.
  geoRows.forEach(item => {
    const cpc = item.cpc || (item.clicks ? item.cost / item.clicks : 0);
    if (item.clicks >= 50 && cpc && totals.clicks && cpc < (totals.cost / totals.clicks) && item.macroGoals > 0) {
      result.push([
        `Гео "${item.object}" даёт трафик дешевле среднего и есть цели`,
        'Рассмотреть отдельную кампанию / отдельную стратегию под это гео.',
        'Поиск / товарная / ЕПК по гео',
        'Средний',
        'Новая'
      ]);
    }
  });

  // Типы РК.
  typeRows.forEach(item => {
    const lower = String(item.object || '').toLowerCase();

    if ((lower.includes('товар') || lower.includes('ecom') || lower.includes('епк')) && item.clicks >= 50 && (item.macroGoals > 0 || item.microGoals > 0)) {
      result.push([
        `Тип РК "${item.object}" даёт трафик и целевые / микроцелевые действия`,
        'Проверить сильные категории. Возможно, вынести их в отдельную товарную / ЕПК кампанию.',
        'Товарная / ЕПК',
        'Средний',
        'Новая'
      ]);
    }

    if ((lower.includes('сети') || lower.includes('рся')) && item.microGoals > 20 && item.macroGoals === 0) {
      result.push([
        `В сетях есть вовлечённость, но нет макроцелей`,
        'Проверить ретаргетинг на вовлечённых пользователей и корректность посадочных.',
        'РСЯ / ретаргетинг',
        'Средний',
        'Новая'
      ]);
    }
  });

  // Частотные темы в запросах.
  const terms = tdmGetFrequentTermsFromQueries_(queries);

  terms.slice(0, 5).forEach(term => {
    result.push([
      `В запросах часто встречается тема "${term.term}"`,
      'Проверить, есть ли под это отдельная группа / кампания. Если нет — рассмотреть тест.',
      'Поиск',
      term.goals > 0 ? 'Высокий' : 'Средний',
      'Новая'
    ]);
  });

  // Если микроцелей много, макро мало.
  if (totals.microGoals >= 20 && totals.macroGoals === 0) {
    result.push([
      'Есть микроцели, но нет макрообращений',
      'Проверить ретаргетинг, формы, звонки, email, TalkMe и путь до обращения.',
      'Ретаргетинг / UX-гипотеза',
      'Высокий',
      'Новая'
    ]);
  }

  if (!result.length) {
    result.push([
      'Явных идей по новым РК по текущим данным не найдено',
      'Оставить блок под наблюдением. После накопления данных проверить повторно.',
      '—',
      'Средний',
      'Новая'
    ]);
  }

  return result.slice(0, 20);
}

/**
 * Формирование задач на неделю.
 */
function tdmBuildTasks_(ctx) {
  const today = Utilities.formatDate(new Date(), TDM_CONFIG.TIMEZONE, 'dd.MM.yyyy');
  const period = tdmGetPreviousWeekPeriodObj_().text;

  const tasks = [];

  function add(task, reason, action, priority) {
    tasks.push([today, period, task, reason, action, priority, 'Новая', '']);
  }

  add(
    'Чистка площадок',
    ctx.placementSignals.length ? 'Найдены площадки с расходом без целей / площадки на проверку' : 'Проверить площадки за прошлую неделю',
    'Проверить площадки и исключить слабые после ручной проверки',
    ctx.placementSignals.length ? 'Высокий' : 'Средний'
  );

  add(
    'Минус-слова',
    ctx.querySignals.length ? 'Найдены поисковые запросы с расходом без целей / рабочие запросы' : 'Проверить поисковые запросы за прошлую неделю',
    'Проверить поисковые запросы и добавить минус-фразы',
    ctx.querySignals.length ? 'Высокий' : 'Средний'
  );

  add(
    'Статистика по группам',
    ctx.groupSignals.length ? 'Есть сильные / слабые группы' : 'Нужно проверить статистику по группам',
    'Подсветить сильные и слабые группы, дать рекомендации',
    'Высокий'
  );

  add(
    'Статистика по ключам',
    ctx.keywordSignals.length ? 'Есть сильные / слабые ключи' : 'Нужно проверить статистику по ключам',
    'Подсветить сильные и слабые ключи, дать рекомендации',
    'Высокий'
  );

  add(
    'Позиция показа на Поиске',
    ctx.positionSignals.length ? 'Найдены сигналы по позиции показа' : 'Проверить позицию показа по поисковым РК',
    'Проверить, где нужно поднять / снизить ставки с учётом целей и CPC',
    'Высокий'
  );

  add(
    'Контроль бюджета',
    ctx.budgetSignals.length ? ctx.budgetSignals[0][0] : 'Проверить перерасход / недорасход',
    'Сравнить факт с планом и проверить дневные лимиты',
    'Средний'
  );

  add(
    'Перераспределение бюджета',
    ctx.campaignSignals.length ? 'Есть кампании сильнее и слабее по цене цели / расходу без целей' : 'Проверить распределение бюджета между РК',
    'Перенести фокус в кампании с более выгодной ценой обращения',
    'Средний'
  );

  add(
    'Идеи новых кампаний / типов РК',
    ctx.growthIdeas.length ? 'Найдены гипотезы для роста / проверки' : 'Проверить точки роста',
    'Проверить, есть ли смысл добавить новую РК, ЕПК, товарную, РСЯ или ретаргетинг',
    'Средний'
  );

  add(
    'Проверка новых кампаний',
    'Если появилась новая РК, она должна попасть во все блоки анализа',
    'Проверить, что новая кампания есть в отчёте, группах, ключах и задачах',
    'Высокий'
  );

  if (!ctx.data.rawDetectedRows) {
    add(
      'Проверить подключение данных',
      'Скрипт не нашёл строки с метриками в листах',
      'Проверить названия колонок и структуру вкладок ЕНО. Возможно, нужно настроить маппинг колонок.',
      'Высокий'
    );
  }

  return tasks;
}

/**
 * Поиск строки заголовков.
 */
function tdmFindHeaderRow_(values) {
  const maxRows = Math.min(values.length, 30);
  let best = null;

  for (let r = 0; r < maxRows; r++) {
    const row = values[r].map(v => tdmNorm_(v));
    let score = 0;

    row.forEach(cell => {
      if (!cell) return;
      if (tdmHeaderMatch_(cell, ['кампания', 'campaign'])) score += 2;
      if (tdmHeaderMatch_(cell, ['группа', 'ad group'])) score += 2;
      if (tdmHeaderMatch_(cell, ['ключ', 'keyword', 'фраза'])) score += 2;
      if (tdmHeaderMatch_(cell, ['поисковый запрос', 'запрос', 'search query'])) score += 2;
      if (tdmHeaderMatch_(cell, ['площадка', 'место показа', 'placement', 'домен'])) score += 2;
      if (tdmHeaderMatch_(cell, ['расход', 'cost', 'spend'])) score += 2;
      if (tdmHeaderMatch_(cell, ['клики', 'clicks'])) score += 1;
      if (tdmHeaderMatch_(cell, ['показы', 'impressions'])) score += 1;
      if (tdmHeaderMatch_(cell, ['cpc', 'цена клика'])) score += 1;
      if (tdmHeaderMatch_(cell, ['ctr'])) score += 1;
      if (tdmHeaderMatch_(cell, ['позиция'])) score += 1;
    });

    if (!best || score > best.score) {
      best = { rowIndex: r, headers: values[r], score: score };
    }
  }

  if (best && best.score >= 3) return best;
  return null;
}

/**
 * Маппинг колонок.
 */
function tdmBuildColumnMap_(headers) {
  const map = {
    goalCols: [],
    microGoalCols: []
  };

  headers.forEach((header, index) => {
    const h = tdmNorm_(header);

    if (!h) return;

    if (!map.date && tdmHeaderMatch_(h, ['дата', 'date', 'день'])) map.date = index;
    if (!map.campaign && tdmHeaderMatch_(h, ['кампания', 'campaign', 'рк'])) map.campaign = index;
    if (!map.group && tdmHeaderMatch_(h, ['группа', 'группа объявлений', 'ad group', 'adgroup'])) map.group = index;
    if (!map.keyword && tdmHeaderMatch_(h, ['ключ', 'ключевая фраза', 'keyword', 'фраза'])) map.keyword = index;
    if (!map.query && tdmHeaderMatch_(h, ['поисковый запрос', 'search query', 'search term'])) map.query = index;
    if (!map.placement && tdmHeaderMatch_(h, ['площадка', 'место показа', 'placement', 'домен', 'сайт'])) map.placement = index;
    if (!map.geo && tdmHeaderMatch_(h, ['гео', 'регион', 'страна', 'город', 'location'])) map.geo = index;
    if (!map.campaignType && tdmHeaderMatch_(h, ['тип рк', 'тип кампании', 'тип размещения', 'тип'])) map.campaignType = index;

    if (!map.impressions && tdmHeaderMatch_(h, ['показы', 'impressions'])) map.impressions = index;
    if (!map.clicks && tdmHeaderMatch_(h, ['клики', 'clicks'])) map.clicks = index;
    if (!map.cost && tdmHeaderMatch_(h, ['расход', 'cost', 'spend', 'стоимость'])) map.cost = index;
    if (!map.ctr && tdmHeaderMatch_(h, ['ctr'])) map.ctr = index;
    if (!map.cpc && tdmHeaderMatch_(h, ['cpc', 'средний cpc', 'ср cpc', 'средняя цена клика', 'цена клика'])) map.cpc = index;
    if (!map.conversions && tdmHeaderMatch_(h, ['конверсии', 'конверсия', 'цели', 'достижения цели'])) map.conversions = index;
    if (!map.cpa && tdmHeaderMatch_(h, ['cpa', 'цена цели', 'стоимость цели', 'цена конверсии'])) map.cpa = index;
    if (!map.position && tdmHeaderMatch_(h, ['позиция показа', 'средняя позиция', 'ср позиция', 'позиция', 'avg position'])) map.position = index;

    if (tdmMatchesGoal_(h, TDM_CONFIG.MACRO_GOALS)) map.goalCols.push(index);
    if (tdmMatchesGoal_(h, TDM_CONFIG.MICRO_GOALS)) map.microGoalCols.push(index);
  });

  return map;
}

/**
 * Определяем тип листа.
 */
function tdmDetectSheetType_(sheetName, map) {
  const name = tdmNorm_(sheetName);

  if (map.query || name.includes('запрос')) return 'query';
  if (map.placement || name.includes('площад') || name.includes('placement')) return 'placement';
  if (map.keyword || name.includes('ключ')) return 'keyword';
  if (map.group || name.includes('групп')) return 'group';
  if (map.geo || name.includes('гео') || name.includes('регион')) return 'geo';

  if (map.campaignType || name.includes('тип')) return 'campaign_type';
  if (map.campaign || name.includes('кампан') || name.includes('рк')) return 'campaign';

  return 'unknown';
}

/**
 * Парсим строку метрик.
 */
function tdmParseMetricRow_(row, map, headers, sheetName, sheetType) {
  const object = tdmGetObject_(row, map, sheetType);
  const campaign = map.campaign !== undefined ? String(row[map.campaign] || '').trim() : '';
  const date = map.date !== undefined ? tdmParseDate_(row[map.date]) : null;

  const impressions = map.impressions !== undefined ? tdmNum_(row[map.impressions]) : 0;
  const clicks = map.clicks !== undefined ? tdmNum_(row[map.clicks]) : 0;
  const cost = map.cost !== undefined ? tdmNum_(row[map.cost]) : 0;
  const ctr = map.ctr !== undefined ? tdmNum_(row[map.ctr]) : 0;
  const cpc = map.cpc !== undefined ? tdmNum_(row[map.cpc]) : (clicks ? cost / clicks : 0);
  const conversions = map.conversions !== undefined ? tdmNum_(row[map.conversions]) : 0;
  const cpa = map.cpa !== undefined ? tdmNum_(row[map.cpa]) : 0;
  const position = map.position !== undefined ? tdmNum_(row[map.position]) : 0;

  const macroGoals = map.goalCols.length
    ? map.goalCols.reduce((sum, i) => sum + tdmNum_(row[i]), 0)
    : conversions;

  const microGoals = map.microGoalCols.length
    ? map.microGoalCols.reduce((sum, i) => sum + tdmNum_(row[i]), 0)
    : 0;

  const type = tdmNormalizeType_(sheetType, map);

  return {
    sheetName,
    type,
    level: type,
    object,
    campaign,
    date,
    impressions,
    clicks,
    cost,
    ctr,
    cpc,
    conversions,
    cpa,
    macroGoals,
    microGoals,
    position,
    hasMetrics: Boolean(impressions || clicks || cost || ctr || cpc || conversions || macroGoals || microGoals || position)
  };
}

/**
 * Получаем объект строки в зависимости от типа.
 */
function tdmGetObject_(row, map, sheetType) {
  if (sheetType === 'query' && map.query !== undefined) return String(row[map.query] || '').trim();
  if (sheetType === 'placement' && map.placement !== undefined) return String(row[map.placement] || '').trim();
  if (sheetType === 'keyword' && map.keyword !== undefined) return String(row[map.keyword] || '').trim();
  if (sheetType === 'group' && map.group !== undefined) return String(row[map.group] || '').trim();
  if (sheetType === 'geo' && map.geo !== undefined) return String(row[map.geo] || '').trim();
  if (sheetType === 'campaign_type' && map.campaignType !== undefined) return String(row[map.campaignType] || '').trim();
  if (map.campaign !== undefined) return String(row[map.campaign] || '').trim();
  return '';
}

function tdmNormalizeType_(sheetType, map) {
  if (sheetType === 'query') return 'query';
  if (sheetType === 'placement') return 'placement';
  if (sheetType === 'keyword') return 'keyword';
  if (sheetType === 'group') return 'group';
  if (sheetType === 'geo') return 'geo';
  if (sheetType === 'campaign_type') return 'campaign_type';
  return 'campaign';
}

/**
 * Фильтр по периоду.
 */
function tdmIsRowInPeriod_(parsed, periodObj) {
  if (TDM_CONFIG.DATE_FILTER_MODE !== 'auto') return true;
  if (!parsed.date) return true;

  const d = new Date(parsed.date.getFullYear(), parsed.date.getMonth(), parsed.date.getDate());
  return d >= periodObj.start && d <= periodObj.end;
}

/**
 * Агрегация по объекту.
 */
function tdmAggregateByObject_(rows, fieldName) {
  const map = {};

  rows.forEach(row => {
    const key = row.object || row.campaign || '';
    if (!key) return;

    if (!map[key]) {
      map[key] = {
        object: key,
        campaign: row.campaign || '',
        type: row.type,
        level: row.level,
        impressions: 0,
        clicks: 0,
        cost: 0,
        macroGoals: 0,
        microGoals: 0,
        positionSum: 0,
        positionCount: 0
      };
    }

    map[key].impressions += row.impressions || 0;
    map[key].clicks += row.clicks || 0;
    map[key].cost += row.cost || 0;
    map[key].macroGoals += row.macroGoals || 0;
    map[key].microGoals += row.microGoals || 0;

    if (row.position) {
      map[key].positionSum += row.position;
      map[key].positionCount++;
    }
  });

  return Object.keys(map).map(key => {
    const item = map[key];
    item.cpc = item.clicks ? item.cost / item.clicks : 0;
    item.ctr = item.impressions ? item.clicks / item.impressions * 100 : 0;
    item.cpa = item.macroGoals ? item.cost / item.macroGoals : 0;
    item.position = item.positionCount ? item.positionSum / item.positionCount : 0;
    return item;
  }).sort((a, b) => b.cost - a.cost);
}

/**
 * Если нет отдельной статистики кампаний, строим её из детальных строк.
 */
function tdmBuildCampaignsFromDetails_(rows) {
  const converted = rows
    .filter(row => row.campaign)
    .map(row => {
      const copy = Object.assign({}, row);
      copy.object = row.campaign;
      copy.type = 'campaign';
      copy.level = 'campaign';
      return copy;
    });

  return tdmAggregateByObject_(converted, 'campaign');
}

/**
 * Общие итоги.
 */
function tdmCalcTotals_(rows) {
  return rows.reduce((acc, row) => {
    acc.impressions += row.impressions || 0;
    acc.clicks += row.clicks || 0;
    acc.cost += row.cost || 0;
    acc.macroGoals += row.macroGoals || 0;
    acc.microGoals += row.microGoals || 0;
    return acc;
  }, {
    impressions: 0,
    clicks: 0,
    cost: 0,
    macroGoals: 0,
    microGoals: 0
  });
}

/**
 * Частотные темы из поисковых запросов.
 */
function tdmGetFrequentTermsFromQueries_(queries) {
  const map = {};

  queries.forEach(row => {
    const query = String(row.object || '').toLowerCase();
    const words = query
      .replace(/[^\p{L}\p{N}\s-]/gu, ' ')
      .split(/\s+/)
      .filter(w => w.length >= 4)
      .filter(w => !TDM_CONFIG.STOPWORDS.includes(w));

    words.forEach(word => {
      if (!map[word]) {
        map[word] = { term: word, count: 0, clicks: 0, cost: 0, goals: 0 };
      }
      map[word].count++;
      map[word].clicks += row.clicks || 0;
      map[word].cost += row.cost || 0;
      map[word].goals += row.macroGoals || 0;
    });
  });

  return Object.keys(map)
    .map(k => map[k])
    .filter(item => item.count >= 2 || item.goals > 0 || item.clicks >= 20)
    .sort((a, b) => (b.goals - a.goals) || (b.clicks - a.clicks) || (b.count - a.count));
}

/**
 * Чтение задач из листа.
 */
function tdmReadTasksFromPlan_(values) {
  const tasks = [];
  let headerRow = -1;
  let map = {};

  for (let i = 0; i < values.length; i++) {
    const row = values[i].map(v => String(v || '').trim());
    if (row.includes('Задача') && row.includes('Статус')) {
      headerRow = i;
      row.forEach((h, idx) => map[h] = idx);
      break;
    }
  }

  if (headerRow === -1) return tasks;

  for (let i = headerRow + 1; i < values.length; i++) {
    const row = values[i];

    if (!row[map['Задача']]) break;

    const status = String(row[map['Статус']] || '');
    if (status === 'Выполнено' || status === 'Не актуально') continue;

    tasks.push({
      task: row[map['Задача']] || '',
      reason: row[map['Основание']] || '',
      action: row[map['Что сделать']] || '',
      priority: row[map['Приоритет']] || '',
      status: status || ''
    });
  }

  return tasks;
}

/**
 * Статусы в выпадающих списках.
 */
function tdmApplyStatusValidations_(sheet) {
  const values = sheet.getDataRange().getValues();
  const allowedStatuses = [
    'Новая',
    'В работе',
    'На проверке',
    'Выполнено',
    'Отложено',
    'Не актуально'
  ];

  const statusRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(allowedStatuses, true)
    .setAllowInvalid(true)
    .build();

  for (let r = 0; r < values.length; r++) {
    const headerRow = values[r].map(v => String(v || '').trim());

    for (let c = 0; c < headerRow.length; c++) {
      if (headerRow[c] !== 'Статус') continue;

      // Выпадающий список статусов ставим только на рабочие таблицы.
      // Например, блок диагностики тоже имеет колонку «Статус», но там значения «ОК» / «Нет строк»,
      // поэтому туда проверку данных ставить не нужно.
      const isTaskTable = headerRow.includes('Приоритет') || headerRow.includes('Задача') || headerRow.includes('Рекомендация');
      if (!isTaskTable) continue;

      const startRow = r + 2;
      const col = c + 1;
      const len = tdmFindTableLength_(values, r + 1, c);
      if (len <= 0) continue;

      const range = sheet.getRange(startRow, col, len, 1);
      const cleaned = range.getValues().map(row => {
        const status = String(row[0] || '').trim();
        return [allowedStatuses.includes(status) ? status : 'Новая'];
      });

      // Сначала чистим значения, потом ставим выпадающий список.
      // Проверка мягкая, чтобы старые значения не ломали запуск скрипта.
      range.setValues(cleaned);
      range.setDataValidation(statusRule);
    }
  }
}

function tdmFindTableLength_(values, startRowIndex, statusColIndex) {
  let len = 0;

  for (let r = startRowIndex; r < values.length; r++) {
    const row = values[r];
    const firstCell = String(row[0] || '').trim();

    if (!firstCell && !String(row[statusColIndex] || '').trim()) break;
    if (firstCell && row.slice(1).every(v => !String(v || '').trim()) && len > 0) break;

    len++;
  }

  return len;
}

/**
 * Бюджетный процент из текста.
 */
function tdmFindBudgetPctInValues_(values) {
  const maxRows = Math.min(values.length, 80);
  const maxCols = values[0] ? Math.min(values[0].length, 20) : 0;

  for (let r = 0; r < maxRows; r++) {
    for (let c = 0; c < maxCols; c++) {
      const text = String(values[r][c] || '').toLowerCase();
      if (!text) continue;

      if (text.includes('бюджет') && text.includes('%')) {
        const match = text.replace(',', '.').match(/(\d+(\.\d+)?)\s*%/);
        if (match) return Number(match[1]);
      }

      if (text.includes('от плана')) {
        const match = text.replace(',', '.').match(/(\d+(\.\d+)?)/);
        if (match) return Number(match[1]);
      }
    }
  }

  return null;
}

/**
 * Период прошлой полной недели.
 */
function tdmGetPreviousWeekPeriodObj_() {
  const today = new Date();
  const day = today.getDay(); // 0 воскресенье, 1 понедельник
  const diffToMonday = day === 0 ? -6 : 1 - day;

  const currentMonday = new Date(today);
  currentMonday.setHours(0, 0, 0, 0);
  currentMonday.setDate(today.getDate() + diffToMonday);

  const previousMonday = new Date(currentMonday);
  previousMonday.setDate(currentMonday.getDate() - 7);

  const previousSunday = new Date(previousMonday);
  previousSunday.setDate(previousMonday.getDate() + 6);
  previousSunday.setHours(23, 59, 59, 999);

  const startText = Utilities.formatDate(previousMonday, TDM_CONFIG.TIMEZONE, 'dd.MM.yyyy');
  const endText = Utilities.formatDate(previousSunday, TDM_CONFIG.TIMEZONE, 'dd.MM.yyyy');

  return {
    start: previousMonday,
    end: previousSunday,
    text: `${startText}–${endText}`
  };
}

/**
 * Служебные функции.
 */
function tdmGetSpreadsheet_() {
  const active = SpreadsheetApp.getActiveSpreadsheet();
  if (active) {
    return active;
  }

  if (!TDM_CONFIG.SPREADSHEET_ID) {
    throw new Error('Не указан ID таблицы TDM_CONFIG.SPREADSHEET_ID.');
  }

  return SpreadsheetApp.openById(TDM_CONFIG.SPREADSHEET_ID);
}

function tdmShowMessage_(message) {
  try {
    SpreadsheetApp.getUi().alert(message);
  } catch (e) {
    Logger.log(message);
  }
}

function tdmGetOrCreatePlanSheet_(ss) {
  let sheet = ss.getSheetByName(TDM_CONFIG.PLAN_SHEET_NAME);

  if (!sheet) {
    sheet = ss.insertSheet(TDM_CONFIG.PLAN_SHEET_NAME);
  }

  // Переносим вкладку в конец.
  ss.setActiveSheet(sheet);
  ss.moveActiveSheet(ss.getSheets().length);

  return sheet;
}

function tdmHeaderMatch_(header, aliases) {
  return aliases.some(alias => {
    const a = tdmNorm_(alias);
    return header === a || header.includes(a);
  });
}

function tdmMatchesGoal_(header, goals) {
  return goals.some(goal => header.includes(tdmNorm_(goal)));
}

function tdmNorm_(value) {
  return String(value || '')
    .toLowerCase()
    .replace(/\u00a0/g, ' ')
    .replace(/[ё]/g, 'е')
    .replace(/[.,;:()"'«»]/g, ' ')
    .replace(/\s+/g, ' ')
    .trim();
}

function tdmNum_(value) {
  if (typeof value === 'number') return value;
  if (value === null || value === undefined || value === '') return 0;

  let text = String(value)
    .replace(/\u00a0/g, '')
    .replace(/\s/g, '')
    .replace('%', '')
    .replace('руб.', '')
    .replace('руб', '')
    .replace(',', '.')
    .replace(/[^\d.-]/g, '');

  const num = Number(text);
  return isNaN(num) ? 0 : num;
}

function tdmParseDate_(value) {
  if (Object.prototype.toString.call(value) === '[object Date]' && !isNaN(value)) {
    return value;
  }

  const text = String(value || '').trim();
  if (!text) return null;

  const match = text.match(/(\d{1,2})[.\-/](\d{1,2})[.\-/](\d{2,4})/);
  if (!match) return null;

  const day = Number(match[1]);
  const month = Number(match[2]) - 1;
  let year = Number(match[3]);
  if (year < 100) year += 2000;

  const date = new Date(year, month, day);
  return isNaN(date) ? null : date;
}

function tdmIsEmptyRow_(row) {
  return row.every(cell => String(cell || '').trim() === '');
}

function tdmMoney_(value) {
  const num = Number(value || 0);
  return Utilities.formatString('%.2f р.', num).replace('.', ',');
}

function tdmPct_(value) {
  const num = Number(value || 0);
  return Utilities.formatString('%.2f%%', num).replace('.', ',');
}

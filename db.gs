// VAG-House DB monthly block from May template v04.1
// Файл для Apps Script. Вставляй в отдельный файл, например db_month_v04.gs.
// Код создаёт/пересобирает месячный блок ДБ на основе форматирования майского блока и заполняет:
// 1) показы/клики/расход из Директа;
// 2) цели из Метрики по логике ЕНО, чтобы совпадало с Метрикой;
// 3) Факт сумма конверсий = Звонки_Уник - ЮИС + Заявки Юис.
//
// Что запускать:
// createDbJune2026FromMayTemplate_v4() — создать/пересобрать июнь 2026 с форматированием как в мае, с верхним блоком формул и чёрными границами.
// createDbCurrentMonthFromMayTemplate_v4() — создать/пересобрать текущий месяц с форматированием как в мае.
// checkDbMetrikaAccess_v4() — проверить доступ к Метрике.

const DB_VAG_V04_CONFIG = {
  SHEET_NAME: 'ДБ',
  CLIENT_LOGIN: 'vag-house98',
  TIMEZONE: 'Europe/Moscow',
  MONTH_BUDGET: 180000,
  METRIKA_COUNTER_ID: '99052519',
  METRIKA_ATTRIBUTION: 'automatic',
  DIRECT_ATTRIBUTION_MODEL: 'AUTO',
  UIS_TOKEN_PROPERTY: 'UIS_TOKEN',
  UIS_API_URL: 'https://dataapi.uiscom.ru/v2.0',
  UIS_QUALITY_VALUE: 'Качественные',
  PROTECTED_ROWS: {
    SKIP_ROW: 234,
    SKIP_FROM_ROW: 242
  },
  START_COLUMN: 2, // B
  TABLE_WIDTH: 20, // B:U
  CLEAR_ROWS: 80,
  GOALS: {
    bb60sec2pages: '555861179',    // I — B-B_60сек+2стр
    phone8835: '511787501',        // J — Клик на 79110008835
    phone8865: '511788019',        // K — Клик на 79110008865
    max: '511788150',              // L — Клик на Макс
    telegramSlider: '511791937',   // M — Успешно пройден слайдер для перехода в TG
    route: '547214611'             // N — Клик по кнопке "Маршрут"
  }
};

// =====================
// ЗАПУСК
// =====================

function createDbJune2026FromMayTemplate_v4() {
  DB_VAG_V04_createMonthBlockFromMayTemplate_(2026, 6);
}

function createDbCurrentMonthFromMayTemplate_v4() {
  const now = new Date();
  const year = Number(Utilities.formatDate(now, DB_VAG_V04_CONFIG.TIMEZONE, 'yyyy'));
  const month = Number(Utilities.formatDate(now, DB_VAG_V04_CONFIG.TIMEZONE, 'M'));
  DB_VAG_V04_createMonthBlockFromMayTemplate_(year, month);
}

function checkDbMetrikaAccess_v4() {
  const period = DB_VAG_V04_getMonthPeriodForReport_(2026, 6);
  const data = DB_VAG_V04_getMetrikaGoalsByDate_(period.dateFrom, period.dateTo);
  Logger.log('Доступ к Метрике есть. Дней с данными по целям найдено: ' + Object.keys(data).length);
}

function testUisApiConnection() {
  const token = DB_VAG_V04_getUisToken_();
  const yesterday = new Date();
  yesterday.setDate(yesterday.getDate() - 1);
  const date = Utilities.formatDate(yesterday, DB_VAG_V04_CONFIG.TIMEZONE, 'yyyy-MM-dd');
  const rows = DB_VAG_V04_getUisReportRows_('get.calls_report', date, date, DB_VAG_V04_getUisCallFields_(), token, 1);
  const firstRow = rows[0] || {};

  Logger.log('UIS API status: OK');
  Logger.log('UIS rows found in test response: ' + rows.length);
  Logger.log('UIS first row structure without personal data: ' + JSON.stringify(DB_VAG_V04_maskUisRowForLog_(firstRow)));
}

function previewUisDailyDataForDb() {
  const now = new Date();
  const year = Number(Utilities.formatDate(now, DB_VAG_V04_CONFIG.TIMEZONE, 'yyyy'));
  const month = Number(Utilities.formatDate(now, DB_VAG_V04_CONFIG.TIMEZONE, 'M'));
  const period = DB_VAG_V04_getMonthPeriodForReport_(year, month);
  const uisByDate = DB_VAG_V04_getUisDailyDataForDb_(period.dateFrom, period.dateTo);

  Object.keys(uisByDate).sort().forEach(date => {
    const row = uisByDate[date];
    Logger.log(date + ' -> O=' + row.callsTotalFromUis + ', P=' + row.qualityCallsFromUis + ', Q=' + row.leadsFromUis);
  });
}

function diagnoseUisDailyDataForDb() {
  const now = new Date();
  const year = Number(Utilities.formatDate(now, DB_VAG_V04_CONFIG.TIMEZONE, 'yyyy'));
  const month = Number(Utilities.formatDate(now, DB_VAG_V04_CONFIG.TIMEZONE, 'M'));
  const period = DB_VAG_V04_getMonthPeriodForReport_(year, month);
  const token = DB_VAG_V04_getUisToken_();

  const calls = DB_VAG_V04_getUisReportRows_('get.calls_report', period.dateFrom, period.dateTo, DB_VAG_V04_getUisCallFields_(), token);
  const leads = DB_VAG_V04_getUisReportRows_('get.offline_messages_report', period.dateFrom, period.dateTo, DB_VAG_V04_getUisLeadFields_(), token);
  const directCalls = calls.filter(row => DB_VAG_V04_isYandexDirectUisRow_(row));
  const directLeads = leads.filter(row => DB_VAG_V04_isYandexDirectUisRow_(row));

  Logger.log('UIS period: ' + period.dateFrom + ' — ' + period.dateTo);
  Logger.log('UIS calls rows total before source filter: ' + calls.length);
  Logger.log('UIS calls rows matched as Yandex Direct: ' + directCalls.length);
  Logger.log('UIS leads rows total before source filter: ' + leads.length);
  Logger.log('UIS leads rows matched as Yandex Direct: ' + directLeads.length);
  Logger.log('UIS calls source-like values: ' + JSON.stringify(DB_VAG_V04_collectUisSourceDebugValues_(calls)));
  Logger.log('UIS leads source-like values: ' + JSON.stringify(DB_VAG_V04_collectUisSourceDebugValues_(leads)));
  Logger.log('UIS first call row keys: ' + Object.keys(calls[0] || {}).join(', '));
  Logger.log('UIS first call row without personal data: ' + JSON.stringify(DB_VAG_V04_maskUisRowForLog_(calls[0] || {})));
  Logger.log('UIS first lead row keys: ' + Object.keys(leads[0] || {}).join(', '));
  Logger.log('UIS first lead row without personal data: ' + JSON.stringify(DB_VAG_V04_maskUisRowForLog_(leads[0] || {})));
}

// =====================
// ОСНОВНАЯ ЛОГИКА
// =====================

function DB_VAG_V04_createMonthBlockFromMayTemplate_(year, month) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(DB_VAG_V04_CONFIG.SHEET_NAME);

  if (!sheet) {
    throw new Error('Не найден лист ' + DB_VAG_V04_CONFIG.SHEET_NAME);
  }

  const period = DB_VAG_V04_getMonthPeriodForReport_(year, month);
  const monthLabel = DB_VAG_V04_getMonthLabel_(year, month);
  const daysInMonth = DB_VAG_V04_getDaysInMonth_(year, month);
  const template = DB_VAG_V04_getMayTemplateInfo_(sheet);
  const topRow = DB_VAG_V04_findOrCreateMonthTopRow_(sheet, monthLabel, template);

  const headerRow = topRow + template.headerOffset;
  const dataStartRow = headerRow + 1;
  const totalRow = dataStartRow + daysInMonth;

  DB_VAG_V04_assertMonthBlockSafeToRebuild_(topRow, totalRow + template.afterTotalRows);

  DB_VAG_V04_prepareBlockAreaFromTemplate_(sheet, topRow, template, daysInMonth);

  const directByDate = DB_VAG_V04_getDirectDailyStats_(period.dateFrom, period.dateTo);
  const goalsByDate = DB_VAG_V04_getMetrikaGoalsByDate_(period.dateFrom, period.dateTo);
  const uisByDate = DB_VAG_V04_getUisDailyDataForDb_(period.dateFrom, period.dateTo);

  DB_VAG_V04_writeMonthHeader_(sheet, topRow, monthLabel, period, template, totalRow);
  DB_VAG_V04_writeTableHeader_(sheet, headerRow);
  DB_VAG_V04_writeDailyRows_(sheet, dataStartRow, year, month, daysInMonth, directByDate, goalsByDate, uisByDate, period.dateTo);
  DB_VAG_V04_writeTotalsAndPlan_(sheet, totalRow, dataStartRow, dataStartRow + daysInMonth - 1, period);
  DB_VAG_V04_applyFormatting_(sheet, topRow, headerRow, dataStartRow, totalRow, template, year, month);

  SpreadsheetApp.flush();
  Logger.log('Готово: блок ДБ "' + monthLabel + '" создан/пересобран с форматированием как в мае за период ' + period.dateFrom + ' — ' + period.dateTo);
}

function DB_VAG_V04_getMonthPeriodForReport_(year, month) {
  const todayStr = Utilities.formatDate(new Date(), DB_VAG_V04_CONFIG.TIMEZONE, 'yyyy-MM-dd');
  const today = DB_VAG_V04_parseApiDate_(todayStr);
  const monthStart = new Date(year, month - 1, 1);
  const monthEnd = new Date(year, month, 0);
  const yesterday = new Date(today);
  yesterday.setDate(today.getDate() - 1);

  let reportEnd = monthEnd;
  if (year === today.getFullYear() && month === today.getMonth() + 1) {
    reportEnd = yesterday < monthStart ? monthStart : yesterday;
  }

  return {
    dateFrom: Utilities.formatDate(monthStart, DB_VAG_V04_CONFIG.TIMEZONE, 'yyyy-MM-dd'),
    dateTo: Utilities.formatDate(reportEnd, DB_VAG_V04_CONFIG.TIMEZONE, 'yyyy-MM-dd'),
    monthStart: Utilities.formatDate(monthStart, DB_VAG_V04_CONFIG.TIMEZONE, 'yyyy-MM-dd'),
    monthEnd: Utilities.formatDate(monthEnd, DB_VAG_V04_CONFIG.TIMEZONE, 'yyyy-MM-dd')
  };
}

// =====================
// СОЗДАНИЕ БЛОКА В ТАБЛИЦЕ
// =====================

function DB_VAG_V04_findOrCreateMonthTopRow_(sheet, monthLabel, template) {
  const lastRow = Math.max(sheet.getLastRow(), 1);
  const values = sheet.getRange(1, 2, lastRow, 1).getDisplayValues(); // B

  for (let i = 0; i < values.length; i++) {
    if (String(values[i][0]).trim() === monthLabel) {
      const candidateRow = i + 1;
      // Старый ошибочный июнь мог оказаться выше майского шаблона.
      // Такой блок игнорируем, чтобы не стереть майский формат при пересборке.
      if (candidateRow > template.topRow) {
        return candidateRow;
      }
    }
  }

  // Новый блок ставим после последней заполненной строки в B.
  const newTopRow = DB_VAG_V04_findLastFilledRowInColumnB_(sheet) + 4;
  const rowsNeeded = template.headerOffset + 1 + 31 + template.afterTotalRows + 8;
  if (sheet.getMaxRows() < newTopRow + rowsNeeded) {
    sheet.insertRowsAfter(sheet.getMaxRows(), newTopRow + rowsNeeded - sheet.getMaxRows());
  }
  return newTopRow;
}

function DB_VAG_V04_findLastFilledRowInColumnB_(sheet) {
  const lastRow = Math.max(sheet.getLastRow(), 1);
  const values = sheet.getRange(1, 2, lastRow, 1).getDisplayValues();

  for (let i = values.length - 1; i >= 0; i--) {
    if (String(values[i][0]).trim() !== '') return i + 1;
  }

  return 1;
}

function DB_VAG_V04_getMayTemplateInfo_(sheet) {
  const topRow = DB_VAG_V04_findMonthTopRowByLabel_(sheet, 'Май 2026');
  if (!topRow) {
    throw new Error('Не нашла шаблонный блок "Май 2026" на листе ДБ. Нужен майский блок для копирования форматирования.');
  }

  const headerRow = DB_VAG_V04_findHeaderRowInBlock_(sheet, topRow);
  if (!headerRow) {
    throw new Error('Не нашла шапку с "Дата" в майском блоке.');
  }

  const totalRow = DB_VAG_V04_findTotalRowAfter_(sheet, headerRow);
  if (!totalRow) {
    throw new Error('Не нашла строку ИТОГО в майском блоке.');
  }

  const headerOffset = headerRow - topRow;
  const dataStartRow = headerRow + 1;
  const templateDays = Math.max(1, totalRow - dataStartRow);
  const afterTotalRows = 8;

  return {
    topRow: topRow,
    headerRow: headerRow,
    headerOffset: headerOffset,
    dataStartRow: dataStartRow,
    totalRow: totalRow,
    templateDays: templateDays,
    afterTotalRows: afterTotalRows
  };
}

function DB_VAG_V04_findMonthTopRowByLabel_(sheet, label) {
  const lastRow = Math.max(sheet.getLastRow(), 1);
  const values = sheet.getRange(1, 2, lastRow, 1).getDisplayValues();

  for (let i = 0; i < values.length; i++) {
    if (String(values[i][0]).trim() === label) return i + 1;
  }

  return 0;
}

function DB_VAG_V04_findHeaderRowInBlock_(sheet, topRow) {
  const maxRow = Math.min(sheet.getLastRow(), topRow + 20);
  const width = DB_VAG_V04_CONFIG.TABLE_WIDTH;

  for (let row = topRow; row <= maxRow; row++) {
    const values = sheet.getRange(row, DB_VAG_V04_CONFIG.START_COLUMN, 1, width).getDisplayValues()[0];
    const normalized = values.map(value => DB_VAG_V04_normalize_(value));
    if (normalized.indexOf('дата') !== -1 && normalized.some(value => value.indexOf('показы') !== -1)) {
      return row;
    }
  }

  return 0;
}

function DB_VAG_V04_findTotalRowAfter_(sheet, headerRow) {
  const maxRow = Math.min(sheet.getLastRow(), headerRow + 90);

  for (let row = headerRow + 1; row <= maxRow; row++) {
    const value = DB_VAG_V04_normalize_(sheet.getRange(row, DB_VAG_V04_CONFIG.START_COLUMN).getDisplayValue());
    if (value === 'итого') return row;
  }

  return 0;
}

function DB_VAG_V04_prepareBlockAreaFromTemplate_(sheet, topRow, template, daysInMonth) {
  const width = DB_VAG_V04_CONFIG.TABLE_WIDTH;
  const startCol = DB_VAG_V04_CONFIG.START_COLUMN;
  const rowsNeeded = template.headerOffset + 1 + daysInMonth + 1 + template.afterTotalRows;

  if (sheet.getMaxRows() < topRow + rowsNeeded + 5) {
    sheet.insertRowsAfter(sheet.getMaxRows(), topRow + rowsNeeded + 5 - sheet.getMaxRows());
  }

  // Очищаем старый/кривой блок в зоне B:U.
  sheet.getRange(topRow, startCol, Math.max(DB_VAG_V04_CONFIG.CLEAR_ROWS, rowsNeeded + 5), width)
    .clearContent()
    .breakApart()
    .setBackground('#ffffff')
    .setFontColor('#000000')
    .setFontWeight('normal')
    .setFontSize(10)
    .setWrap(false)
    .setBorder(false, false, false, false, false, false);

  // Копируем формат верхней части майского блока до шапки включительно.
  const topRows = template.headerOffset + 1;
  sheet.getRange(template.topRow, startCol, topRows, width)
    .copyTo(sheet.getRange(topRow, startCol, topRows, width), SpreadsheetApp.CopyPasteType.PASTE_FORMAT, false);

  for (let i = 0; i < topRows; i++) {
    sheet.setRowHeight(topRow + i, sheet.getRowHeight(template.topRow + i));
  }

  const targetHeaderRow = topRow + template.headerOffset;
  const targetDataStartRow = targetHeaderRow + 1;
  const targetTotalRow = targetDataStartRow + daysInMonth;

  // Копируем формат строк дней из мая. Если дней меньше/больше, берём ближайшую строку шаблона.
  for (let i = 0; i < daysInMonth; i++) {
    const sourceRow = template.dataStartRow + Math.min(i, template.templateDays - 1);
    const targetRow = targetDataStartRow + i;
    sheet.getRange(sourceRow, startCol, 1, width)
      .copyTo(sheet.getRange(targetRow, startCol, 1, width), SpreadsheetApp.CopyPasteType.PASTE_FORMAT, false);
    sheet.setRowHeight(targetRow, sheet.getRowHeight(sourceRow));
  }

  // ИТОГО и служебные строки ниже — из майского блока.
  sheet.getRange(template.totalRow, startCol, 1, width)
    .copyTo(sheet.getRange(targetTotalRow, startCol, 1, width), SpreadsheetApp.CopyPasteType.PASTE_FORMAT, false);
  sheet.setRowHeight(targetTotalRow, sheet.getRowHeight(template.totalRow));

  if (template.afterTotalRows > 0) {
    sheet.getRange(template.totalRow + 1, startCol, template.afterTotalRows, width)
      .copyTo(sheet.getRange(targetTotalRow + 1, startCol, template.afterTotalRows, width), SpreadsheetApp.CopyPasteType.PASTE_FORMAT, false);

    for (let i = 0; i < template.afterTotalRows; i++) {
      sheet.setRowHeight(targetTotalRow + 1 + i, sheet.getRowHeight(template.totalRow + 1 + i));
    }
  }
}

function DB_VAG_V04_writeMonthHeader_(sheet, topRow, monthLabel, period, template, totalRow) {
  const startCol = DB_VAG_V04_CONFIG.START_COLUMN;
  const width = DB_VAG_V04_CONFIG.TABLE_WIDTH;
  const periodRow = topRow + 1;
  const daysTotalRow = topRow + 2;
  const daysPassedRow = topRow + 3;
  const daysLeftRow = topRow + 4;
  const reportDateRow = topRow + 5;
  const budgetRow = topRow + template.headerOffset - 2;

  // Верхний блок B:S как в майском блоке: месяц + сводные формулы по ИТОГО.
  sheet.getRange(topRow, startCol, 1, width).clearContent();
  sheet.getRange(topRow, 2).setValue(monthLabel);

  sheet.getRange(topRow, 7).setValue('Расход');
  sheet.getRange(topRow, 8).setFormula('=F' + totalRow);
  sheet.getRange(topRow, 9).setValue('Клики');
  sheet.getRange(topRow, 10).setFormula('=D' + totalRow);
  sheet.getRange(topRow, 11).setValue('cpc');
  sheet.getRange(topRow, 12).setFormula('=H' + totalRow);
  sheet.getRange(topRow, 14).setValue('Кол-во Лид');
  sheet.getRange(topRow, 15).setFormula('=R' + totalRow);
  sheet.getRange(topRow, 16).setValue('CR');
  sheet.getRange(topRow, 17).setFormula('=S' + totalRow);
  sheet.getRange(topRow, 18).setValue('Факт CPA');
  sheet.getRange(topRow, 19).setFormula('=T' + totalRow);

  // Блок периода РК / дней / даты отчёта как в строках 128–133 майского блока.
  sheet.getRange(periodRow, 2).setValue('Период РК');
  sheet.getRange(periodRow, 3).setValue(DB_VAG_V04_parseApiDate_(period.monthStart));
  sheet.getRange(periodRow, 4).setValue(DB_VAG_V04_parseApiDate_(period.monthEnd));

  sheet.getRange(daysTotalRow, 2).setValue('Дней всего');
  sheet.getRange(daysTotalRow, 3).setFormula('=D' + periodRow + '-C' + periodRow + '+1');

  sheet.getRange(daysPassedRow, 2).setValue('Дней прошло');
  sheet.getRange(daysPassedRow, 3).setFormula('=C' + daysTotalRow + '-C' + daysLeftRow);
  sheet.getRange(daysPassedRow, 4).setFormula('=IFERROR(C' + daysPassedRow + '/C' + daysTotalRow + ';0)');

  sheet.getRange(daysLeftRow, 2).setValue('Осталось');
  sheet.getRange(daysLeftRow, 3).setFormula('=MAX(0;D' + periodRow + '-TODAY())');
  sheet.getRange(daysLeftRow, 4).setFormula('=IFERROR(C' + daysLeftRow + '/C' + daysTotalRow + ';0)');

  sheet.getRange(reportDateRow, 2).setValue('Дата отчета');
  sheet.getRange(reportDateRow, 3).setFormula('=TODAY()');

  // Нижний заголовок месяца перед таблицей: месяц / бюджет / резерв.
  if (template.headerOffset > 1) {
    sheet.getRange(topRow + 6, startCol, Math.max(0, template.headerOffset - 7), width).clearContent();
  }
  sheet.getRange(budgetRow, startCol, 1, width).clearContent();
  sheet.getRange(budgetRow, 2).setValue(monthLabel);
  sheet.getRange(budgetRow, 5).setValue('Бюджет с НДС');
  sheet.getRange(budgetRow, 6).setValue(DB_VAG_V04_CONFIG.MONTH_BUDGET);
  sheet.getRange(budgetRow, 7).setValue('10к резерв');

  // Форматы верхнего блока.
  sheet.getRange(periodRow, 3, 1, 2).setNumberFormat('dd.mm.yyyy');
  sheet.getRange(reportDateRow, 3).setNumberFormat('dd.mm.yyyy');
  sheet.getRange(daysPassedRow, 4, 2, 1).setNumberFormat('0.00%');
  sheet.getRange(topRow, 8).setNumberFormat('[$р.-419]#,##0.00');
  sheet.getRange(topRow, 10).setNumberFormat('#,##0');
  sheet.getRange(topRow, 12).setNumberFormat('[$р.-419]#,##0.00');
  sheet.getRange(topRow, 15).setNumberFormat('#,##0');
  sheet.getRange(topRow, 17).setNumberFormat('0.00%');
  sheet.getRange(topRow, 19).setNumberFormat('[$р.-419]#,##0.00');
  sheet.getRange(budgetRow, 6).setNumberFormat('[$р.-419]#,##0.00');
}

function DB_VAG_V04_writeTableHeader_(sheet, headerRow) {
  const headers = [
    'Дата',
    'Показы',
    'Клики',
    'Расход План c НДС',
    'Расход ФАКТ с НДС (руб.)',
    'CTR',
    'CPC',
    'B-B_60сек+2стр (555861179)',
    'Клик на 79110008835 (511787501)',
    'Клик на 79110008865 (511788019)',
    'Клик на Макс (511788150)',
    'Успешно пройден слайдер для перехода в TG (511791937)',
    'Клик по кнопке "Маршрут" (547214611)',
    'Звонки, количество из ЮИС',
    'Качественные звонки из ЮИС',
    'Заявки, количество из ЮИС',
    'Факт сумма конверсий',
    'CR',
    'Стоимость факт CPA',
    'Сумма по неделям'
  ];

  sheet.getRange(headerRow, 2, 1, headers.length).setValues([headers]);
}

function DB_VAG_V04_writeDailyRows_(sheet, dataStartRow, year, month, daysInMonth, directByDate, goalsByDate, uisByDate, reportDateTo) {
  const rows = [];
  const reportEndDate = DB_VAG_V04_parseApiDate_(reportDateTo);

  for (let day = 1; day <= daysInMonth; day++) {
    const date = new Date(year, month - 1, day);
    const apiDate = Utilities.formatDate(date, DB_VAG_V04_CONFIG.TIMEZONE, 'yyyy-MM-dd');
    const isFuture = date > reportEndDate;
    const direct = directByDate[apiDate] || { impressions: '', clicks: '', cost: '' };
    const goals = goalsByDate[apiDate] || DB_VAG_V04_emptyGoals_();
    const uis = uisByDate[apiDate] || DB_VAG_V04_emptyUisDailyRow_();

    rows.push([
      DB_VAG_V04_formatDateRuFull_(apiDate),
      isFuture ? '' : direct.impressions,
      isFuture ? '' : direct.clicks,
      DB_VAG_V04_CONFIG.MONTH_BUDGET / daysInMonth,
      isFuture ? '' : direct.cost,
      '',
      '',
      isFuture ? '' : DB_VAG_V04_blankZero_(goals.bb60sec2pages),
      isFuture ? '' : DB_VAG_V04_blankZero_(goals.phone8835),
      isFuture ? '' : DB_VAG_V04_blankZero_(goals.phone8865),
      isFuture ? '' : DB_VAG_V04_blankZero_(goals.max),
      isFuture ? '' : DB_VAG_V04_blankZero_(goals.telegramSlider),
      isFuture ? '' : DB_VAG_V04_blankZero_(goals.route),
      isFuture ? '' : DB_VAG_V04_blankZero_(uis.callsTotalFromUis),
      isFuture ? '' : DB_VAG_V04_blankZero_(uis.qualityCallsFromUis),
      isFuture ? '' : DB_VAG_V04_blankZero_(uis.leadsFromUis),
      '',
      '',
      '',
      ''
    ]);
  }

  sheet.getRange(dataStartRow, 2, rows.length, DB_VAG_V04_CONFIG.TABLE_WIDTH).setValues(rows);

  for (let i = 0; i < daysInMonth; i++) {
    const row = dataStartRow + i;
    sheet.getRange(row, 7).setFormula('=IFERROR(D' + row + '/C' + row + ';0)');   // G CTR
    sheet.getRange(row, 8).setFormula('=IFERROR(F' + row + '/D' + row + ';0)');   // H CPC
    sheet.getRange(row, 18).setFormula('=SUM(P' + row + ':Q' + row + ')');         // R Fact conversions = P + Q
    sheet.getRange(row, 19).setFormula('=IFERROR(R' + row + '/D' + row + ';0)');   // S CR
    sheet.getRange(row, 20).setFormula('=IFERROR(F' + row + '/R' + row + ';0)');   // T CPA

    const apiDate = Utilities.formatDate(new Date(year, month - 1, i + 1), DB_VAG_V04_CONFIG.TIMEZONE, 'yyyy-MM-dd');
    const date = DB_VAG_V04_parseApiDate_(apiDate);
    const isSunday = date.getDay() === 0;
    const isLastDay = i + 1 === daysInMonth;
    if (isSunday || isLastDay) {
      const weekStartRow = Math.max(dataStartRow, row - (date.getDay() === 0 ? 6 : date.getDay() - 1));
      sheet.getRange(row, 21).setFormula('=SUM(R' + weekStartRow + ':R' + row + ')');
    }
  }
}

function DB_VAG_V04_writeTotalsAndPlan_(sheet, totalRow, dataStartRow, dataEndRow, period) {
  sheet.getRange(totalRow, 2).setValue('ИТОГО');

  const sumColumns = [3, 4, 5, 6, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18];
  sumColumns.forEach(col => {
    const letter = DB_VAG_V04_columnToLetter_(col);
    sheet.getRange(totalRow, col).setFormula('=SUBTOTAL(9;' + letter + dataStartRow + ':' + letter + dataEndRow + ')');
  });

  sheet.getRange(totalRow, 7).setFormula('=IFERROR(D' + totalRow + '/C' + totalRow + ';0)');
  sheet.getRange(totalRow, 8).setFormula('=IFERROR(F' + totalRow + '/D' + totalRow + ';0)');
  sheet.getRange(totalRow, 19).setFormula('=IFERROR(R' + totalRow + '/D' + totalRow + ';0)');
  sheet.getRange(totalRow, 20).setFormula('=IFERROR(F' + totalRow + '/R' + totalRow + ';0)');

  const planStart = totalRow + 1;
  const planRows = [
    ['ПЛАН по МП', '', '', DB_VAG_V04_CONFIG.MONTH_BUDGET, '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', ''],
    ['Остаток до конца месяца', '', '', '', '=E' + planStart + '-F' + totalRow, '', '', '', '', '', '', '', '', '', '', '', '', '', '', ''],
    ['Разница в бюджете', '', '', '', '=E' + planStart + '-F' + totalRow, '', '', '', '', '', '', '', '', '', '', '', '', '', '', ''],
    ['План и факт расход на вчера', '', '', '=E' + totalRow, '=F' + totalRow, '', '', '', '', '', '', '', '', '', '', '', '', '', '', ''],
    ['План и факт на дату %', '', '', '=IFERROR(E' + totalRow + '/E' + planStart + ';0)', '=IFERROR(F' + totalRow + '/E' + planStart + ';0)', '=IFERROR(F' + totalRow + '/E' + planStart + ';0)', '', '', '', '', '', '', '', '', '', '', '', '', '', '']
  ];

  sheet.getRange(planStart, 2, planRows.length, DB_VAG_V04_CONFIG.TABLE_WIDTH).setValues(planRows);
}

// =====================
// ДИРЕКТ: ДНЕВНЫЕ МЕТРИКИ
// =====================

function DB_VAG_V04_getDirectDailyStats_(dateFrom, dateTo) {
  const reportBody = {
    params: {
      SelectionCriteria: {
        DateFrom: dateFrom,
        DateTo: dateTo
      },
      FieldNames: [
        'Date',
        'CampaignName',
        'Impressions',
        'Clicks',
        'Cost'
      ],
      OrderBy: [
        { Field: 'Date', SortOrder: 'ASCENDING' }
      ],
      ReportName: 'db_daily_v4_' + DB_VAG_V04_CONFIG.CLIENT_LOGIN + '_' + dateFrom + '_' + dateTo + '_' + Utilities.getUuid(),
      ReportType: 'CAMPAIGN_PERFORMANCE_REPORT',
      DateRangeType: 'CUSTOM_DATE',
      Format: 'TSV',
      IncludeVAT: 'YES',
      IncludeDiscount: 'NO'
    }
  };

  const text = DB_VAG_V04_requestDirectReport_(reportBody);
  return DB_VAG_V04_parseDirectDailyTsv_(text);
}

function DB_VAG_V04_requestDirectReport_(reportBody) {
  const token = PropertiesService.getScriptProperties().getProperty('YANDEX_TOKEN');
  if (!token) throw new Error('Не найден YANDEX_TOKEN в свойствах скрипта');

  const url = 'https://api.direct.yandex.com/json/v5/reports';
  const options = {
    method: 'post',
    contentType: 'application/json',
    headers: {
      Authorization: 'Bearer ' + token,
      'Client-Login': DB_VAG_V04_CONFIG.CLIENT_LOGIN,
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

  throw new Error('Отчёт Директа долго формируется. Запусти позже.');
}

function DB_VAG_V04_parseDirectDailyTsv_(tsvText) {
  const lines = String(tsvText || '').trim().split(/\r?\n/);
  if (lines.length < 2) return {};

  const headers = lines[0].split('\t').map(h => String(h).replace(/^\uFEFF/, '').trim());
  const idx = {};
  headers.forEach((h, i) => idx[h] = i);

  const result = {};

  lines.slice(1).forEach(line => {
    if (!String(line).trim()) return;
    const row = line.split('\t');
    const date = String(row[idx['Date']] || '').trim();
    if (!date) return;

    if (!result[date]) {
      result[date] = { impressions: 0, clicks: 0, cost: 0 };
    }

    result[date].impressions += DB_VAG_V04_toNumber_(row[idx['Impressions']]);
    result[date].clicks += DB_VAG_V04_toNumber_(row[idx['Clicks']]);
    result[date].cost += DB_VAG_V04_toNumber_(row[idx['Cost']]);
  });

  return result;
}

// =====================
// МЕТРИКА: ЦЕЛИ ПО ДНЯМ И КАМПАНИЯМ ДИРЕКТА
// =====================

function DB_VAG_V04_getMetrikaGoalsByDate_(dateFrom, dateTo) {
  const token = DB_VAG_V04_getMetrikaToken_();
  if (!token) throw new Error('Не найден METRIKA_TOKEN или YANDEX_TOKEN в свойствах скрипта');

  const goalEntries = DB_VAG_V04_getGoalEntries_();
  const metrics = goalEntries.map(entry => 'ym:s:goal' + entry.id + 'reaches').join(',');

  const params = {
    ids: DB_VAG_V04_CONFIG.METRIKA_COUNTER_ID,
    date1: dateFrom,
    date2: dateTo,
    metrics: metrics,
    dimensions: 'ym:s:date,ym:s:' + DB_VAG_V04_CONFIG.METRIKA_ATTRIBUTION + 'DirectClickOrderName',
    accuracy: 'full',
    limit: '100000',
    lang: 'ru'
  };

  const url = 'https://api-metrika.yandex.net/stat/v1/data?' + DB_VAG_V04_toQueryString_(params);
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
    throw new Error('Ошибка Метрики: ' + code + '\n' + text);
  }

  const json = JSON.parse(text);
  const result = {};

  (json.data || []).forEach(item => {
    const dimensions = item.dimensions || [];
    const date = dimensions[0] && dimensions[0].name ? String(dimensions[0].name).trim() : '';
    const campaignName = dimensions[1] && dimensions[1].name ? String(dimensions[1].name).trim() : '';
    const cleanCampaign = DB_VAG_V04_cleanCampaignName_(campaignName);

    if (!date || !cleanCampaign || cleanCampaign === 'none' || cleanCampaign === 'not set') return;

    if (!result[date]) result[date] = DB_VAG_V04_emptyGoals_();

    goalEntries.forEach((entry, index) => {
      result[date][entry.key] += DB_VAG_V04_toNumber_((item.metrics || [])[index]);
    });
  });

  return result;
}

function DB_VAG_V04_getMetrikaToken_() {
  const props = PropertiesService.getScriptProperties();
  return props.getProperty('METRIKA_TOKEN') || props.getProperty('YANDEX_TOKEN');
}

function DB_VAG_V04_getGoalEntries_() {
  return [
    { key: 'bb60sec2pages', id: DB_VAG_V04_CONFIG.GOALS.bb60sec2pages },
    { key: 'phone8835', id: DB_VAG_V04_CONFIG.GOALS.phone8835 },
    { key: 'phone8865', id: DB_VAG_V04_CONFIG.GOALS.phone8865 },
    { key: 'max', id: DB_VAG_V04_CONFIG.GOALS.max },
    { key: 'telegramSlider', id: DB_VAG_V04_CONFIG.GOALS.telegramSlider },
    { key: 'route', id: DB_VAG_V04_CONFIG.GOALS.route }
  ];
}

function DB_VAG_V04_emptyGoals_() {
  return {
    bb60sec2pages: 0,
    phone8835: 0,
    phone8865: 0,
    max: 0,
    telegramSlider: 0,
    route: 0
  };
}

function DB_VAG_V04_cleanCampaignName_(name) {
  return String(name || '')
    .replace(/\s*\(N-\d+\)\s*$/i, '')
    .replace(/\s+/g, ' ')
    .trim()
    .toLowerCase();
}

// =====================
// UIS: ЗВОНКИ И ЗАЯВКИ ДЛЯ O:P:Q
// =====================

function DB_VAG_V04_getUisDailyDataForDb_(dateFrom, dateTo) {
  const token = DB_VAG_V04_getUisToken_();
  const result = DB_VAG_V04_makeEmptyUisDailyMap_(dateFrom, dateTo);

  const calls = DB_VAG_V04_getUisReportRows_('get.calls_report', dateFrom, dateTo, DB_VAG_V04_getUisCallFields_(), token);
  calls.forEach(row => {
    if (!DB_VAG_V04_isYandexDirectUisRow_(row)) return;
    const date = DB_VAG_V04_getUisRowDate_(row);
    if (!result[date]) return;

    result[date].callsTotalFromUis += 1;
    if (DB_VAG_V04_isQualityUisRow_(row)) {
      result[date].qualityCallsFromUis += 1;
    }
  });

  const leads = DB_VAG_V04_getUisReportRows_('get.offline_messages_report', dateFrom, dateTo, DB_VAG_V04_getUisLeadFields_(), token);
  leads.forEach(row => {
    if (!DB_VAG_V04_isYandexDirectUisRow_(row)) return;
    const date = DB_VAG_V04_getUisRowDate_(row);
    if (!result[date]) return;

    result[date].leadsFromUis += 1;
  });

  return result;
}

function DB_VAG_V04_getUisToken_() {
  const token = String(PropertiesService.getScriptProperties().getProperty(DB_VAG_V04_CONFIG.UIS_TOKEN_PROPERTY) || '').trim();
  if (!token) {
    throw new Error('Не найден UIS API token в Script Properties. Ожидаемое имя свойства: ' + DB_VAG_V04_CONFIG.UIS_TOKEN_PROPERTY);
  }
  return token;
}

function DB_VAG_V04_getUisReportRows_(method, dateFrom, dateTo, fields, token, customLimit) {
  let activeFields = fields.slice();

  while (true) {
    const limit = customLimit || 1000;
    let offset = 0;
    let rows = [];

    try {
      while (true) {
        const payload = DB_VAG_V04_buildUisRequest_(method, {
          date_from: dateFrom + ' 00:00:00',
          date_till: dateTo + ' 23:59:59',
          limit: limit,
          offset: offset,
          fields: activeFields
        }, token);

        const response = DB_VAG_V04_fetchUis_(payload);
        const json = DB_VAG_V04_parseUisJson_(response.text);
        const pageRows = DB_VAG_V04_extractUisRows_(json);

        rows = rows.concat(pageRows);
        if (customLimit || pageRows.length < limit) return rows;
        offset += limit;
      }
    } catch (error) {
      const removedField = DB_VAG_V04_removeInvalidUisField_(activeFields, error);
      if (!removedField) throw error;
      Logger.log('UIS field skipped for ' + method + ': ' + removedField);
    }
  }
}

function DB_VAG_V04_buildUisRequest_(method, params, token) {
  const safeParams = Object.assign({}, params, { access_token: token });
  return {
    jsonrpc: '2.0',
    id: Utilities.getUuid(),
    method: method,
    params: safeParams
  };
}

function DB_VAG_V04_fetchUis_(payload) {
  const response = UrlFetchApp.fetch(DB_VAG_V04_CONFIG.UIS_API_URL, {
    method: 'post',
    contentType: 'application/json',
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  });
  const code = response.getResponseCode();
  const text = response.getContentText();

  if (code < 200 || code >= 300) {
    throw new Error('Ошибка UIS API: ' + code + '\n' + DB_VAG_V04_maskTokenInText_(text));
  }

  return { code: code, text: text };
}

function DB_VAG_V04_parseUisJson_(text) {
  const json = JSON.parse(text || '{}');
  if (json.error) {
    if (json.error.data && json.error.data.mnemonic === 'access_token_invalid') {
      throw new Error('UIS API отклонил токен: access_token_invalid. Свойство UIS_TOKEN найдено, но его значение недействительное. Замени значение UIS_TOKEN в Script Properties на действующий API-токен UIS и запусти testUisApiConnection() ещё раз.');
    }
    if (json.error.data && json.error.data.mnemonic === 'ip_not_whitelisted') {
      const ip = json.error.data.params && json.error.data.params.ip ? json.error.data.params.ip : 'неизвестен';
      throw new Error('UIS API отклонил запрос: IP не добавлен в whitelist. UIS увидел IP Google Apps Script: ' + ip + '. Добавь этот IP в разрешённые IP для API-токена UIS или отключи IP-ограничение для этого токена, затем запусти testUisApiConnection() ещё раз.');
    }
    throw new Error('Ошибка UIS API: ' + JSON.stringify(json.error));
  }
  return json;
}

function DB_VAG_V04_extractUisRows_(json) {
  const result = json && json.result ? json.result : json;
  if (Array.isArray(result)) return result;
  if (result && Array.isArray(result.data)) return result.data;
  if (result && Array.isArray(result.rows)) return result.rows;
  if (result && Array.isArray(result.items)) return result.items;
  return [];
}

function DB_VAG_V04_getUisCallFields_() {
  return [
    'start_time',
    'communication_type',
    'communication_type_name',
    'utm_source',
    'utm_medium',
    'utm_campaign',
    'referrer',
    'search_engine',
    'visitor_source'
  ];
}

function DB_VAG_V04_getUisLeadFields_() {
  return [
    'date_time',
    'create_time',
    'communication_type',
    'communication_type_name',
    'utm_source',
    'utm_medium',
    'utm_campaign',
    'referrer',
    'search_engine',
    'visitor_source'
  ];
}

function DB_VAG_V04_getUisRowDate_(row) {
  const value = row.start_time || row.date_time || row.create_time || row.date || '';
  const match = String(value).match(/\d{4}-\d{2}-\d{2}/);
  return match ? match[0] : '';
}

function DB_VAG_V04_isQualityUisRow_(row) {
  const quality = DB_VAG_V04_normalize_(DB_VAG_V04_CONFIG.UIS_QUALITY_VALUE);
  const values = [
    row.communication_type,
    row.communication_type_name,
    row.appeal_type,
    row.appeal_type_name,
    row.contact_type,
    row.contact_type_name
  ];
  return values.some(value => DB_VAG_V04_normalize_(value) === quality);
}

function DB_VAG_V04_isYandexDirectUisRow_(row) {
  const values = [
    row.utm_source,
    row.utm_medium,
    row.utm_campaign,
    row.referrer,
    row.search_engine,
    row.visitor_source,
    row.source,
    row.source_name,
    row.channel,
    row.channel_name
  ].join(' ').toLowerCase();

  return values.indexOf('yandex') !== -1 ||
    values.indexOf('яндекс') !== -1 ||
    values.indexOf('direct') !== -1 ||
    values.indexOf('директ') !== -1;
}

function DB_VAG_V04_makeEmptyUisDailyMap_(dateFrom, dateTo) {
  const result = {};
  const current = DB_VAG_V04_parseApiDate_(dateFrom);
  const end = DB_VAG_V04_parseApiDate_(dateTo);

  while (current <= end) {
    const date = Utilities.formatDate(current, DB_VAG_V04_CONFIG.TIMEZONE, 'yyyy-MM-dd');
    result[date] = DB_VAG_V04_emptyUisDailyRow_();
    current.setDate(current.getDate() + 1);
  }

  return result;
}

function DB_VAG_V04_emptyUisDailyRow_() {
  return {
    callsTotalFromUis: 0,
    qualityCallsFromUis: 0,
    leadsFromUis: 0
  };
}

function DB_VAG_V04_maskUisRowForLog_(row) {
  const safe = {};
  Object.keys(row || {}).forEach(key => {
    const lower = String(key).toLowerCase();
    if (lower.indexOf('phone') !== -1 || lower.indexOf('email') !== -1 || lower.indexOf('name') !== -1 || lower.indexOf('client') !== -1) {
      safe[key] = '[hidden]';
    } else {
      safe[key] = row[key];
    }
  });
  return safe;
}

function DB_VAG_V04_collectUisSourceDebugValues_(rows) {
  const sourceKeys = [
    'utm_source',
    'utm_medium',
    'utm_campaign',
    'referrer',
    'search_engine',
    'source',
    'source_name',
    'channel',
    'channel_name',
    'communication_type',
    'appeal_type',
    'contact_type'
  ];
  const found = {};

  rows.slice(0, 50).forEach(row => {
    sourceKeys.forEach(key => {
      if (row && row[key] !== undefined && row[key] !== null && String(row[key]).trim() !== '') {
        if (!found[key]) found[key] = [];
        const value = String(row[key]).trim();
        if (found[key].indexOf(value) === -1 && found[key].length < 10) {
          found[key].push(value);
        }
      }
    });
  });

  return found;
}

function DB_VAG_V04_maskTokenInText_(text) {
  const token = PropertiesService.getScriptProperties().getProperty(DB_VAG_V04_CONFIG.UIS_TOKEN_PROPERTY);
  return token ? String(text || '').split(token).join('[UIS_TOKEN]') : String(text || '');
}

function DB_VAG_V04_removeInvalidUisField_(fields, error) {
  const message = String(error && error.message ? error.message : error);
  if (message.indexOf('invalid_parameter_value') === -1) return '';

  const match = message.match(/"field"\s*:\s*"fields\[(\d+)\]"/);
  if (!match) return '';

  const index = Number(match[1]);
  if (isNaN(index) || index < 0 || index >= fields.length) return '';

  const removed = fields.splice(index, 1)[0];
  if (!fields.length) {
    throw new Error('UIS API отклонил все запрошенные поля для отчёта. Последнее отклонённое поле: ' + removed);
  }
  return removed;
}

function DB_VAG_V04_isProtectedDbRow_(row) {
  return row === DB_VAG_V04_CONFIG.PROTECTED_ROWS.SKIP_ROW ||
    row >= DB_VAG_V04_CONFIG.PROTECTED_ROWS.SKIP_FROM_ROW;
}

function DB_VAG_V04_assertMonthBlockSafeToRebuild_(startRow, endRow) {
  if (startRow <= DB_VAG_V04_CONFIG.PROTECTED_ROWS.SKIP_ROW && endRow >= DB_VAG_V04_CONFIG.PROTECTED_ROWS.SKIP_ROW) {
    throw new Error('Пересборка блока остановлена: диапазон затрагивает защищенную строку 234. Для текущего обновления используй fillDbYesterdayDaily_v4().');
  }
  if (endRow >= DB_VAG_V04_CONFIG.PROTECTED_ROWS.SKIP_FROM_ROW) {
    throw new Error('Пересборка блока остановлена: диапазон затрагивает строки 242 и ниже. Для текущего обновления используй fillDbYesterdayDaily_v4().');
  }
}

function DB_VAG_V04_updateCurrentMonthDailySafe_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(DB_VAG_V04_CONFIG.SHEET_NAME);
  if (!sheet) throw new Error('Не найден лист ' + DB_VAG_V04_CONFIG.SHEET_NAME);

  const now = new Date();
  const year = Number(Utilities.formatDate(now, DB_VAG_V04_CONFIG.TIMEZONE, 'yyyy'));
  const month = Number(Utilities.formatDate(now, DB_VAG_V04_CONFIG.TIMEZONE, 'M'));
  const monthLabel = DB_VAG_V04_getMonthLabel_(year, month);
  const topRow = DB_VAG_V04_findMonthTopRowByLabel_(sheet, monthLabel);
  if (!topRow) throw new Error('Не найден блок ДБ "' + monthLabel + '". Сначала создай месяц через createDbCurrentMonthFromMayTemplate_v4().');

  const headerRow = DB_VAG_V04_findHeaderRowInBlock_(sheet, topRow);
  const totalRow = DB_VAG_V04_findTotalRowAfter_(sheet, headerRow);
  if (!headerRow || !totalRow) throw new Error('Не найдены шапка или строка ИТОГО в блоке "' + monthLabel + '".');

  const period = DB_VAG_V04_getMonthPeriodForReport_(year, month);
  const directByDate = DB_VAG_V04_getDirectDailyStats_(period.dateFrom, period.dateTo);
  const goalsByDate = DB_VAG_V04_getMetrikaGoalsByDate_(period.dateFrom, period.dateTo);
  const uisByDate = DB_VAG_V04_getUisDailyDataForDb_(period.dateFrom, period.dateTo);
  const dataStartRow = headerRow + 1;
  const reportEndDate = DB_VAG_V04_parseApiDate_(period.dateTo);
  const daysInMonth = DB_VAG_V04_getDaysInMonth_(year, month);

  for (let day = 1; day <= daysInMonth; day++) {
    const date = new Date(year, month - 1, day);
    if (date > reportEndDate) break;

    const row = dataStartRow + day - 1;
    if (row >= totalRow || DB_VAG_V04_isProtectedDbRow_(row)) {
      Logger.log('DB row skipped by protection: ' + row);
      continue;
    }

    const apiDate = Utilities.formatDate(date, DB_VAG_V04_CONFIG.TIMEZONE, 'yyyy-MM-dd');
    const direct = directByDate[apiDate] || { impressions: '', clicks: '', cost: '' };
    const goals = goalsByDate[apiDate] || DB_VAG_V04_emptyGoals_();
    const uis = uisByDate[apiDate] || DB_VAG_V04_emptyUisDailyRow_();

    sheet.getRange(row, 3, 1, 2).setValues([[
      DB_VAG_V04_blankZero_(direct.impressions),
      DB_VAG_V04_blankZero_(direct.clicks)
    ]]);
    sheet.getRange(row, 6).setValue(DB_VAG_V04_blankZero_(direct.cost));
    sheet.getRange(row, 9, 1, 9).setValues([[
      DB_VAG_V04_blankZero_(goals.bb60sec2pages),
      DB_VAG_V04_blankZero_(goals.phone8835),
      DB_VAG_V04_blankZero_(goals.phone8865),
      DB_VAG_V04_blankZero_(goals.max),
      DB_VAG_V04_blankZero_(goals.telegramSlider),
      DB_VAG_V04_blankZero_(goals.route),
      DB_VAG_V04_blankZero_(uis.callsTotalFromUis),
      DB_VAG_V04_blankZero_(uis.qualityCallsFromUis),
      DB_VAG_V04_blankZero_(uis.leadsFromUis)
    ]]);
  }

  SpreadsheetApp.flush();
  Logger.log('Готово: безопасно обновлены дневные значения ДБ за период ' + period.dateFrom + ' — ' + period.dateTo + '. Формулы, итоги, строка 234 и строки 242+ не тронуты.');
}

// =====================
// ФОРМАТИРОВАНИЕ
// =====================

function DB_VAG_V04_applyFormatting_(sheet, topRow, headerRow, dataStartRow, totalRow, template, year, month) {
  const width = DB_VAG_V04_CONFIG.TABLE_WIDTH;
  const startCol = DB_VAG_V04_CONFIG.START_COLUMN;
  const daysRows = totalRow - dataStartRow;
  const serviceRows = template.afterTotalRows;
  const rowsCount = totalRow - topRow + 1 + serviceRows;

  // Сохраняем формат майского шаблона, но проставляем нужные форматы чисел и дат.
  sheet.getRange(topRow, startCol, rowsCount, width)
    .setVerticalAlignment('middle')
    .setBorder(true, true, true, true, true, true, '#000000', SpreadsheetApp.BorderStyle.SOLID);

  sheet.getRange(headerRow, startCol, 1, width)
    .setFontWeight('bold')
    .setWrap(true)
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle');

  sheet.getRange(totalRow, startCol, 1, width)
    .setFontWeight('bold');

  sheet.getRange(dataStartRow, 2, daysRows, 1).setNumberFormat('dd.mm.yyyy');
  sheet.getRange(dataStartRow, 3, daysRows + 1, 2).setNumberFormat('#,##0');
  sheet.getRange(dataStartRow, 5, daysRows + 1, 2).setNumberFormat('#,##0.00');
  sheet.getRange(dataStartRow, 7, daysRows + 1, 1).setNumberFormat('0.00%');
  sheet.getRange(dataStartRow, 8, daysRows + 1, 1).setNumberFormat('#,##0.00');
  sheet.getRange(dataStartRow, 9, daysRows + 1, 10).setNumberFormat('#,##0');
  sheet.getRange(dataStartRow, 19, daysRows + 1, 1).setNumberFormat('0.00%');
  sheet.getRange(dataStartRow, 20, daysRows + 1, 1).setNumberFormat('#,##0.00');

  sheet.getRange(topRow, startCol, rowsCount, width).setHorizontalAlignment('center');
  sheet.getRange(topRow, startCol, rowsCount, 1).setHorizontalAlignment('left');

  // Чёрные контуры как в майском блоке: верхний блок, таблица, ИТОГО и служебные строки.
  sheet.getRange(topRow, startCol, rowsCount, width)
    .setBorder(true, true, true, true, true, true, '#000000', SpreadsheetApp.BorderStyle.SOLID);

  // Ширины колонок подтягиваем из майского блока, чтобы июнь выглядел так же.
  for (let col = startCol; col < startCol + width; col++) {
    sheet.setColumnWidth(col, sheet.getColumnWidth(col));
  }

  // Рабочие дни белые, выходные серые — как в обычном ДБ-блоке.
  for (let i = 0; i < daysRows; i++) {
    const row = dataStartRow + i;
    const date = new Date(year, month - 1, i + 1);
    const isWeekend = date.getDay() === 0 || date.getDay() === 6;
    sheet.getRange(row, startCol, 1, width).setBackground(isWeekend ? '#f3f3f3' : '#ffffff');
  }

  // Шапка и ИТОГО остаются с форматами из мая; дополнительно фиксируем читаемость.
  sheet.getRange(headerRow, startCol, 1, width).setWrap(true);
  sheet.getRange(totalRow, startCol, 1, width).setFontWeight('bold');
}


// =====================
// УТИЛИТЫ
// =====================

function DB_VAG_V04_getMonthLabel_(year, month) {
  const names = ['Январь', 'Февраль', 'Март', 'Апрель', 'Май', 'Июнь', 'Июль', 'Август', 'Сентябрь', 'Октябрь', 'Ноябрь', 'Декабрь'];
  return names[month - 1] + ' ' + year;
}

function DB_VAG_V04_getDaysInMonth_(year, month) {
  return new Date(year, month, 0).getDate();
}

function DB_VAG_V04_getDaysBetween_(dateFrom, dateTo) {
  const start = DB_VAG_V04_parseApiDate_(dateFrom);
  const end = DB_VAG_V04_parseApiDate_(dateTo);
  return Math.round((end - start) / (24 * 60 * 60 * 1000)) + 1;
}

function DB_VAG_V04_formatDateRuFull_(apiDate) {
  const parts = String(apiDate).split('-');
  return parts[2] + '.' + parts[1] + '.' + parts[0];
}

function DB_VAG_V04_parseApiDate_(apiDate) {
  const parts = String(apiDate).split('-').map(Number);
  return new Date(parts[0], parts[1] - 1, parts[2]);
}

function DB_VAG_V04_toNumber_(value) {
  if (value === null || value === undefined || value === '' || value === '-' || value === '--') return 0;
  const text = String(value).replace(/\s/g, '').replace(',', '.').replace(/[^\d.-]/g, '');
  const number = Number(text);
  return isNaN(number) ? 0 : number;
}

function DB_VAG_V04_blankZero_(value) {
  const number = Number(value || 0);
  return number > 0 ? number : '';
}

function DB_VAG_V04_normalize_(value) {
  return String(value || '')
    .toLowerCase()
    .replace(/ё/g, 'е')
    .replace(/[«»"']/g, '')
    .replace(/\s+/g, ' ')
    .trim();
}

function DB_VAG_V04_columnToLetter_(column) {
  let temp = '';
  let letter = '';

  while (column > 0) {
    temp = (column - 1) % 26;
    letter = String.fromCharCode(temp + 65) + letter;
    column = (column - temp - 1) / 26;
  }

  return letter;
}

function DB_VAG_V04_toQueryString_(params) {
  return Object.keys(params)
    .map(key => encodeURIComponent(key) + '=' + encodeURIComponent(params[key]))
    .join('&');
}
// =====================
// ЕЖЕДНЕВНЫЙ АВТОЗАПУСК ДБ В 8:00
// Добавь этот блок в конец выбранного скрипта DB v04_1.
//
// Что запускать:
// 1) createDbDaily8amTrigger_v4() — один раз создать ежедневный автозапуск.
// 2) deleteDbDaily8amTriggers_v4() — удалить старый автозапуск, если понадобится.
// 3) fillDbYesterdayDaily_v4() — ручной запуск обновления текущего месяца по вчерашний день.
// =====================

function fillDbYesterdayDaily_v4() {
  // Безопасно обновляет текущий месяц с 1 числа по вчерашний день.
  // Не очищает блок, не трогает формулы, итоги, строку 234 и строки 242+.
  DB_VAG_V04_updateCurrentMonthDailySafe_();
}

function createDbDaily8amTrigger_v4() {
  deleteDbDaily8amTriggers_v4();

  ScriptApp.newTrigger('fillDbYesterdayDaily_v4')
    .timeBased()
    .everyDays(1)
    .atHour(8)
    .create();

  Logger.log('Готово: ежедневный автозапуск ДБ создан. Отчёт будет обновляться каждый день около 08:00 по текущему месяцу до вчерашнего дня.');
}

function deleteDbDaily8amTriggers_v4() {
  const handlers = [
    'fillDbYesterdayDaily_v4',
    'createDbCurrentMonthFromMayTemplate_v4'
  ];

  ScriptApp.getProjectTriggers()
    .filter(trigger => handlers.indexOf(trigger.getHandlerFunction()) !== -1)
    .forEach(trigger => ScriptApp.deleteTrigger(trigger));

  Logger.log('Готово: старые ежедневные триггеры ДБ удалены.');
}

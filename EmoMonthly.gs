// EMO Monthly Report v13
// Вставь этот код целиком в Apps Script вместо текущего файла ЕМО.
// Исправлено:
// 1) чистим весь блок "Доля по лидам" 228–249, чтобы не оставался хвост ИТОГО 125;
// 2) удаляем старые/пустые графики рядом с блоком;
// 3) не используем merge(), чтобы не было ошибки с закреплённым столбцом A;
// 4) логин уже стоит vag-house98;
// 5) комментарии оформляются слева, текст комментариев пишется в столбец A;
// 6) график по источникам ставится напротив блока «Доля по лидам»;
// 7) блок «Доля по лидам» считает те же основные цели, что и общий итог по лидам.

const EMO_CONFIG = {
  SHEET_NAME: 'ЕМО',
  CLIENT_LOGIN: 'vag-house98',
  TIMEZONE: 'Europe/Moscow',
  ATTRIBUTION_MODEL: 'AUTO',

  TEMPLATE_START_ROW: 76,
  TEMPLATE_END_ROW: 158,
  TOTAL_COLS: 25,
  COMMENT_COLS: 6, // комментарии оформляем только A:F
  GAP_ROWS: 3,

  // Для нового майского блока, который начинается примерно со строки 161,
  // блок "Доля по лидам" попадёт в зону 228–249.
  LEAD_SHARE_TEMPLATE_ROW: 143,
  LEAD_SHARE_CLEAR_ROWS: 22, // 228–249

  // Динамика по месяцам с февраля 2026
  DYNAMICS_START_DATE: '2026-02-01',

  GOALS: {
    callMain: null,
    phone1: '511787501',
    phone2: '511788019',
    max: '511788150',
    telegramSlider: '511791937',
    zayavkaUis: '517626592',
    jivoOffline: '530094694',
    callKval: '540780589',
    callsUis: '517626340',
    uniqueCallsUis: '562354501'
  },

  LEAD_SHARE_GOALS: [
    '517626340', // Звонки Юис
    '517626592', // Заявки Юис
    '530094694'  // Jivo: офлайн-сообщение
  ]
};

function fillEmoMayReport() {
  const today = EMO_getToday_();
  let dateTo = today;

  if (dateTo > '2026-05-31') {
    dateTo = '2026-05-31';
  }

  EMO_fillMonthlyReportForPeriod_('2026-05-01', dateTo);
}

function fillEmoCurrentMonthReport() {
  const today = EMO_getToday_();
  const dateFrom = today.slice(0, 8) + '01';

  EMO_fillMonthlyReportForPeriod_(dateFrom, today);
}

function fillEmoPreviousMonthReport() {
  const period = EMO_getPreviousMonthPeriod_();

  EMO_fillMonthlyReportForPeriod_(period.dateFrom, period.dateTo);
}

function createEmoMonthlyTrigger10am() {
  ScriptApp.getProjectTriggers()
    .filter(t => t.getHandlerFunction() === 'fillEmoPreviousMonthReport')
    .forEach(t => ScriptApp.deleteTrigger(t));

  ScriptApp.newTrigger('fillEmoPreviousMonthReport')
    .timeBased()
    .onMonthDay(1)
    .atHour(10)
    .create();
}

function EMO_fillMonthlyReportForPeriod_(dateFrom, dateTo) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(EMO_CONFIG.SHEET_NAME);

  if (!sheet) {
    throw new Error('Не найдена вкладка: ' + EMO_CONFIG.SHEET_NAME);
  }

  const rows = EMO_getMonthlyCampaignReport_(dateFrom, dateTo);
  rows.sort((a, b) => b.cost - a.cost);

  const block = EMO_getOrCreateBlock_(sheet, dateFrom, dateTo);
  const totalRow = EMO_prepareBlock_(sheet, block, rows.length);
  EMO_extendTotalRowFormulas_(sheet, block.headerRow, totalRow);

  EMO_writeTitle_(sheet, block, dateFrom, dateTo);
  EMO_writeRawData_(sheet, block.headerRow, rows);
  EMO_writeCommentBlock_(sheet, block, totalRow, rows, dateFrom, dateTo);
  EMO_writeLeadShareFromMetrika_(sheet, block, dateFrom, dateTo);
  EMO_updateMonthlyDynamicsFromFebruary_(dateTo);
}

// ================== ДИРЕКТ ==================

function EMO_getMonthlyCampaignReport_(dateFrom, dateTo) {
  const goalIds = Object.values(EMO_CONFIG.GOALS)
    .filter(id => id)
    .map(id => Number(id));

  const reportBody = {
    params: {
      SelectionCriteria: {
        DateFrom: dateFrom,
        DateTo: dateTo
      },
      Goals: goalIds,
      AttributionModels: [
        EMO_CONFIG.ATTRIBUTION_MODEL
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
      ReportName: 'emo_monthly_' + EMO_CONFIG.CLIENT_LOGIN + '_' + dateFrom + '_' + dateTo + '_' + Utilities.getUuid(),
      ReportType: 'CAMPAIGN_PERFORMANCE_REPORT',
      DateRangeType: 'CUSTOM_DATE',
      Format: 'TSV',
      IncludeVAT: 'YES',
      IncludeDiscount: 'NO'
    }
  };

  const text = EMO_requestDirectReport_(reportBody);
  return EMO_parseCampaignTsv_(text);
}

function EMO_requestDirectReport_(reportBody) {
  const token = PropertiesService.getScriptProperties().getProperty('YANDEX_TOKEN');

  if (!token) {
    throw new Error('Не найден YANDEX_TOKEN');
  }

  const url = 'https://api.direct.yandex.com/json/v5/reports';

  const options = {
    method: 'post',
    contentType: 'application/json',
    headers: {
      Authorization: 'Bearer ' + token,
      'Client-Login': EMO_CONFIG.CLIENT_LOGIN,
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

    if (code === 200) {
      return text;
    }

    if (code === 201 || code === 202) {
      Utilities.sleep(5000);
      continue;
    }

    throw new Error('Ошибка Директа: ' + code + '\n' + text);
  }

  throw new Error('Отчёт долго формируется. Запусти позже.');
}

function EMO_parseCampaignTsv_(tsvText) {
  const lines = String(tsvText || '').trim().split(/\r?\n/);

  if (lines.length < 2) return [];

  const headers = lines[0]
    .split('\t')
    .map(h => String(h).replace(/^\uFEFF/, '').trim());

  const idx = {};
  headers.forEach((h, i) => idx[h] = i);

  const rows = lines.slice(1).filter(line => String(line).trim() !== '');
  const result = [];

  rows.forEach(line => {
    const row = line.split('\t');

    const campaignName = String(row[idx['CampaignName']] || '').trim();
    if (!campaignName) return;

    const impressions = EMO_toNumber_(row[idx['Impressions']]);
    const clicks = EMO_toNumber_(row[idx['Clicks']]);
    const cost = EMO_toNumber_(row[idx['Cost']]);

    const phone1 = EMO_goalValue_(row, idx, EMO_CONFIG.GOALS.phone1);
    const phone2 = EMO_goalValue_(row, idx, EMO_CONFIG.GOALS.phone2);
    const max = EMO_goalValue_(row, idx, EMO_CONFIG.GOALS.max);
    const telegramSlider = EMO_goalValue_(row, idx, EMO_CONFIG.GOALS.telegramSlider);
    const zayavkaUis = EMO_goalValue_(row, idx, EMO_CONFIG.GOALS.zayavkaUis);
    const jivoOffline = EMO_goalValue_(row, idx, EMO_CONFIG.GOALS.jivoOffline);
    const callKval = EMO_goalValue_(row, idx, EMO_CONFIG.GOALS.callKval);
    const callsUis = EMO_goalValue_(row, idx, EMO_CONFIG.GOALS.callsUis);
    const uniqueCallsUis = EMO_goalValue_(row, idx, EMO_CONFIG.GOALS.uniqueCallsUis);

    const callMain = EMO_CONFIG.GOALS.callMain
      ? EMO_goalValue_(row, idx, EMO_CONFIG.GOALS.callMain)
      : 0;

    const kvalPlusZayavki = callKval + zayavkaUis;

    result.push({
      campaignName,
      campaignType: EMO_getCampaignType_(campaignName),

      impressions,
      clicks,
      cost,

      avgCpc: idx['AvgCpc'] !== undefined ? EMO_toNumber_(row[idx['AvgCpc']]) : 0,
      avgPosition: idx['AvgImpressionPosition'] !== undefined ? EMO_toNumber_(row[idx['AvgImpressionPosition']]) : 0,
      avgTrafficVolume: idx['AvgTrafficVolume'] !== undefined ? EMO_toNumber_(row[idx['AvgTrafficVolume']]) : 0,
      bounceRate: idx['BounceRate'] !== undefined ? EMO_toNumber_(row[idx['BounceRate']]) : 0,
      avgPageviews: idx['AvgPageviews'] !== undefined ? EMO_toNumber_(row[idx['AvgPageviews']]) : 0,

      callMain,
      phone1,
      phone2,
      max,
      telegramSlider,

      kvalPlusZayavki,
      jivoOffline,
      callsUis,
      uniqueCallsUis,
      zayavkaUis
    });
  });

  return result.filter(r => r.impressions > 0 || r.clicks > 0 || r.cost > 0 || EMO_rowHasGoals_(r));
}

function EMO_goalValue_(row, idx, goalId) {
  if (!goalId) return 0;

  const exact = 'Conversions_' + goalId + '_' + EMO_CONFIG.ATTRIBUTION_MODEL;

  if (idx[exact] !== undefined) {
    return EMO_toNumber_(row[idx[exact]]);
  }

  const prefix = 'Conversions_' + goalId + '_';

  for (const header in idx) {
    if (String(header).indexOf(prefix) === 0) {
      return EMO_toNumber_(row[idx[header]]);
    }
  }

  return 0;
}

// ================== БЛОК ЕМО ==================

function EMO_getOrCreateBlock_(sheet, dateFrom, dateTo) {
  const existing = EMO_findBlockByPeriod_(sheet, dateFrom, dateTo);

  if (existing) return existing;

  return EMO_createNewBlock_(sheet);
}

function EMO_createNewBlock_(sheet) {
  const template = EMO_getTemplateBlock_(sheet);
  const lastEndRow = EMO_findLastBlockEndRow_(sheet);

  const newTopRow = lastEndRow + EMO_CONFIG.GAP_ROWS + 1;
  const templateRowsCount = EMO_CONFIG.TEMPLATE_END_ROW - EMO_CONFIG.TEMPLATE_START_ROW + 1;
  const neededLastRow = newTopRow + templateRowsCount - 1;

  if (sheet.getMaxRows() < neededLastRow) {
    sheet.insertRowsAfter(sheet.getMaxRows(), neededLastRow - sheet.getMaxRows());
  }

  sheet
    .getRange(EMO_CONFIG.TEMPLATE_START_ROW, 1, templateRowsCount, EMO_CONFIG.TOTAL_COLS)
    .copyTo(
      sheet.getRange(newTopRow, 1, templateRowsCount, EMO_CONFIG.TOTAL_COLS),
      SpreadsheetApp.CopyPasteType.PASTE_NORMAL,
      false
    );

  return {
    topRow: newTopRow,
    headerRow: newTopRow + (template.headerRow - template.topRow),
    totalRow: newTopRow + (template.totalRow - template.topRow),
    endRow: newTopRow + templateRowsCount - 1
  };
}

function EMO_getTemplateBlock_(sheet) {
  const topRow = EMO_CONFIG.TEMPLATE_START_ROW;
  const headerRow = EMO_findHeaderRowAfter_(sheet, topRow);
  const totalRow = EMO_findTotalRowAfter_(sheet, headerRow);

  if (!headerRow || !totalRow) {
    throw new Error('Не нашла шаблон ЕМО. Проверь строки 76–158, строку РК и строку ИТОГО.');
  }

  return {
    topRow,
    headerRow,
    totalRow,
    endRow: EMO_CONFIG.TEMPLATE_END_ROW
  };
}

function EMO_prepareBlock_(sheet, block, rowsNeeded) {
  let totalRow = EMO_findTotalRowAfter_(sheet, block.headerRow);
  const dataStartRow = block.headerRow + 1;
  const existingRows = totalRow - dataStartRow;

  if (rowsNeeded > existingRows) {
    const rowsToAdd = rowsNeeded - existingRows;

    sheet.insertRowsBefore(totalRow, rowsToAdd);

    sheet
      .getRange(dataStartRow, 1, 1, EMO_CONFIG.TOTAL_COLS)
      .copyTo(
        sheet.getRange(totalRow, 1, rowsToAdd, EMO_CONFIG.TOTAL_COLS),
        SpreadsheetApp.CopyPasteType.PASTE_NORMAL,
        false
      );

    totalRow += rowsToAdd;
  }

  const dataRowsCount = totalRow - dataStartRow;

  if (dataRowsCount > 0) {
    // Чистим только сырые данные.
    // Формулы E, F, S, W, X, Y не трогаем.
    sheet.getRange(dataStartRow, 1, dataRowsCount, 4).clearContent();   // A:D
    sheet.getRange(dataStartRow, 7, dataRowsCount, 12).clearContent();  // G:R
    sheet.getRange(dataStartRow, 20, dataRowsCount, 3).clearContent();  // T:V
  }

  return totalRow;
}

function EMO_writeTitle_(sheet, block, dateFrom, dateTo) {
  const title = EMO_buildTitle_(dateFrom, dateTo);
  sheet.getRange(block.topRow, 1).setValue(title);
}

function EMO_writeRawData_(sheet, headerRow, rows) {
  const dataStartRow = headerRow + 1;

  if (!rows.length) return;

  rows.forEach((item, index) => {
    const row = dataStartRow + index;

    // A:D
    sheet.getRange(row, 1, 1, 4).setValues([[
      item.campaignName,
      item.campaignType,
      item.impressions,
      item.clicks
    ]]);

    // G:R
    sheet.getRange(row, 7, 1, 12).setValues([[
      item.cost,
      EMO_valueOrDash_(item.avgCpc),
      EMO_valueOrDash_(item.avgPosition),
      EMO_valueOrDash_(item.avgTrafficVolume),
      EMO_valueOrDash_(item.bounceRate),
      EMO_valueOrDash_(item.avgPageviews),
      EMO_valueOrDash_(item.callMain),
      EMO_valueOrDash_(item.phone1),
      EMO_valueOrDash_(item.phone2),
      EMO_valueOrDash_(item.max),
      EMO_valueOrDash_(item.telegramSlider),
      EMO_valueOrDash_(item.kvalPlusZayavki)
    ]]);

    // T:V
    sheet.getRange(row, 20, 1, 3).setValues([[
      EMO_valueOrDash_(item.jivoOffline),
      EMO_valueOrDash_(item.callsUis),
      EMO_valueOrDash_(item.zayavkaUis)
    ]]);
  });
}

// ================== КОММЕНТАРИИ ==================


function EMO_writeCommentBlock_(sheet, block, totalRow, rows, dateFrom, dateTo) {
  const commentStartRow = totalRow + 3;
  const leadShareRow = EMO_getLeadShareRow_(block.topRow);
  const commentEndRow = Math.max(commentStartRow + 38, leadShareRow - 4);
  const rowsToClear = Math.max(1, commentEndRow - commentStartRow + 1);

  EMO_clearCommentArea_(sheet, commentStartRow, rowsToClear);

  const totals = EMO_calcTotals_(rows);
  const prevPeriod = EMO_getPreviousMonthForPeriod_(dateFrom);
  let prevRows = [];

  try {
    prevRows = EMO_getMonthlyCampaignReport_(prevPeriod.dateFrom, prevPeriod.dateTo);
  } catch (e) {
    prevRows = [];
  }

  const prevTotals = EMO_calcTotals_(prevRows);
  const bestRows = EMO_getBestRows_(rows);
  const weakRows = EMO_getWeakRows_(rows, totals);
  const typeLines = EMO_buildTypeLines_(rows);

  let row = commentStartRow;

  EMO_writeCommentTitle_(sheet, row++, 'Комментарий к отчёту за период с ' + EMO_formatShortDate_(dateFrom) + ' по ' + EMO_formatShortDateFull_(dateTo));
  row++;

  EMO_writeCommentSection_(sheet, row++, 'Итог месяца');
  EMO_writeCommentText_(sheet, row++, 'За период потрачено ' + EMO_formatRub_(totals.cost) + ', получено ' + EMO_formatNumber_(totals.clicks) + ' кликов и ' + totals.factConversions + ' лидов. Средний CPA — ' + EMO_formatRub_(totals.factConversions > 0 ? totals.cost / totals.factConversions : 0) + ', CR — ' + EMO_formatPercent_(totals.clicks > 0 ? totals.factConversions / totals.clicks : 0) + '.');
  row++;

  EMO_writeCommentSection_(sheet, row++, 'Сравнение с прошлым месяцем');
  EMO_writeCommentText_(sheet, row++, EMO_buildCompareText_(totals, prevTotals));
  row++;

  EMO_writeCommentSection_(sheet, row++, 'Бюджет и трафик');
  EMO_writeCommentText_(sheet, row++, 'Расход — ' + EMO_formatRub_(totals.cost) + ', показы — ' + EMO_formatNumber_(totals.impressions) + ', клики — ' + EMO_formatNumber_(totals.clicks) + ', средний CPC — ' + EMO_formatRub_(totals.clicks > 0 ? totals.cost / totals.clicks : 0) + '.');
  row++;

  EMO_writeCommentSection_(sheet, row++, 'Достигнуты цели в Директ / UIS / чат');
  EMO_writeCommentText_(sheet, row++, 'Конверсий получено — ' + totals.factConversions + ', CPA — ' + EMO_formatRub_(totals.factConversions > 0 ? totals.cost / totals.factConversions : 0) + ', CR — ' + EMO_formatPercent_(totals.clicks > 0 ? totals.factConversions / totals.clicks : 0) + '. Jivo — ' + totals.jivoOffline + ', заявки Юис — ' + totals.zayavkaUis + ', звонки Юис — ' + totals.callsUis + '.');
  row++;

  EMO_writeCommentSection_(sheet, row++, 'Квал лиды + заявки');
  EMO_writeCommentText_(sheet, row++, 'Квал лиды — ' + totals.callKval + ', заявки — ' + totals.zayavkaUis + '. ИТОГО: ' + totals.kvalPlusZayavki + ' конверсий, CPA — ' + EMO_formatRub_(totals.kvalPlusZayavki > 0 ? totals.cost / totals.kvalPlusZayavki : 0) + '.');
  row++;

  EMO_writeCommentSection_(sheet, row++, 'По типам РК');
  typeLines.forEach(line => EMO_writeCommentText_(sheet, row++, line));
  row++;

  EMO_writeCommentSection_(sheet, row++, 'Лучшие кампании');
  bestRows.forEach(line => EMO_writeCommentText_(sheet, row++, line));
  row++;

  EMO_writeCommentSection_(sheet, row++, 'Зона внимания', '#f4cccc');
  EMO_writeCommentText_(sheet, row++, 'Сюда попадают кампании с заметным расходом без конверсий: от 1 000 руб. или от половины среднего CPA по отчёту. Также подсвечиваются кампании, где CPA сильно выше среднего.');
  weakRows.forEach(line => EMO_writeCommentText_(sheet, row++, line));
  row++;

  EMO_writeCommentSection_(sheet, row++, 'План на следующий месяц');
  EMO_writeCommentText_(sheet, row++, 'Удержать CPA, усилить рабочие кампании, проверить кампании с расходом без лидов и отдельно посмотреть качество обращений по UIS/CRM.');
}


function EMO_clearCommentArea_(sheet, startRow, rowsCount) {
  // Сначала убираем старую широкую заливку по всей строке, чтобы не было зелёных полос на весь лист.
  const fullRange = sheet.getRange(startRow, 1, rowsCount, EMO_CONFIG.TOTAL_COLS);
  fullRange.breakApart();
  fullRange
    .clearContent()
    .setBackground('#ffffff')
    .setFontWeight('normal')
    .setFontColor('#000000')
    .setWrap(false)
    .setBorder(false, false, false, false, false, false);

  // Новый блок оформляем только слева: A:F.
  const commentRange = sheet.getRange(startRow, 1, rowsCount, EMO_CONFIG.COMMENT_COLS);
  commentRange
    .setWrap(true)
    .setVerticalAlignment('middle');
}


function EMO_writeCommentTitle_(sheet, row, text) {
  const range = sheet.getRange(row, 1, 1, EMO_CONFIG.COMMENT_COLS);
  range.breakApart();
  range
    .setBackground('#b6d7a8')
    .setFontWeight('bold')
    .setFontSize(10)
    .setBorder(true, true, true, true, false, false)
    .setVerticalAlignment('middle');

  sheet.getRange(row, 1).setValue(text);
  if (EMO_CONFIG.COMMENT_COLS > 1) {
    sheet.getRange(row, 2, 1, EMO_CONFIG.COMMENT_COLS - 1).clearContent();
  }
  sheet.setRowHeight(row, 30);
}


function EMO_writeCommentSection_(sheet, row, text, bg) {
  const color = bg || '#d9ead3';
  const range = sheet.getRange(row, 1, 1, EMO_CONFIG.COMMENT_COLS);
  range.breakApart();
  range
    .setBackground(color)
    .setFontWeight('bold')
    .setFontSize(10)
    .setBorder(true, true, true, true, false, false)
    .setVerticalAlignment('middle');

  sheet.getRange(row, 1).setValue(text);
  if (EMO_CONFIG.COMMENT_COLS > 1) {
    sheet.getRange(row, 2, 1, EMO_CONFIG.COMMENT_COLS - 1).clearContent();
  }
  sheet.setRowHeight(row, 24);
}


function EMO_writeCommentText_(sheet, row, text) {
  // Текст комментария пишем прямо в столбец A.
  // B:F остаются только частью визуального блока, без текста.
  const range = sheet.getRange(row, 1, 1, EMO_CONFIG.COMMENT_COLS);

  range.breakApart();
  range
    .setBackground('#fffdf5')
    .setFontWeight('normal')
    .setFontSize(10)
    .setBorder(true, true, true, true, false, false)
    .setVerticalAlignment('top')
    .setHorizontalAlignment('left')
    .setWrap(true);

  sheet.getRange(row, 1)
    .setValue(text)
    .setHorizontalAlignment('left')
    .setVerticalAlignment('top')
    .setWrap(true);

  if (EMO_CONFIG.COMMENT_COLS > 1) {
    sheet.getRange(row, 2, 1, EMO_CONFIG.COMMENT_COLS - 1).clearContent();
  }

  sheet.setRowHeight(row, 58);
}

function EMO_calcTotals_(rows) {
  return rows.reduce((acc, row) => {
    acc.impressions += row.impressions;
    acc.clicks += row.clicks;
    acc.cost += row.cost;

    acc.jivoOffline += row.jivoOffline;
    acc.callsUis += row.callsUis;
    acc.uniqueCallsUis += row.uniqueCallsUis;
    acc.zayavkaUis += row.zayavkaUis;

    acc.callKval += Math.max(0, row.kvalPlusZayavki - row.zayavkaUis);
    acc.kvalPlusZayavki += row.kvalPlusZayavki;

    acc.factConversions += row.jivoOffline + row.callsUis + row.zayavkaUis;

    return acc;
  }, {
    impressions: 0,
    clicks: 0,
    cost: 0,
    jivoOffline: 0,
    callsUis: 0,
    uniqueCallsUis: 0,
    zayavkaUis: 0,
    callKval: 0,
    kvalPlusZayavki: 0,
    factConversions: 0
  });
}

function EMO_buildCompareText_(totals, prevTotals) {
  if (!prevTotals || prevTotals.cost === 0 && prevTotals.clicks === 0) {
    return 'Данных за прошлый месяц для сравнения пока нет.';
  }

  const currentCpa = totals.factConversions > 0 ? totals.cost / totals.factConversions : 0;
  const prevCpa = prevTotals.factConversions > 0 ? prevTotals.cost / prevTotals.factConversions : 0;

  return 'Расход ' + EMO_compareMoney_(totals.cost, prevTotals.cost) + ', лиды ' + EMO_compareNumber_(totals.factConversions, prevTotals.factConversions) + ', CPA ' + EMO_compareMoney_(currentCpa, prevCpa) + '.';
}

function EMO_buildTypeLines_(rows) {
  const groups = {};

  rows.forEach(row => {
    const type = row.campaignType || 'н/а';

    if (!groups[type]) {
      groups[type] = {
        clicks: 0,
        impressions: 0,
        cost: 0,
        conversions: 0
      };
    }

    groups[type].clicks += row.clicks;
    groups[type].impressions += row.impressions;
    groups[type].cost += row.cost;
    groups[type].conversions += row.jivoOffline + row.callsUis + row.zayavkaUis;
  });

  return Object.entries(groups)
    .sort((a, b) => b[1].cost - a[1].cost)
    .map(([type, data]) => {
      const ctr = data.impressions > 0 ? data.clicks / data.impressions : 0;
      const cpc = data.clicks > 0 ? data.cost / data.clicks : 0;
      const cpa = data.conversions > 0 ? data.cost / data.conversions : 0;

      return type + ' — ' +
        EMO_formatNumber_(data.clicks) + ' кликов, расход ' +
        EMO_formatRub_(data.cost) + ', средний CTR — ' +
        EMO_formatPercent_(ctr) + ', CPC — ' +
        EMO_formatRub_(cpc) + ', ' +
        data.conversions + ' конверсий, CPA — ' +
        EMO_formatRub_(cpa);
    });
}

function EMO_getBestRows_(rows) {
  const result = rows
    .map(row => {
      const conversions = row.jivoOffline + row.callsUis + row.zayavkaUis;
      const cpa = conversions > 0 ? row.cost / conversions : 0;
      return {
        name: row.campaignName,
        conversions,
        cpa
      };
    })
    .filter(item => item.conversions > 0)
    .sort((a, b) => b.conversions - a.conversions || a.cpa - b.cpa)
    .slice(0, 3)
    .map(item => item.name + ' — ' + item.conversions + ' лидов, CPA ' + EMO_formatRub_(item.cpa));

  return result.length ? result : ['Нет кампаний с лидами за период.'];
}


function EMO_getWeakRows_(rows, totals) {
  const avgCpa = totals.factConversions > 0 ? totals.cost / totals.factConversions : 0;

  // Правило для зоны внимания:
  // 1) есть заметный расход, но конверсий нет;
  // 2) CPA сильно выше среднего по отчёту.
  const noConversionCostLimit = Math.max(1000, avgCpa * 0.5);

  const result = rows
    .map(row => {
      const conversions = row.jivoOffline + row.callsUis + row.zayavkaUis;
      const cpa = conversions > 0 ? row.cost / conversions : 0;

      return {
        name: row.campaignName,
        cost: row.cost,
        conversions,
        cpa
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
    .slice(0, 5)
    .map(item => {
      if (item.conversions === 0) {
        return item.name + ' — расход ' + EMO_formatRub_(item.cost) + ', конверсий нет. Проверить запросы, площадки и качество трафика.';
      }

      return item.name + ' — CPA ' + EMO_formatRub_(item.cpa) + ', выше среднего по отчёту.';
    });

  return result.length ? result : ['Критичных кампаний с высоким расходом без лидов не выявлено.'];
}

// ================== ДОЛЯ ПО ЛИДАМ ИЗ МЕТРИКИ ==================

function EMO_writeLeadShareFromMetrika_(sheet, block, dateFrom, dateTo) {
  const sourceRows = EMO_getLeadShareFromMetrika_(dateFrom, dateTo);
  const startRow = EMO_getLeadShareRow_(block.topRow);

  const rows = [
    ['Директ', sourceRows.direct],
    ['Прямые переходы', sourceRows.directVisits],
    ['Яндекс поиск', sourceRows.yandexSearch],
    ['Яндекс карты', sourceRows.yandexMaps],
    ['Google Search', sourceRows.googleSearch],
    ['Телеграм', sourceRows.telegram],
    ['Внутренние переходы', sourceRows.internal],
    ['Другие входящие', sourceRows.other],
    ['Без источника', sourceRows.undefinedSource]
  ];

  const total = rows.reduce((sum, row) => sum + row[1], 0);

  const values = rows.map(row => {
    const percent = total > 0 ? row[1] / total : 0;
    return [row[0], row[1], percent];
  });

  values.push(['ИТОГО', total, '']);

  EMO_formatLeadShareBlock_(sheet, startRow, values.length);

  sheet.getRange(startRow + 1, 1, values.length, 3).setValues(values);
  sheet.getRange(startRow + 1, 3, values.length - 1, 1).setNumberFormat('0.0%');

  EMO_createLeadShareChart_(sheet, startRow, values.length, total);
  EMO_writeLeadShareConclusion_(sheet, startRow, sourceRows, total);
}


function EMO_writeLeadShareConclusion_(sheet, startRow, sourceRows, total) {
  const conclusionRow = startRow + 12;

  const direct = Number(sourceRows.direct || 0);
  const directShare = total > 0 ? direct / total : 0;

  const text1 = 'По всем источникам за период получено ' + total + ' лидов. Основной объём дал Директ — ' + direct + ' лидов, это ' + EMO_formatPercent_(directShare) + ' от всех обращений.';

  const extraSources = [];
  if (sourceRows.directVisits > 0) extraSources.push('прямых переходов');
  if (sourceRows.yandexSearch > 0) extraSources.push('Яндекс Поиска');
  if (sourceRows.googleSearch > 0) extraSources.push('Google Search');
  if (sourceRows.yandexMaps > 0) extraSources.push('Яндекс Карт');
  if (sourceRows.telegram > 0) extraSources.push('Телеграм');
  if (sourceRows.internal > 0) extraSources.push('внутренних переходов');
  if (sourceRows.other > 0) extraSources.push('других входящих');

  const text2 = extraSources.length > 0
    ? 'Дополнительно лиды пришли из ' + extraSources.join(', ') + '. Органика и прямые переходы дают вспомогательный объём, но ключевым источником заявок остаётся Директ.'
    : 'Дополнительных лидов из других источников за период не было. Ключевым источником заявок остаётся Директ.';

  const inactiveSources = [];
  if (sourceRows.yandexMaps === 0) inactiveSources.push('Яндекс Карты');
  if (sourceRows.telegram === 0) inactiveSources.push('Телеграм');
  if (sourceRows.internal === 0) inactiveSources.push('внутренние переходы');
  if (sourceRows.other === 0) inactiveSources.push('другие входящие');

  const text3 = inactiveSources.length > 0
    ? inactiveSources.join(', ') + ' за период лидов не дали. Их стоит оставить в наблюдении, но основной фокус — на Директе и качестве лидов из него.'
    : 'Все основные источники дали обращения. Основной фокус — на Директе и качестве лидов из него.';

  // Оформляем вывод только A:F. Ниже 244 строки остаются чистыми.
  sheet.getRange(conclusionRow, 1, 4, EMO_CONFIG.COMMENT_COLS)
    .breakApart()
    .clearContent()
    .setBackground('#ffffff')
    .setFontWeight('normal')
    .setFontColor('#000000')
    .setWrap(true)
    .setBorder(false, false, false, false, false, false);

  EMO_writeCommentSection_(sheet, conclusionRow, 'Вывод по источникам лидов');
  EMO_writeCommentText_(sheet, conclusionRow + 1, text1);
  EMO_writeCommentText_(sheet, conclusionRow + 2, text2);
  EMO_writeCommentText_(sheet, conclusionRow + 3, text3);
}

function EMO_getLeadShareFromMetrika_(dateFrom, dateTo) {
  const token = PropertiesService.getScriptProperties().getProperty('METRIKA_TOKEN');
  const counterId = PropertiesService.getScriptProperties().getProperty('METRIKA_COUNTER_ID');

  if (!token) throw new Error('Не найден METRIKA_TOKEN');
  if (!counterId) throw new Error('Не найден METRIKA_COUNTER_ID');

  const metrics = EMO_CONFIG.LEAD_SHARE_GOALS
    .map(id => 'ym:s:goal' + id + 'visits')
    .join(',');

  const dimensions = [
    'ym:s:lastsignTrafficSource',
    'ym:s:lastsignSourceEngine'
  ].join(',');

  const params = {
    ids: counterId,
    date1: dateFrom,
    date2: dateTo,
    metrics: metrics,
    dimensions: dimensions,
    accuracy: 'full',
    limit: 100000,
    lang: 'ru',
    include_undefined: 'true'
  };

  const url = 'https://api-metrika.yandex.net/stat/v1/data?' + EMO_toQueryString_(params);

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

  const result = {
    direct: 0,
    directVisits: 0,
    yandexSearch: 0,
    yandexMaps: 0,
    googleSearch: 0,
    telegram: 0,
    internal: 0,
    other: 0,
    undefinedSource: 0
  };

  (json.data || []).forEach(item => {
    const dimensions = item.dimensions || [];
    const trafficSource = dimensions[0] ? String(dimensions[0].name || '').toLowerCase() : '';
    const sourceEngine = dimensions[1] ? String(dimensions[1].name || '').toLowerCase() : '';

    const leads = (item.metrics || []).reduce((sum, value) => sum + Number(value || 0), 0);
    if (leads === 0) return;

    const bucket = EMO_detectLeadSourceBucket_(trafficSource, sourceEngine);
    result[bucket] += leads;
  });

  return result;
}


function EMO_formatLeadShareBlock_(sheet, startRow, valuesLength) {
  // Чистим всю зону 228–249, чтобы убрать старые хвосты, серые строки и выделения.
  const clearRows = EMO_CONFIG.LEAD_SHARE_CLEAR_ROWS || 22;

  sheet.getRange(startRow, 1, clearRows, EMO_CONFIG.TOTAL_COLS)
    .breakApart()
    .clearContent()
    .setBackground('#ffffff')
    .setFontColor('#000000')
    .setFontWeight('normal')
    .setWrap(false)
    .setBorder(false, false, false, false, false, false);

  const tableRows = valuesLength + 1; // заголовок + строки источников
  const tableRange = sheet.getRange(startRow, 1, tableRows, 3);

  tableRange
    .setBackground('#d9d9d9')
    .setFontColor('#000000')
    .setWrap(true)
    .setBorder(true, true, true, true, true, true);

  sheet.getRange(startRow, 1, 1, 3)
    .setBackground('#d9ead3')
    .setFontWeight('bold');

  sheet.getRange(startRow, 1).setValue('Доля по лидам:');

  sheet.getRange(startRow + 1, 1, valuesLength, 1).setHorizontalAlignment('left');
  sheet.getRange(startRow + 1, 2, valuesLength, 2).setHorizontalAlignment('center');
  sheet.getRange(startRow + valuesLength, 1, 1, 3).setFontWeight('bold');
}

function EMO_createLeadShareChart_(sheet, leadBlockRow, valuesLength, total) {
  const dataStartRow = leadBlockRow + 1;
  const dataRowsCount = valuesLength - 1;

  // Ставим график напротив блока «Доля по лидам», а не выше.
  const chartRow = leadBlockRow;
  const chartCol = 6; // F

  EMO_removeLeadShareVisuals_(sheet, chartRow, leadBlockRow);

  const sourceRange = sheet.getRange(dataStartRow, 1, dataRowsCount, 2);

  const chart = sheet.newChart()
    .asPieChart()
    .addRange(sourceRange)
    .setOption('title', 'Доля лидов по источникам — ' + total + ' лидов')
    .setOption('pieHole', 0.55)
    .setOption('pieSliceText', 'percentage')
    .setOption('legend', {
      position: 'right',
      textStyle: {
        fontSize: 10
      }
    })
    .setOption('titleTextStyle', {
      fontSize: 15,
      bold: true
    })
    .setOption('width', 760)
    .setOption('height', 300)
    .setPosition(chartRow, chartCol, 0, 0)
    .build();

  sheet.insertChart(chart);
}

function EMO_removeLeadShareVisuals_(sheet, chartRow, leadBlockRow) {
  // Удаляем все старые графики/картинки рядом с блоком «Доля по лидам»,
  // чтобы не оставались две диаграммы.
  const minRow = Math.max(1, leadBlockRow - 80);
  const maxRow = leadBlockRow + 120;
  const minCol = 4;

  sheet.getCharts().forEach(chart => {
    const options = chart.getOptions();
    const title = String(options.get('title') || '').toLowerCase();
    const info = chart.getContainerInfo();
    const anchorRow = info ? info.getAnchorRow() : 0;
    const anchorCol = info ? info.getAnchorColumn() : 0;

    const byTitle = title.includes('доля лидов') || title.includes('источникам');
    const nearBlock = anchorRow >= minRow && anchorRow <= maxRow && anchorCol >= minCol;

    if (byTitle || nearBlock) {
      sheet.removeChart(chart);
    }
  });

  try {
    sheet.getImages().forEach(image => {
      const cell = image.getAnchorCell();
      const row = cell.getRow();
      const col = cell.getColumn();

      if (row >= minRow && row <= maxRow && col >= minCol) {
        image.remove();
      }
    });
  } catch (e) {}

  try {
    sheet.getDrawings().forEach(drawing => {
      try {
        const info = drawing.getContainerInfo ? drawing.getContainerInfo() : null;
        const row = info ? info.getAnchorRow() : 0;
        const col = info ? info.getAnchorColumn() : 0;

        if (row >= minRow && row <= maxRow && col >= minCol) {
          drawing.remove();
        }
      } catch (e) {}
    });
  } catch (e) {}
}

function EMO_updateMonthlyDynamicsFromFebruary_(currentDateTo) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let dynSheet = ss.getSheetByName('ЕМО_Динамика');

  if (!dynSheet) {
    dynSheet = ss.insertSheet('ЕМО_Динамика');
  }

  dynSheet.clear();
  dynSheet.getRange(1, 1, 1, 4).setValues([[
    'Месяц',
    'Расход',
    'Лиды',
    'CPA'
  ]]);

  dynSheet.getRange(1, 1, 1, 4)
    .setFontWeight('bold')
    .setBackground('#d9ead3');

  const periods = EMO_getMonthlyPeriodsFromStart_(EMO_CONFIG.DYNAMICS_START_DATE, currentDateTo);
  const values = [];

  periods.forEach(period => {
    try {
      const rows = EMO_getMonthlyCampaignReport_(period.dateFrom, period.dateTo);
      const totals = EMO_calcTotals_(rows);
      const cpa = totals.factConversions > 0 ? totals.cost / totals.factConversions : 0;

      values.push([
        EMO_monthLabel_(period.dateFrom),
        totals.cost,
        totals.factConversions,
        cpa
      ]);
    } catch (e) {
      values.push([
        EMO_monthLabel_(period.dateFrom),
        0,
        0,
        0
      ]);
    }
  });

  if (values.length > 0) {
    dynSheet.getRange(2, 1, values.length, 4).setValues(values);
    dynSheet.getRange(2, 2, values.length, 1).setNumberFormat('# ##0.00 руб.');
    dynSheet.getRange(2, 4, values.length, 1).setNumberFormat('# ##0.00 руб.');
  }

  EMO_createMonthlyDynamicsChart_(dynSheet);
}

function EMO_createMonthlyDynamicsChart_(sheet) {
  sheet.getCharts().forEach(chart => {
    const title = String(chart.getOptions().get('title') || '');
    if (title === 'Динамика по месяцам') sheet.removeChart(chart);
  });

  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return;

  const chart = sheet.newChart()
    .asLineChart()
    .addRange(sheet.getRange(1, 1, lastRow, 4))
    .setOption('title', 'Динамика по месяцам')
    .setOption('legend', {
      position: 'bottom'
    })
    .setOption('width', 760)
    .setOption('height', 360)
    .setPosition(2, 6, 0, 0)
    .build();

  sheet.insertChart(chart);
}


function EMO_extendTotalRowFormulas_(sheet, headerRow, totalRow) {
  // Протягиваем формулы строки ИТОГО до последней строки кампаний.
  // Например C162:C177 станет C162:C185.
  const dataStartRow = headerRow + 1;
  const dataEndRow = totalRow - 1;
  if (dataEndRow < dataStartRow) return;

  const totalRange = sheet.getRange(totalRow, 1, 1, EMO_CONFIG.TOTAL_COLS);
  const formulas = totalRange.getFormulas()[0];

  formulas.forEach((formula, index) => {
    if (!formula) return;

    let updated = formula;

    // Обычные ссылки: C162:C177
    updated = updated.replace(/(\$?[A-Z]{1,3}\$?)\d+:\1\d+/g, function(match, col) {
      return col + dataStartRow + ':' + col + dataEndRow;
    });

    // На всякий случай ссылки с разными $: $C$162:$C$177
    updated = updated.replace(/(\$?[A-Z]{1,3}\$?)\d+:(\$?[A-Z]{1,3}\$?)\d+/g, function(match, col1, col2) {
      const c1 = col1.replace(/\$/g, '');
      const c2 = col2.replace(/\$/g, '');
      if (c1 === c2) {
        return col1 + dataStartRow + ':' + col2 + dataEndRow;
      }
      return match;
    });

    if (updated !== formula) {
      sheet.getRange(totalRow, index + 1).setFormula(updated);
    }
  });

  SpreadsheetApp.flush();
}

// ================== ПОИСК СТРОК ==================

function EMO_findBlockByPeriod_(sheet, dateFrom, dateTo) {
  const title = EMO_buildTitle_(dateFrom, dateTo);
  const lastRow = sheet.getLastRow();
  const values = sheet.getRange(1, 1, lastRow, 1).getDisplayValues();

  for (let i = 0; i < values.length; i++) {
    if (String(values[i][0]).trim() === title) {
      const topRow = i + 1;
      const headerRow = EMO_findHeaderRowAfter_(sheet, topRow);
      const totalRow = EMO_findTotalRowAfter_(sheet, headerRow);

      return {
        topRow,
        headerRow,
        totalRow
      };
    }
  }

  return null;
}

function EMO_findHeaderRowAfter_(sheet, startRow) {
  const lastRow = sheet.getLastRow();

  for (let row = startRow; row <= lastRow; row++) {
    const value = String(sheet.getRange(row, 1).getDisplayValue()).trim();

    if (value === 'РК') {
      return row;
    }
  }

  return 0;
}

function EMO_findTotalRowAfter_(sheet, startRow) {
  const lastRow = sheet.getLastRow();

  for (let row = startRow; row <= lastRow; row++) {
    const value = String(sheet.getRange(row, 1).getDisplayValue()).trim();

    if (value === 'ИТОГО') {
      return row;
    }
  }

  return 0;
}

function EMO_findLastBlockEndRow_(sheet) {
  const lastRow = sheet.getLastRow();
  const lastCol = EMO_CONFIG.TOTAL_COLS;
  const values = sheet.getRange(1, 1, lastRow, lastCol).getDisplayValues();

  for (let r = values.length - 1; r >= 0; r--) {
    const hasContent = values[r].some(v => String(v).trim() !== '');

    if (hasContent) {
      return r + 1;
    }
  }

  return EMO_CONFIG.TEMPLATE_END_ROW;
}

function EMO_getLeadShareRow_(blockTopRow) {
  return blockTopRow + (EMO_CONFIG.LEAD_SHARE_TEMPLATE_ROW - EMO_CONFIG.TEMPLATE_START_ROW);
}

// ================== ПЕРИОДЫ ==================

function EMO_getToday_() {
  return Utilities.formatDate(new Date(), EMO_CONFIG.TIMEZONE, 'yyyy-MM-dd');
}

function EMO_getPreviousMonthPeriod_() {
  const today = EMO_getToday_();
  const parts = today.split('-').map(Number);

  const firstDayCurrentMonth = new Date(parts[0], parts[1] - 1, 1);
  const lastDayPreviousMonth = new Date(firstDayCurrentMonth);
  lastDayPreviousMonth.setDate(0);

  const firstDayPreviousMonth = new Date(
    lastDayPreviousMonth.getFullYear(),
    lastDayPreviousMonth.getMonth(),
    1
  );

  return {
    dateFrom: Utilities.formatDate(firstDayPreviousMonth, EMO_CONFIG.TIMEZONE, 'yyyy-MM-dd'),
    dateTo: Utilities.formatDate(lastDayPreviousMonth, EMO_CONFIG.TIMEZONE, 'yyyy-MM-dd')
  };
}

function EMO_getPreviousMonthForPeriod_(dateFrom) {
  const parts = dateFrom.split('-').map(Number);

  const firstCurrent = new Date(parts[0], parts[1] - 1, 1);
  const lastPrev = new Date(firstCurrent);
  lastPrev.setDate(0);

  const firstPrev = new Date(
    lastPrev.getFullYear(),
    lastPrev.getMonth(),
    1
  );

  return {
    dateFrom: Utilities.formatDate(firstPrev, EMO_CONFIG.TIMEZONE, 'yyyy-MM-dd'),
    dateTo: Utilities.formatDate(lastPrev, EMO_CONFIG.TIMEZONE, 'yyyy-MM-dd')
  };
}

function EMO_getMonthlyPeriodsFromStart_(startDate, currentDateTo) {
  const startParts = startDate.split('-').map(Number);
  const endParts = currentDateTo.split('-').map(Number);

  const start = new Date(startParts[0], startParts[1] - 1, 1);
  const end = new Date(endParts[0], endParts[1] - 1, 1);

  const periods = [];
  let cursor = new Date(start);

  while (cursor <= end) {
    const first = new Date(cursor.getFullYear(), cursor.getMonth(), 1);
    const last = new Date(cursor.getFullYear(), cursor.getMonth() + 1, 0);

    let dateTo = Utilities.formatDate(last, EMO_CONFIG.TIMEZONE, 'yyyy-MM-dd');
    const monthKey = Utilities.formatDate(first, EMO_CONFIG.TIMEZONE, 'yyyy-MM');
    const currentMonthKey = currentDateTo.slice(0, 7);

    if (monthKey === currentMonthKey) {
      dateTo = currentDateTo;
    }

    periods.push({
      dateFrom: Utilities.formatDate(first, EMO_CONFIG.TIMEZONE, 'yyyy-MM-dd'),
      dateTo
    });

    cursor.setMonth(cursor.getMonth() + 1);
  }

  return periods;
}

function EMO_buildTitle_(dateFrom, dateTo) {
  return 'Отчет месячный по РК  ' +
    EMO_formatShortDate_(dateFrom) +
    ' по ' +
    EMO_formatShortDateFull_(dateTo);
}

function EMO_formatShortDate_(apiDate) {
  const parts = apiDate.split('-');
  return Number(parts[2]) + '.' + parts[1] + '.';
}

function EMO_formatShortDateFull_(apiDate) {
  const parts = apiDate.split('-');
  return Number(parts[2]) + '.' + parts[1] + '.' + parts[0];
}

function EMO_monthLabel_(apiDate) {
  const parts = apiDate.split('-');
  return parts[1] + '.' + parts[0];
}

// ================== ВСПОМОГАТЕЛЬНЫЕ ==================

function EMO_getCampaignType_(campaignName) {
  const name = String(campaignName).toLowerCase();

  if (name.includes('поиск/сети') || name.includes('m-k')) {
    return 'Поиск/Сети';
  }

  if (name.includes('сет') || name.includes('_s_') || name.includes('ретаргет')) {
    return 'Сети';
  }

  if (name.includes('карт')) {
    return 'Карты';
  }

  if (name.includes('поиск') || name.includes('_p_') || name.includes('brand')) {
    return 'Поиск';
  }

  return 'н/а';
}

function EMO_rowHasGoals_(row) {
  return (
    row.callMain > 0 ||
    row.phone1 > 0 ||
    row.phone2 > 0 ||
    row.max > 0 ||
    row.telegramSlider > 0 ||
    row.kvalPlusZayavki > 0 ||
    row.jivoOffline > 0 ||
    row.callsUis > 0 ||
    row.zayavkaUis > 0
  );
}

function EMO_valueOrDash_(value) {
  const num = Number(value || 0);
  return num === 0 ? '-' : num;
}

function EMO_detectLeadSourceBucket_(trafficSource, sourceEngine) {
  const sourceText = (trafficSource + ' ' + sourceEngine).toLowerCase();

  if (
    sourceText.includes('direct') ||
    sourceText.includes('директ') ||
    sourceText.includes('yandex direct') ||
    sourceText.includes('яндекс.директ')
  ) {
    return 'direct';
  }

  if (
    sourceText.includes('прям') ||
    sourceText.includes('direct traffic') ||
    sourceText.includes('typed')
  ) {
    return 'directVisits';
  }

  if (
    sourceText.includes('maps') ||
    sourceText.includes('карты') ||
    sourceText.includes('справочник') ||
    sourceText.includes('business')
  ) {
    return 'yandexMaps';
  }

  if (
    sourceText.includes('telegram') ||
    sourceText.includes('t.me') ||
    sourceText.includes('телеграм')
  ) {
    return 'telegram';
  }

  if (sourceText.includes('google')) {
    return 'googleSearch';
  }

  if (
    sourceText.includes('yandex') ||
    sourceText.includes('яндекс')
  ) {
    return 'yandexSearch';
  }

  if (
    sourceText.includes('internal') ||
    sourceText.includes('внутрен')
  ) {
    return 'internal';
  }

  if (!trafficSource && !sourceEngine) {
    return 'undefinedSource';
  }

  if (
    sourceText.includes('не определ') ||
    sourceText.includes('undefined') ||
    sourceText.includes('not set')
  ) {
    return 'undefinedSource';
  }

  return 'other';
}

function EMO_compareMoney_(current, previous) {
  if (!previous || previous === 0) {
    return 'составил ' + EMO_formatRub_(current);
  }

  if (current > previous) {
    return 'вырос с ' + EMO_formatRub_(previous) + ' до ' + EMO_formatRub_(current);
  }

  if (current < previous) {
    return 'снизился с ' + EMO_formatRub_(previous) + ' до ' + EMO_formatRub_(current);
  }

  return 'остался на уровне ' + EMO_formatRub_(current);
}

function EMO_compareNumber_(current, previous) {
  if (!previous || previous === 0) {
    return 'составили ' + EMO_formatNumber_(current);
  }

  if (current > previous) {
    return 'выросли с ' + EMO_formatNumber_(previous) + ' до ' + EMO_formatNumber_(current);
  }

  if (current < previous) {
    return 'снизились с ' + EMO_formatNumber_(previous) + ' до ' + EMO_formatNumber_(current);
  }

  return 'остались на уровне ' + EMO_formatNumber_(current);
}

function EMO_toQueryString_(params) {
  return Object.keys(params)
    .map(key => encodeURIComponent(key) + '=' + encodeURIComponent(params[key]))
    .join('&');
}

function EMO_toNumber_(value) {
  if (
    value === undefined ||
    value === null ||
    value === '' ||
    value === '--' ||
    value === '-'
  ) {
    return 0;
  }

  const normalized = String(value)
    .replace(/\s/g, '')
    .replace(',', '.');

  const num = Number(normalized);

  return isNaN(num) ? 0 : num;
}

function EMO_formatRub_(value) {
  return Number(value || 0).toLocaleString('ru-RU', {
    minimumFractionDigits: 2,
    maximumFractionDigits: 2
  }) + ' руб.';
}

function EMO_formatNumber_(value) {
  return Number(value || 0).toLocaleString('ru-RU');
}

function EMO_formatPercent_(value) {
  return Number(value || 0).toLocaleString('ru-RU', {
    minimumFractionDigits: 2,
    maximumFractionDigits: 2
  }) + '%';
}

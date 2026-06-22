/* 
============================================================
ЕНО — ЧИСТЫЙ ЗАЩИЩЁННЫЙ КОД + ПРАВИЛА
Версия: v20_geo_cpa + rules v1
============================================================

ПРАВИЛА, КОТОРЫЕ НУЖНО ПОМНИТЬ ПРИ ЛЮБЫХ ПРАВКАХ ЕНО:

ЕНО — ПРАВИЛО, ЧТОБЫ БЛОК НЕ ПАДАЛ И НЕ МОЛЧАЛ

Версия: 1.1

Дополнение v13:
- колонка R больше не содержит цель TalkMe онлайн;
- колонка U содержит TalkMe: Клиент написал в чат (онлайн), цель 432303787;
- колонка V содержит Уникальный звонок Alytics, цель 512458206;
- колонка W очищается в новом недельном блоке;
- комментарий по ГЕО считает конверсии как Ecommerce: покупка + TalkMe онлайн + Уникальный звонок Alytics.
- колонка R содержит цель 359915942: Клик: По ссылке в мессенджер.
- лишняя колонка W с названием 'Уникальный звонок' удаляется локально внутри текущего недельного блока;
- колонка 'Уникальный звонок Alytics' остаётся;
- колонка 'Факт сумма конверсий' не удаляется и должна встать сразу после 'Уникальный звонок Alytics'.
- в комментарии по ГЕО добавляется CPA = расход / конверсии.

1. Не использовать временные hotfix-вставки как постоянное решение.
После проверки hotfix нужно переносить в основной чистый код.

2. В коде не должно быть дублей функций.
Особенно:
- tdmeFindTotalRow_
- tdmeFitCampaignRows_
- tdmeFixTotalRowFormulaRanges_

Если функция объявлена два раза, Apps Script использует последнюю, но код становится непредсказуемым.

3. Поиск строки ИТОГО должен быть устойчивым:
- искать ИТОГО по всей строке, а не только в первой колонке;
- если ИТОГО не найдено, искать строку перед комментариями;
- если и это не сработало, определять ИТОГО по последней строке кампаний;
- выдавать понятную ошибку с номером строки.

4. После сборки недельного отчёта обязательно запускать проверку:
- строка ИТОГО ниже последней строки кампаний;
- слово ИТОГО не попало в строки кампаний;
- формулы ИТОГО тянутся до последней строки кампаний;
- формулы не перезаписаны значениями.

5. Автозапуск не должен молчать:
- функция fillTdmEnoPreviousFullWeek должна быть обёрнута в try/catch;
- при ошибке отправлять письмо;
- ошибку всё равно пробрасывать дальше, чтобы она была видна в журнале выполнения.

6. Перед выдачей нового кода:
- отдавать TXT-файлом;
- проверять, что нет дублей функций;
- проверять, что ИТОГО ищется безопасно;
- проверять, что есть post-build validation;
- проверять, что формулы не перезаписываются.


============================================================
НИЖЕ ОСНОВНОЙ КОД ДЛЯ eno.gs
============================================================
*/

// =====================
// ТДМ — ЕНО недельный отчёт
// Клиент: ilhaleontj3v
// Вкладка: ЕНО
// v20 clean protected: R = Клик по ссылке в мессенджер, U = TalkMe онлайн, V = Уникальный звонок Alytics, в ГЕО добавлен CPA,
// формулы не перезаписываем, отчёт валидируем после создания,
// при ошибке отправляем уведомление на почту.
// =====================

const TDME_CLIENT_LOGIN = 'ilhaleontj3v';
const TDME_SHEET_NAME = 'ЕНО';
const TDME_TIMEZONE = 'Europe/Moscow';

const TDME_TEMPLATE_TITLE_PART = '11.05.2026 - 17.05.2026';
const TDME_FIRST_REPORT_TOP_ROW = 527;
const TDME_MONTH_PLAN = 210000;
const TDME_METRIKA_COUNTER_ID = '92370926';

// Почта для уведомлений об ошибках. Если пусто — берём почту активного пользователя.
const TDME_NOTIFY_EMAIL = '';

// Цели как в Мастере отчётов
const TDME_GOALS = {
  view3Pages: ['453453800'],      // Просмотр 3х страниц
  phoneClick: ['432293138'],      // Клик: По номеру телефона
  emailClick: ['388393244'],      // Клик: По email адресу
  messengerClick: ['359915942'],  // Клик: По ссылке в мессенджер
  talkMeOnline: ['432303787'],    // TalkMe: Клиент написал в чат онлайн
  addToCart: ['504318736'],       // Ecommerce: добавление в корзину
  purchase: ['504318735'],        // Ecommerce: покупка
  talkMe: [],                     // Не используем в текущей структуре ЕНО
  email: [],                      // Цель удалённая цель Roistat Email удалена
  calls: ['512458206']            // Уникальный звонок Alytics
};

// Тестовый запуск — пересобрать неделю 25.05–31.05.2026
function fillTdmEnoPeriodTest() {
  tdmeFillReport_('2026-05-25', '2026-05-31', null);
}

// Запуск вручную — отчёт за 18.05–24.05
function fillTdmEno18to24May() {
  tdmeFillReport_('2026-05-18', '2026-05-24', TDME_FIRST_REPORT_TOP_ROW);
}

// Автоматический отчёт за прошлую полную неделю.
// Важно: если отчёт не собрался, отправляем ошибку на почту,
// чтобы ЕНО не "молчал" и проблема была видна сразу.
function fillTdmEnoPreviousFullWeek() {
  try {
    const period = tdmePreviousFullWeek_();
    tdmeFillReport_(period.dateFrom, period.dateTo, null);
  } catch (e) {
    tdmeNotifyError_('ЕНО не обновился', e);
    throw e;
  }
}

// Создать триггер на понедельник 9:00
function createTdmEnoWeeklyTriggerMonday9am() {
  const triggers = ScriptApp.getProjectTriggers();

  triggers.forEach(trigger => {
    if (trigger.getHandlerFunction() === 'fillTdmEnoPreviousFullWeek') {
      ScriptApp.deleteTrigger(trigger);
    }
  });

  ScriptApp.newTrigger('fillTdmEnoPreviousFullWeek')
    .timeBased()
    .onWeekDay(ScriptApp.WeekDay.MONDAY)
    .atHour(9)
    .inTimezone(TDME_TIMEZONE)
    .create();
}

function tdmeFillReport_(dateFrom, dateTo, forcedTopRow) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(TDME_SHEET_NAME);

  if (!sheet) {
    throw new Error('Не найден лист: ' + TDME_SHEET_NAME);
  }

  const rows = tdmeLoadCampaignRows_(dateFrom, dateTo);
  rows.sort((a, b) => b.cost - a.cost);

  const monthRows = tdmeLoadCampaignRows_(tdmeMonthStart_(dateTo), dateTo);

  const title = tdmeTitle_(dateFrom, dateTo);
  const topRow = forcedTopRow || tdmeNextTopRow_(sheet, title);

  // Важно: копируем предыдущий недельный блок, чтобы сохранить формулы/форматы.
  // Если текущий блок уже есть, он пересобирается поверх него.
  tdmeCopyPreviousReportBlock_(sheet, topRow, title);

  let block = tdmeGetBlock_(sheet, topRow);

  // Если из старого шаблона/ручных правок приехала лишняя колонка "Уникальный звонок",
  // удаляем её только внутри текущего недельного блока, не трогая весь лист.
  tdmeRemoveExtraUniqueCallColumn_(sheet, block);

  // После локального сдвига перечитываем блок, чтобы строки/колонки были актуальны.
  block = tdmeGetBlock_(sheet, topRow);
  block = tdmeFitCampaignRows_(sheet, block, rows.length);

  const header = tdmeHeaderMap_(sheet, block.headerRow);
  const dataStartRow = block.headerRow + 1;
  const dataRowsCount = Math.max(rows.length, 1);
  const dataEndRow = dataStartRow + dataRowsCount - 1;

  // Критично: строка ИТОГО должна быть строго сразу после последней строки кампаний.
  // Не доверяем старому поиску по скопированному блоку, потому что при копировании/удалении строк
  // метка ИТОГО может съехать или исчезнуть.
  block.totalRow = tdmeForceTotalRowAfterData_(sheet, header, dataEndRow, block.totalRow);
  tdmeApplyUpdatedGoalColumns_(sheet, block.headerRow, block.totalRow, header);

  tdmeSetPlain_(sheet, block.topRow, 1, title);

  // Чистим только заполняемые колонки строк кампаний.
  // Это убирает случайные формулы ИТОГО из строк кампаний,
  // но не трогает расчётные колонки CTR / CPC / CPA / CR и т.д.
  tdmeClearDataInputCells_(sheet, header, dataStartRow, dataRowsCount);
  tdmeFillRows_(sheet, header, dataStartRow, rows);

  // Строку ИТОГО значениями НЕ перезаписываем.
  // Только фиксируем метку и диапазоны формул под фактический конец текущей недели.
  tdmeEnsureTotalLabel_(sheet, header, block.totalRow);
  tdmeFixTotalRowFormulaRanges_(sheet, block.totalRow, dataStartRow, dataEndRow);
  tdmeFixTopSummaryFormulas_(sheet, block.topRow, block.totalRow);

  tdmeFormatBounceColumn_(sheet, header, dataStartRow, dataRowsCount, block.totalRow);
  tdmeHighlightGreenRows_(sheet, header, dataStartRow, dataRowsCount, rows);

  const comments = tdmeBuildComments_(rows, monthRows, dateFrom, dateTo);
  const commentRow = tdmeFindCommentRow_(sheet, block.totalRow) || block.totalRow + 3;
  tdmeWriteComments_(sheet, commentRow, comments);

  // Контрольная проверка, чтобы отчёт не считался успешным, если блок собрался криво.
  tdmeValidateReportBlock_(sheet, block, dataStartRow, dataEndRow, title);
}

// =====================
// ДИРЕКТ
// =====================

function tdmeLoadCampaignRows_(dateFrom, dateTo) {
  const goalIds = tdmeAllGoalIds_();

  const body = {
    params: {
      SelectionCriteria: {
        DateFrom: dateFrom,
        DateTo: dateTo
      },
      Goals: goalIds.map(id => Number(id)),
      AttributionModels: ['AUTO'],
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
      ReportName: 'tdm_eno_' + dateFrom + '_' + dateTo + '_' + Utilities.getUuid(),
      ReportType: 'CAMPAIGN_PERFORMANCE_REPORT',
      DateRangeType: 'CUSTOM_DATE',
      Format: 'TSV',
      IncludeVAT: 'YES',
      IncludeDiscount: 'NO'
    }
  };

  const text = tdmeDirectRequest_(body);
  const rows = tdmeParseCampaignTsv_(text);

  tdmeApplyMetrikaAutomaticGoals_(rows, dateFrom, dateTo);

  return rows;
}

function tdmeDirectRequest_(body) {
  let token = PropertiesService.getScriptProperties().getProperty('YANDEX_TOKEN');

  if (!token) {
    throw new Error('Не найден YANDEX_TOKEN в Script Properties.');
  }

  token = String(token)
    .replace(/^OAuth\s+/i, '')
    .replace(/^Bearer\s+/i, '')
    .trim();

  const url = 'https://api.direct.yandex.com/json/v5/reports';

  const options = {
    method: 'post',
    contentType: 'application/json',
    payload: JSON.stringify(body),
    headers: {
      Authorization: 'Bearer ' + token,
      'Client-Login': TDME_CLIENT_LOGIN,
      'Accept-Language': 'ru',
      processingMode: 'auto',
      returnMoneyInMicros: 'false',
      skipReportHeader: 'true',
      skipColumnHeader: 'false',
      skipReportSummary: 'true'
    },
    muteHttpExceptions: true
  };

  for (let attempt = 1; attempt <= 15; attempt++) {
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

    throw new Error('Ошибка Директа. Код: ' + code + '\n' + text);
  }

  throw new Error('Отчёт Директа не успел сформироваться.');
}

function tdmeApplyMetrikaAutomaticGoals_(rows, dateFrom, dateTo) {
  const metrika = tdmeFetchMetrikaAutomaticGoalsByCampaign_(dateFrom, dateTo);

  rows.forEach(row => {
    const key = tdmeCampaignKey_(row.campaignName);
    const goals = metrika.byCampaign[key] || tdmeEmptyMetrikaGoals_();

    // Эти три цели берём из Метрики по automaticTrafficSource, как в ДБ.
    row.goals.purchase = goals.purchase;
    row.goals.talkMeOnline = goals.talkMeOnline;
    row.goals.calls = goals.calls;
  });

  if (tdmeHasMetrikaGoals_(metrika.noCampaign)) {
    rows.push(tdmeBuildMetrikaNoCampaignRow_(metrika.noCampaign));
  }
}

function tdmeFetchMetrikaAutomaticGoalsByCampaign_(dateFrom, dateTo) {
  let token =
    PropertiesService.getScriptProperties().getProperty('YANDEX_METRIKA_TOKEN') ||
    PropertiesService.getScriptProperties().getProperty('YANDEX_TOKEN');

  if (!token) {
    throw new Error('Не найден YANDEX_METRIKA_TOKEN или YANDEX_TOKEN в Script Properties.');
  }

  token = String(token)
    .replace(/^OAuth\s+/i, '')
    .replace(/^Bearer\s+/i, '')
    .trim();

  const params = {
    ids: TDME_METRIKA_COUNTER_ID,
    date1: dateFrom,
    date2: dateTo,
    dimensions: 'ym:s:automaticTrafficSource,ym:s:lastsignUTMCampaign',
    metrics: [
      'ym:s:goal504318735reaches',
      'ym:s:goal432303787reaches',
      'ym:s:goal512458206reaches'
    ].join(','),
    accuracy: 'full',
    limit: 10000
  };

  const query = Object.keys(params)
    .map(key => encodeURIComponent(key) + '=' + encodeURIComponent(params[key]))
    .join('&');

  const response = UrlFetchApp.fetch('https://api-metrika.yandex.net/stat/v1/data?' + query, {
    method: 'get',
    headers: {
      Authorization: 'OAuth ' + token
    },
    muteHttpExceptions: true
  });

  const code = response.getResponseCode();
  const text = response.getContentText();

  if (code < 200 || code >= 300) {
    throw new Error('Ошибка Метрики. Код: ' + code + '. Ответ: ' + text);
  }

  const json = JSON.parse(text);
  const result = {
    byCampaign: {},
    noCampaign: tdmeEmptyMetrikaGoals_()
  };

  (json.data || []).forEach(item => {
    const source = item.dimensions[0] || {};

    if (!tdmeIsMetrikaAdTraffic_(source)) {
      return;
    }

    const campaignName = item.dimensions[1] && item.dimensions[1].name;
    const goals = {
      purchase: tdmeToNumber_(item.metrics[0]),
      talkMeOnline: tdmeToNumber_(item.metrics[1]),
      calls: tdmeToNumber_(item.metrics[2])
    };

    if (!campaignName) {
      tdmeAddMetrikaGoals_(result.noCampaign, goals);
      return;
    }

    const key = tdmeCampaignKey_(campaignName);

    if (!result.byCampaign[key]) {
      result.byCampaign[key] = tdmeEmptyMetrikaGoals_();
    }

    tdmeAddMetrikaGoals_(result.byCampaign[key], goals);
  });

  return result;
}

function tdmeBuildMetrikaNoCampaignRow_(goals) {
  return {
    campaignName: 'Без источника. Метрика',
    campaignType: 'н/а',
    geo: 'РФ',
    impressions: 0,
    clicks: 0,
    cost: 0,
    avgBid: 0,
    avgPosition: '',
    avgTraffic: '',
    bounce: 0,
    depth: 0,
    goals: {
      view3Pages: 0,
      phoneClick: 0,
      emailClick: 0,
      messengerClick: 0,
      talkMeOnline: goals.talkMeOnline,
      addToCart: 0,
      purchase: goals.purchase,
      talkMe: 0,
      email: 0,
      calls: goals.calls
    }
  };
}

function tdmeEmptyMetrikaGoals_() {
  return {
    purchase: 0,
    talkMeOnline: 0,
    calls: 0
  };
}

function tdmeAddMetrikaGoals_(target, source) {
  target.purchase += Number(source.purchase || 0);
  target.talkMeOnline += Number(source.talkMeOnline || 0);
  target.calls += Number(source.calls || 0);
}

function tdmeHasMetrikaGoals_(goals) {
  return Number(goals.purchase || 0) +
    Number(goals.talkMeOnline || 0) +
    Number(goals.calls || 0) > 0;
}

function tdmeIsMetrikaAdTraffic_(source) {
  const id = tdmeNormalize_(source && source.id);
  const name = tdmeNormalize_(source && source.name);

  return id === 'ad' || name === 'ad traffic' || name === 'переходы по рекламе';
}

function tdmeCampaignKey_(value) {
  let text = tdmeNormalize_(value);

  if (text.indexOf('|') !== -1) {
    const parts = text.split('|').map(part => part.trim()).filter(Boolean);
    text = parts[parts.length - 1] || text;
  }

  text = text.replace(/-\d+$/g, '');
  text = text.replace(/[^a-zа-я0-9]+/g, '_');
  text = text.replace(/_+/g, '_').replace(/^_+|_+$/g, '');

  return text;
}

function tdmeParseCampaignTsv_(tsvText) {
  const text = String(tsvText || '').trim();

  if (!text) return [];

  const lines = text.split(/\r?\n/).filter(line => String(line).trim() !== '');

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

  const idxCampaign = tdmeRequireColumn_(headers, 'CampaignName');
  const idxImpressions = tdmeRequireColumn_(headers, 'Impressions');
  const idxClicks = tdmeRequireColumn_(headers, 'Clicks');
  const idxCost = tdmeRequireColumn_(headers, 'Cost');

  const idxAvgBid = headers.indexOf('AvgEffectiveBid');
  const idxAvgPosition = headers.indexOf('AvgImpressionPosition');
  const idxAvgTraffic = headers.indexOf('AvgTrafficVolume');
  const idxBounce = headers.indexOf('BounceRate');
  const idxDepth = headers.indexOf('AvgPageviews');

  const goalIndexes = {};

  tdmeAllGoalIds_().forEach(goalId => {
    goalIndexes[goalId] = tdmeFindGoalColumn_(headers, goalId);

    if (goalIndexes[goalId] === -1) {
      Logger.log('Не найден столбец цели в отчёте Директа: ' + goalId);
    }
  });

  const result = [];

  for (let i = headerIndex + 1; i < lines.length; i++) {
    const row = lines[i].split('\t');
    const campaignName = String(row[idxCampaign] || '').trim();

    if (!campaignName) continue;

    result.push({
      campaignName: campaignName,
      campaignType: tdmeCampaignType_(campaignName),
      geo: tdmeGeo_(campaignName),
      impressions: tdmeToNumber_(row[idxImpressions]),
      clicks: tdmeToNumber_(row[idxClicks]),
      cost: tdmeToNumber_(row[idxCost]),

      avgBid: idxAvgBid === -1 ? 0 : tdmeToNumber_(row[idxAvgBid]),
      avgPosition: idxAvgPosition === -1 ? '' : tdmeMetric_(row[idxAvgPosition]),
      avgTraffic: idxAvgTraffic === -1 ? '' : tdmeMetric_(row[idxAvgTraffic]),
      bounce: idxBounce === -1 ? 0 : tdmeToNumber_(row[idxBounce]) / 100,
      depth: idxDepth === -1 ? 0 : tdmeToNumber_(row[idxDepth]),

      goals: {
        view3Pages: tdmeGoalSum_(row, goalIndexes, TDME_GOALS.view3Pages),
        phoneClick: tdmeGoalSum_(row, goalIndexes, TDME_GOALS.phoneClick),
        emailClick: tdmeGoalSum_(row, goalIndexes, TDME_GOALS.emailClick),
        messengerClick: tdmeGoalSum_(row, goalIndexes, TDME_GOALS.messengerClick),
        talkMeOnline: tdmeGoalSum_(row, goalIndexes, TDME_GOALS.talkMeOnline),
        addToCart: tdmeGoalSum_(row, goalIndexes, TDME_GOALS.addToCart),
        purchase: tdmeGoalSum_(row, goalIndexes, TDME_GOALS.purchase),
        talkMe: tdmeGoalSum_(row, goalIndexes, TDME_GOALS.talkMe),
        email: tdmeGoalSum_(row, goalIndexes, TDME_GOALS.email),
        calls: tdmeGoalSum_(row, goalIndexes, TDME_GOALS.calls)
      }
    });
  }

  return result;
}

function tdmeAllGoalIds_() {
  let result = [];

  Object.keys(TDME_GOALS).forEach(key => {
    result = result.concat(TDME_GOALS[key] || []);
  });

  return Array.from(new Set(result.map(id => String(id))));
}

function tdmeFindGoalColumn_(headers, goalId) {
  const id = String(goalId);

  let index = headers.findIndex(header => {
    return String(header).trim() === 'Conversions_' + id + '_AUTO';
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

function tdmeGoalSum_(row, goalIndexes, goalIds) {
  return goalIds.reduce((sum, goalId) => {
    const idx = goalIndexes[goalId];

    if (idx === -1 || idx === undefined) return sum;

    return sum + tdmeToNumber_(row[idx]);
  }, 0);
}

// =====================
// ШАБЛОН / БЛОК
// =====================

function tdmeCopyPreviousReportBlock_(sheet, topRow, targetTitle) {
  const sourceTopRow = tdmeFindPreviousReportTopRow_(sheet, topRow) || tdmeFindTemplateTopRow_(sheet);
  const copyRows = tdmeReportRowsCount_(sheet, sourceTopRow);
  const lastCol = sheet.getLastColumn();

  if (sheet.getMaxRows() < topRow + copyRows + 20) {
    sheet.insertRowsAfter(
      sheet.getMaxRows(),
      topRow + copyRows + 20 - sheet.getMaxRows()
    );
  }

  sheet
    .getRange(sourceTopRow, 1, copyRows, lastCol)
    .copyTo(
      sheet.getRange(topRow, 1, copyRows, lastCol),
      SpreadsheetApp.CopyPasteType.PASTE_NORMAL,
      false
    );

  sheet.getRange(topRow, 1).setValue(targetTitle);
}

function tdmeFindPreviousReportTopRow_(sheet, currentTopRow) {
  const titles = tdmeFindAllReportTopRows_(sheet)
    .filter(row => row < currentTopRow)
    .sort((a, b) => b - a);

  return titles.length ? titles[0] : 0;
}

function tdmeFindAllReportTopRows_(sheet) {
  const lastRow = sheet.getLastRow();
  const values = sheet.getRange(1, 1, lastRow, 1).getDisplayValues();
  const result = [];

  for (let i = 0; i < values.length; i++) {
    const text = String(values[i][0] || '').trim();

    if (tdmeNormalize_(text).indexOf('отчет недельный по рк') !== -1) {
      result.push(i + 1);
    }
  }

  return result;
}

function tdmeReportRowsCount_(sheet, topRow) {
  const reportRows = tdmeFindAllReportTopRows_(sheet).sort((a, b) => a - b);
  const nextTop = reportRows.find(row => row > topRow);

  if (nextTop) {
    return Math.max(40, nextTop - topRow - 1);
  }

  // Если это последний отчёт, берём блок до конца комментария, но без огромного хвоста.
  const headerRow = tdmeFindHeaderRow_(sheet, topRow);
  const totalRow = tdmeFindTotalRow_(sheet, headerRow);
  const commentRow = tdmeFindCommentRow_(sheet, totalRow) || totalRow + 3;
  const endRow = Math.min(sheet.getLastRow(), commentRow + 55);

  return Math.max(60, endRow - topRow + 1);
}

function tdmeFindTemplateTopRow_(sheet) {
  const lastRow = sheet.getLastRow();
  const values = sheet.getRange(1, 1, lastRow, 1).getDisplayValues();

  for (let i = 0; i < values.length; i++) {
    const text = String(values[i][0] || '');

    if (
      tdmeNormalize_(text).indexOf('отчет недельный по рк') !== -1 &&
      text.indexOf(TDME_TEMPLATE_TITLE_PART) !== -1
    ) {
      return i + 1;
    }
  }

  throw new Error('Не нашла шаблонный отчёт: ' + TDME_TEMPLATE_TITLE_PART);
}

function tdmeNextTopRow_(sheet, title) {
  const existing = tdmeFindTitleRow_(sheet, title);

  if (existing) {
    return existing;
  }

  const firstCell = String(sheet.getRange(TDME_FIRST_REPORT_TOP_ROW, 1).getDisplayValue() || '').trim();

  if (!firstCell) {
    return TDME_FIRST_REPORT_TOP_ROW;
  }

  return tdmeLastFilledRowInColumnA_(sheet) + 2;
}

function tdmeFindTitleRow_(sheet, title) {
  const lastRow = sheet.getLastRow();
  const values = sheet.getRange(1, 1, lastRow, 1).getDisplayValues();

  for (let i = 0; i < values.length; i++) {
    if (String(values[i][0]).trim() === title) {
      return i + 1;
    }
  }

  return 0;
}

function tdmeGetBlock_(sheet, topRow) {
  const headerRow = tdmeFindHeaderRow_(sheet, topRow);
  const totalRow = tdmeFindTotalRow_(sheet, headerRow);

  return {
    topRow: topRow,
    headerRow: headerRow,
    totalRow: totalRow
  };
}

function tdmeFindHeaderRow_(sheet, topRow) {
  const maxRow = Math.min(sheet.getLastRow(), topRow + 40);
  const maxCol = sheet.getLastColumn();

  for (let row = topRow; row <= maxRow; row++) {
    const values = sheet.getRange(row, 1, 1, maxCol).getDisplayValues()[0];
    const normalized = values.map(value => tdmeNormalize_(value));

    const hasCampaign = normalized.some(value => value === 'рк');
    const hasImpressions = normalized.some(value => value.indexOf('показы') !== -1);

    if (hasCampaign && hasImpressions) {
      return row;
    }
  }

  throw new Error('Не нашла шапку таблицы ЕНО.');
}


// Строка ИТОГО должна быть строго после последней строки кампаний.
// Если после операций со строками метка съехала, возвращаем строку на место.
function tdmeForceTotalRowAfterData_(sheet, header, dataEndRow, detectedTotalRow) {
  const expectedTotalRow = dataEndRow + 1;

  if (detectedTotalRow !== expectedTotalRow) {
    Logger.log(
      'ИТОГО было найдено на строке ' + detectedTotalRow +
      ', но должно быть на строке ' + expectedTotalRow +
      '. Используем строку после последней кампании.'
    );
  }

  tdmeEnsureTotalLabel_(sheet, header, expectedTotalRow);

  return expectedTotalRow;
}

function tdmeRemoveExtraUniqueCallColumn_(sheet, block) {
  const headerRow = block.headerRow;
  const totalRow = block.totalRow;
  const lastCol = sheet.getLastColumn();
  const headers = sheet.getRange(headerRow, 1, 1, lastCol).getDisplayValues()[0];

  // Ищем именно лишний столбец "Уникальный звонок".
  // Столбец "Уникальный звонок Alytics" НЕ трогаем.
  const colsToDelete = [];

  headers.forEach((header, index) => {
    const text = tdmeNormalize_(header);

    if (text === 'уникальный звонок') {
      colsToDelete.push(index + 1);
    }
  });

  if (!colsToDelete.length) return;

  // Удаляем справа налево, чтобы индексы не съезжали.
  colsToDelete
    .sort((a, b) => b - a)
    .forEach(col => {
      sheet
        .getRange(block.topRow, col, totalRow - block.topRow + 1, 1)
        .deleteCells(SpreadsheetApp.Dimension.COLUMNS);
    });
}

function tdmeFitCampaignRows_(sheet, block, rowsCount) {
  const dataStartRow = block.headerRow + 1;
  const neededRows = Math.max(rowsCount, 1);
  let existingRows = block.totalRow - dataStartRow;
  const lastCol = sheet.getLastColumn();

  if (existingRows < neededRows) {
    const rowsToAdd = neededRows - existingRows;

    sheet.insertRowsBefore(block.totalRow, rowsToAdd);

    // Копируем именно первую строку кампании: там должны быть формулы расчётных колонок и формат.
    sheet
      .getRange(dataStartRow, 1, 1, lastCol)
      .copyTo(
        sheet.getRange(block.totalRow, 1, rowsToAdd, lastCol),
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

  return {
    topRow: block.topRow,
    headerRow: block.headerRow,
    totalRow: block.totalRow
  };
}

function tdmeLastFilledRowInColumnA_(sheet) {
  const lastRow = sheet.getLastRow();
  const values = sheet.getRange(1, 1, lastRow, 1).getDisplayValues();

  for (let i = values.length - 1; i >= 0; i--) {
    if (String(values[i][0]).trim() !== '') {
      return i + 1;
    }
  }

  return TDME_FIRST_REPORT_TOP_ROW;
}

// =====================
// КОЛОНКИ / ЗАПОЛНЕНИЕ
// =====================

function tdmeHeaderMap_(sheet, headerRow) {
  const headers = sheet.getRange(headerRow, 1, 1, sheet.getLastColumn()).getDisplayValues()[0];
  const map = {};

  headers.forEach((header, index) => {
    const key = tdmeNormalize_(header);

    if (key) {
      map[key] = index + 1;
    }
  });

  // Текущая структура ЕНО:
  // R — Клик: По ссылке в мессенджер
  // U — TalkMe: Клиент написал в чат (онлайн)
  // V — Уникальный звонок Alytics
  const fixedMessengerCol = sheet.getLastColumn() >= 18
    ? 18
    : tdmeContainsCol_(map, ['клик: по ссылке в мессенджер', 'клик по ссылке в мессенджер']);

  const fixedTalkMeOnlineCol = sheet.getLastColumn() >= 21
    ? 21
    : tdmeContainsCol_(map, ['talkme: клиент написал в чат', 'talkme клиент написал в чат']);

  const fixedUniqueCallCol = sheet.getLastColumn() >= 22
    ? 22
    : tdmeRoistatCallsCol_(map);

  return {
    rk: tdmeExactCol_(map, 'рк'),
    type: tdmeExactCol_(map, 'тип рк'),

    impressions: tdmeContainsCol_(map, ['показы']),
    clicks: tdmeContainsCol_(map, ['клики']),
    cost: tdmeContainsCol_(map, ['расход c ндс', 'расход с ндс']),

    avgBid: tdmeContainsCol_(map, ['ср. ставка за клик']),
    avgPosition: tdmeContainsCol_(map, ['средняя позиция показа']),
    avgTraffic: tdmeContainsCol_(map, ['средний объем трафика', 'средний объём трафика']),
    bounce: tdmeContainsCol_(map, ['отказы']),
    depth: tdmeContainsCol_(map, ['глубина']),

    view3Pages: tdmeContainsCol_(map, ['просмотр 3х страниц']),
    phoneClick: tdmeContainsCol_(map, ['клик: по номеру телефона', 'клик по номеру телефона']),
    emailClick: tdmeContainsCol_(map, ['клик: по email адресу', 'клик по email адресу']),
    messengerClick: fixedMessengerCol,
    addToCart: tdmeContainsCol_(map, ['ecommerce: добавление в корзину', 'ecommerce добавление в корзину']),
    purchase: tdmeContainsCol_(map, ['ecommerce: покупка', 'ecommerce покупка']),
    talkMeOnline: fixedTalkMeOnlineCol,

    talkMe: 0,
    email: 0,
    calls: fixedUniqueCallCol
  };
}

function tdmeExactCol_(map, header) {
  return map[tdmeNormalize_(header)] || 0;
}

function tdmeContainsCol_(map, parts) {
  const normalizedParts = parts.map(part => tdmeNormalize_(part));

  for (const key in map) {
    for (let i = 0; i < normalizedParts.length; i++) {
      if (key.indexOf(normalizedParts[i]) !== -1) {
        return map[key];
      }
    }
  }

  return 0;
}


function tdmeRoistatEmailCol_(map) {
  return map['roistat email'] || map['email'] || 0;
}

function tdmeRoistatCallsCol_(map) {
  return map['roistat звонки'] || map['звонки'] || 0;
}

function tdmeApplyUpdatedGoalColumns_(sheet, headerRow, totalRow, header) {
  const lastCol = sheet.getLastColumn();

  // R: новая цель "Клик: По ссылке в мессенджер".
  if (lastCol >= 18) {
    sheet.getRange(headerRow, 18).setValue('Клик: По ссылке в мессенджер');
  }

  // U: TalkMe онлайн.
  if (lastCol >= 21) {
    sheet.getRange(headerRow, 21).setValue('TalkMe: Клиент написал в чат (онлайн)');
  }

  // V: Уникальный звонок Alytics.
  if (lastCol >= 22) {
    sheet.getRange(headerRow, 22).setValue('Уникальный звонок Alytics');
  }

  // W не очищаем и не удаляем:
  // там остаётся "Факт сумма конверсий" и формулы/значения отчёта.
}

function tdmeFixTopSummaryFormulas_(sheet, topRow, totalRow) {
  const lastCol = sheet.getLastColumn();
  const values = sheet.getRange(topRow, 1, 1, lastCol).getDisplayValues()[0];
  const formulasByLabel = {
    'расход': '=G' + totalRow,
    'cpc': '=F' + totalRow,
    'клики': '=D' + totalRow,
    'cpa': '=O' + totalRow,
    'кол-во лид': '=W' + totalRow,
    'cpa лид': '=Y' + totalRow
  };

  Object.keys(formulasByLabel).forEach(label => {
    const labelCol = tdmeFindTopSummaryLabelCol_(values, label);

    if (!labelCol || labelCol >= lastCol) return;

    sheet.getRange(topRow, labelCol + 1).setFormula(formulasByLabel[label]);
  });

  tdmeFixTopLeadSummaryBlock_(sheet, topRow, totalRow);
}

function tdmeFindTopSummaryLabelCol_(values, label) {
  const target = tdmeNormalize_(label);

  for (let i = 0; i < values.length; i++) {
    if (tdmeNormalize_(values[i]) === target) {
      return i + 1;
    }
  }

  return 0;
}

function tdmeFixTopLeadSummaryBlock_(sheet, topRow, totalRow) {
  const lastCol = sheet.getLastColumn();

  if (lastCol < 27) return;

  const oldNoteValues = sheet.getRange(topRow, 26, 1, Math.min(2, lastCol - 25))
    .getDisplayValues()[0];
  const note = oldNoteValues
    .map(value => String(value || '').trim())
    .filter(value => value)[0] || '';

  // Верхняя зелёная строка должна идти так:
  // Кол-во Лид -> W, CR -> X, CPA Лид -> Y, комментарий -> AA.
  // Так мы не смешиваем новый столбец W "Факт сумма конверсий" с CR/CPA и заметкой.
  sheet.getRange(topRow, 20, 1, 8).setValues([[
    'Кол-во Лид',
    '=W' + totalRow,
    'CR',
    '=X' + totalRow,
    'CPA Лид',
    '=Y' + totalRow,
    '',
    note
  ]]);
}

function tdmeInputColumns_(header) {
  return [
    header.rk,
    header.type,
    header.impressions,
    header.clicks,
    header.cost,
    header.avgBid,
    header.avgPosition,
    header.avgTraffic,
    header.bounce,
    header.depth,
    header.view3Pages,
    header.phoneClick,
    header.emailClick,
    header.messengerClick,
    header.talkMeOnline,
    header.addToCart,
    header.purchase,
    header.talkMe,
    header.email,
    header.calls
  ].filter(col => col > 0);
}

function tdmeClearDataInputCells_(sheet, header, startRow, rowsCount) {
  const columns = tdmeInputColumns_(header);

  columns.forEach(col => {
    sheet.getRange(startRow, col, rowsCount, 1).clearContent();
  });
}

function tdmeFillRows_(sheet, header, startRow, rows) {
  rows.forEach((item, index) => {
    const row = startRow + index;

    tdmeSetData_(sheet, row, header.rk, item.campaignName);
    tdmeSetData_(sheet, row, header.type, item.campaignType);

    tdmeSetData_(sheet, row, header.impressions, item.impressions);
    tdmeSetData_(sheet, row, header.clicks, item.clicks);
    tdmeSetData_(sheet, row, header.cost, item.cost);
    tdmeSetData_(sheet, row, header.avgBid, item.avgBid);
    tdmeSetData_(sheet, row, header.avgPosition, item.avgPosition);
    tdmeSetData_(sheet, row, header.avgTraffic, item.avgTraffic);
    tdmeSetData_(sheet, row, header.bounce, item.bounce);
    tdmeSetData_(sheet, row, header.depth, item.depth);

    tdmeSetData_(sheet, row, header.view3Pages, tdmeGoalValue_(item.goals.view3Pages));
    tdmeSetData_(sheet, row, header.phoneClick, tdmeGoalValue_(item.goals.phoneClick));
    tdmeSetData_(sheet, row, header.emailClick, tdmeGoalValue_(item.goals.emailClick));
    tdmeSetData_(sheet, row, header.messengerClick, tdmeGoalValue_(item.goals.messengerClick));
    tdmeSetData_(sheet, row, header.talkMeOnline, tdmeGoalValue_(item.goals.talkMeOnline));
    tdmeSetData_(sheet, row, header.addToCart, tdmeGoalValue_(item.goals.addToCart));
    tdmeSetData_(sheet, row, header.purchase, tdmeGoalValue_(item.goals.purchase));
    tdmeSetData_(sheet, row, header.talkMe, tdmeGoalValue_(item.goals.talkMe));
    tdmeSetData_(sheet, row, header.email, tdmeGoalValue_(item.goals.email));
    tdmeSetData_(sheet, row, header.calls, tdmeGoalValue_(item.goals.calls));
  });
}

function tdmeSetData_(sheet, row, col, value) {
  if (col <= 0) return;
  sheet.getRange(row, col).setValue(value);
}

function tdmeSetPlain_(sheet, row, col, value) {
  if (col <= 0) return;
  sheet.getRange(row, col).setValue(value);
}

function tdmeEnsureTotalLabel_(sheet, header, totalRow) {
  if (header.rk > 0) {
    sheet.getRange(totalRow, header.rk).setValue('ИТОГО');
  }

  if (header.type > 0) {
    sheet.getRange(totalRow, header.type).setValue('Итого');
  }
}

function tdmeGoalValue_(value) {
  const number = Number(value || 0);
  return number > 0 ? number : '-';
}

function tdmeFormatBounceColumn_(sheet, header, startRow, rowsCount, totalRow) {
  if (!header.bounce) return;

  sheet.getRange(startRow, header.bounce, rowsCount, 1).setNumberFormat('0.00%');
  sheet.getRange(totalRow, header.bounce).setNumberFormat('0.00%');
}

function tdmeHighlightGreenRows_(sheet, header, startRow, rowsCount, rows) {
  const lastCol = Math.max(
    header.calls,
    header.email,
    header.talkMe,
    header.purchase,
    header.addToCart,
    header.talkMeOnline,
    header.emailClick,
    header.messengerClick,
    header.phoneClick,
    header.view3Pages,
    header.depth,
    header.cost,
    1
  );

  sheet.getRange(startRow, 1, rowsCount, lastCol)
    .setBackground('#ffffff')
    .setFontWeight('normal');

  rows.forEach((item, index) => {
    const targetActions =
      Number(item.goals.phoneClick || 0) +
      Number(item.goals.emailClick || 0) +
      Number(item.goals.messengerClick || 0) +
      Number(item.goals.talkMeOnline || 0) +
      Number(item.goals.addToCart || 0) +
      Number(item.goals.purchase || 0) +
      Number(item.goals.calls || 0);

    if (targetActions > 0) {
      sheet.getRange(startRow + index, 1, 1, lastCol)
        .setBackground('#d9ead3')
        .setFontWeight('normal');
    }
  });
}

// =====================
// КОММЕНТАРИИ
// =====================

function tdmeBuildComments_(weekRows, monthRows, dateFrom, dateTo) {
  const total = tdmeTotals_(weekRows);
  const monthTotal = tdmeTotals_(monthRows);

  const planToDate = tdmePlanToDate_(TDME_MONTH_PLAN, dateTo);
  const budgetPercent = planToDate > 0 ? monthTotal.cost / planToDate * 100 : 0;

  const byType = tdmeGroupByType_(weekRows);
  const byGeo = tdmeGroupByGeo_(weekRows);

  const rows = [];

  rows.push(['Комментарии к отчёту за прошлую неделю с ' + tdmeShortDate_(dateFrom) + ' по ' + tdmeDateRu_(dateTo), '']);
  rows.push(['', '']);

  rows.push(['Бюджет:', '']);
  rows.push(['Всего потрачено за неделю ' + tdmeRub_(total.cost) + '. Идём по бюджету на ' + tdmePercent_(budgetPercent) + ' от плана на дату', '']);
  rows.push(['', '']);

  rows.push(['Трафик:', '']);
  rows.push(['Показы — ' + tdmeInt_(total.impressions), '']);
  rows.push(['Клики — ' + tdmeInt_(total.clicks), '']);
  rows.push(['Средний CPC — ' + tdmeRub_(total.clicks > 0 ? total.cost / total.clicks : 0), '']);
  rows.push(['', '']);

  rows.push(['Достигнуты цели:', '']);
  rows.push(['Просмотр 3х страниц — ' + total.goals.view3Pages, '']);
  rows.push(['Клик по номеру телефона — ' + total.goals.phoneClick, '']);
  rows.push(['Клик по email — ' + total.goals.emailClick, '']);
  rows.push(['Клик по ссылке в мессенджер — ' + total.goals.messengerClick, '']);
  rows.push(['TalkMe онлайн — ' + total.goals.talkMeOnline, '']);
  rows.push(['Ecommerce: добавление в корзину — ' + total.goals.addToCart, '']);
  rows.push(['Ecommerce: покупка — ' + total.goals.purchase, '']);
  rows.push(['Уникальный звонок Alytics — ' + total.goals.calls, '']);
  rows.push(['', '']);

  rows.push(['По ГЕО:', '']);
  rows.push(['', '']);

  ['Казахстан', 'Санкт-Петербург', 'РФ'].forEach(geo => {
    const item = byGeo[geo];

    if (!item) return;

    const cpc = item.clicks > 0 ? item.cost / item.clicks : 0;

    rows.push([geo, '']);
    rows.push([
      'Бюджет — ' +
      tdmeRub_(item.cost) +
      ', показы — ' +
      tdmeInt_(item.impressions) +
      ', клики — ' +
      tdmeInt_(item.clicks) +
      ', средний CPC — ' +
      tdmeRub_(cpc) +
      ', конверсии — ' +
      tdmeInt_(tdmeTargetConversions_(item)) +
      ', CPA — ' +
      tdmeGeoCpa_(item),
      ''
    ]);
    rows.push(['', '']);
  });

  rows.push(['По типам РК:', '']);

  ['Поиск', 'Сети', 'Товарная', 'н/а'].forEach(type => {
    const item = byType[type];

    if (!item) return;

    const ctr = item.impressions > 0 ? item.clicks / item.impressions * 100 : 0;
    const cpc = item.clicks > 0 ? item.cost / item.clicks : 0;

    rows.push([
      type + ' — ' +
      tdmeInt_(item.clicks) +
      ' кликов, расход ' +
      tdmeRub_(item.cost) +
      ', CTR — ' +
      tdmePercent_(ctr) +
      ', CPC — ' +
      tdmeRub_(cpc) +
      ', Уникальный звонок Alytics — ' +
      item.goals.calls +
      ', TalkMe онлайн — ' +
      item.goals.talkMeOnline +
      ', мессенджер — ' +
      item.goals.messengerClick,
      ''
    ]);
  });

  rows.push(['', '']);
  rows.push(['Вывод:', '']);
  rows.push(['Зелёным выделены кампании, которые принесли целевые действия. Кампании с расходом без целевых действий проверяем по запросам, площадкам, объявлениям и посадочным страницам.', '']);

  return rows;
}

function tdmeFindCommentRow_(sheet, totalRow) {
  const maxRow = Math.min(sheet.getLastRow(), totalRow + 80);

  for (let row = totalRow + 1; row <= maxRow; row++) {
    const text = tdmeNormalize_(sheet.getRange(row, 1).getDisplayValue());

    if (text.indexOf('комментар') !== -1) {
      return row;
    }
  }

  return 0;
}

function tdmeWriteComments_(sheet, startRow, rows) {
  sheet.getRange(startRow, 1, 80, 8)
    .clearContent()
    .setBackground('#ffffff')
    .setFontWeight('normal');

  sheet.getRange(startRow, 1, rows.length, 2)
    .setValues(rows)
    .setWrap(true);

  const boldTitles = [
    'Бюджет:',
    'Трафик:',
    'Достигнуты цели:',
    'По ГЕО:',
    'Казахстан',
    'Санкт-Петербург',
    'РФ',
    'По типам РК:',
    'Вывод:'
  ];

  rows.forEach((row, index) => {
    const text = String(row[0] || '').trim();

    if (index === 0 || boldTitles.indexOf(text) !== -1) {
      sheet.getRange(startRow + index, 1)
        .setFontWeight('bold')
        .setFontSize(index === 0 ? 11 : 10);
    }
  });
}

// =====================
// АГРЕГАЦИЯ
// =====================

function tdmeTotals_(rows) {
  const total = tdmeEmptyTotal_();

  rows.forEach(row => {
    tdmeAddToTotal_(total, row);
  });

  return total;
}

function tdmeGroupByType_(rows) {
  const result = {};

  rows.forEach(row => {
    const type = row.campaignType || 'н/а';

    if (!result[type]) {
      result[type] = tdmeEmptyTotal_();
    }

    tdmeAddToTotal_(result[type], row);
  });

  return result;
}

function tdmeGroupByGeo_(rows) {
  const result = {};

  rows.forEach(row => {
    const geo = row.geo || 'РФ';

    if (!result[geo]) {
      result[geo] = tdmeEmptyTotal_();
    }

    tdmeAddToTotal_(result[geo], row);
  });

  return result;
}

function tdmeGeo_(campaignName) {
  const name = String(campaignName || '').toLowerCase();

  if (
    name.indexOf('kaz') !== -1 ||
    name.indexOf('_kz') !== -1 ||
    name.indexOf('kz') !== -1 ||
    name.indexOf('каз') !== -1
  ) {
    return 'Казахстан';
  }

  if (
    name.indexOf('spb') !== -1 ||
    name.indexOf('_spb') !== -1 ||
    name.indexOf('спб') !== -1 ||
    name.indexOf('санкт') !== -1
  ) {
    return 'Санкт-Петербург';
  }

  return 'РФ';
}

function tdmeAddToTotal_(target, row) {
  target.impressions += row.impressions;
  target.clicks += row.clicks;
  target.cost += row.cost;

  target.goals.view3Pages += row.goals.view3Pages;
  target.goals.phoneClick += row.goals.phoneClick;
  target.goals.emailClick += row.goals.emailClick;
  target.goals.messengerClick += row.goals.messengerClick;
  target.goals.talkMeOnline += row.goals.talkMeOnline;
  target.goals.addToCart += row.goals.addToCart;
  target.goals.purchase += row.goals.purchase;
  target.goals.talkMe += row.goals.talkMe;
  target.goals.email += row.goals.email;
  target.goals.calls += row.goals.calls;
}

function tdmeEmptyTotal_() {
  return {
    impressions: 0,
    clicks: 0,
    cost: 0,
    goals: {
      view3Pages: 0,
      phoneClick: 0,
      emailClick: 0,
      messengerClick: 0,
      talkMeOnline: 0,
      addToCart: 0,
      purchase: 0,
      talkMe: 0,
      email: 0,
      calls: 0
    }
  };
}

function tdmeTargetConversions_(item) {
  if (!item || !item.goals) return 0;

  // Конверсии для комментариев по ГЕО:
  // Ecommerce: покупка + TalkMe онлайн + Уникальный звонок Alytics.
  return Number(item.goals.purchase || 0) +
    Number(item.goals.talkMeOnline || 0) +
    Number(item.goals.calls || 0);
}

function tdmeGeoCpa_(item) {
  const conversions = tdmeTargetConversions_(item);

  if (!conversions) {
    return '-';
  }

  return tdmeRub_(Number(item.cost || 0) / conversions);
}


// =====================
// ТИП РК
// =====================

function tdmeCampaignType_(campaignName) {
  const name = String(campaignName || '').toLowerCase();

  if (
    name.indexOf('_s_') !== -1 ||
    name.indexOf('rsya') !== -1 ||
    name.indexOf('рся') !== -1 ||
    name.indexOf('set') !== -1 ||
    name.indexOf('smart') !== -1
  ) {
    return 'Сети';
  }

  if (
    name.indexOf('tovar') !== -1 ||
    name.indexOf('товар') !== -1 ||
    name.indexOf('tk') !== -1
  ) {
    return 'Товарная';
  }

  if (
    name.indexOf('_p_') !== -1 ||
    name.indexOf('search') !== -1 ||
    name.indexOf('poisk') !== -1 ||
    name.indexOf('поиск') !== -1 ||
    name.indexOf('categ_t-g') !== -1 ||
    name.indexOf('dsa') !== -1
  ) {
    return 'Поиск';
  }

  return 'н/а';
}

// =====================
// ДАТЫ / ФОРМАТЫ
// =====================

function tdmePreviousFullWeek_() {
  const now = new Date();
  const today = Utilities.formatDate(now, TDME_TIMEZONE, 'yyyy-MM-dd');
  const todayDate = tdmeParseDate_(today);

  const day = todayDate.getDay();
  const diffToMonday = day === 0 ? 6 : day - 1;

  const currentMonday = new Date(todayDate);
  currentMonday.setDate(todayDate.getDate() - diffToMonday);

  const previousMonday = new Date(currentMonday);
  previousMonday.setDate(currentMonday.getDate() - 7);

  const previousSunday = new Date(currentMonday);
  previousSunday.setDate(currentMonday.getDate() - 1);

  return {
    dateFrom: Utilities.formatDate(previousMonday, TDME_TIMEZONE, 'yyyy-MM-dd'),
    dateTo: Utilities.formatDate(previousSunday, TDME_TIMEZONE, 'yyyy-MM-dd')
  };
}

function tdmeTitle_(dateFrom, dateTo) {
  return 'Отчет недельный по РК  ' + tdmeDateRu_(dateFrom) + ' - ' + tdmeDateRu_(dateTo);
}

function tdmeMonthStart_(apiDate) {
  const date = tdmeParseDate_(apiDate);

  return Utilities.formatDate(
    new Date(date.getFullYear(), date.getMonth(), 1),
    TDME_TIMEZONE,
    'yyyy-MM-dd'
  );
}

function tdmePlanToDate_(monthPlan, dateTo) {
  const date = tdmeParseDate_(dateTo);
  const day = date.getDate();
  const daysInMonth = new Date(date.getFullYear(), date.getMonth() + 1, 0).getDate();

  return monthPlan / daysInMonth * day;
}

function tdmeParseDate_(apiDate) {
  const parts = String(apiDate).split('-').map(Number);
  return new Date(parts[0], parts[1] - 1, parts[2]);
}

function tdmeShortDate_(apiDate) {
  const parts = String(apiDate).split('-');
  return parts[2] + '.' + parts[1] + '.';
}

function tdmeDateRu_(apiDate) {
  const parts = String(apiDate).split('-');
  return parts[2] + '.' + parts[1] + '.' + parts[0];
}

function tdmeMetric_(value) {
  const text = String(value || '').trim();

  if (!text || text === '-' || text === '--') return '-';

  return tdmeToNumber_(text);
}

function tdmeToNumber_(value) {
  if (value === null || value === undefined) return 0;

  const text = String(value)
    .replace(/\s/g, '')
    .replace(',', '.')
    .replace(/[^\d.-]/g, '');

  const number = Number(text);

  return isNaN(number) ? 0 : number;
}

function tdmeRub_(value) {
  return Number(value || 0).toLocaleString('ru-RU', {
    minimumFractionDigits: 2,
    maximumFractionDigits: 2
  }) + ' руб.';
}

function tdmeInt_(value) {
  return Number(value || 0).toLocaleString('ru-RU');
}

function tdmePercent_(value) {
  return Number(value || 0).toLocaleString('ru-RU', {
    minimumFractionDigits: 2,
    maximumFractionDigits: 2
  }) + '%';
}

function tdmeNormalize_(value) {
  return String(value || '')
    .toLowerCase()
    .replace(/ё/g, 'е')
    .replace(/[«»"']/g, '')
    .replace(/\s+/g, ' ')
    .trim();
}

function tdmeRequireColumn_(headers, columnName) {
  const index = headers.indexOf(columnName);

  if (index === -1) {
    throw new Error('Не найден столбец "' + columnName + '". Заголовки: ' + headers.join(' | '));
  }

  return index;
}


// =====================
// ВАЛИДАЦИЯ ЕНО ПОСЛЕ СБОРКИ
// =====================

function tdmeValidateReportBlock_(sheet, block, dataStartRow, dataEndRow, title) {
  const lastCol = sheet.getLastColumn();

  if (!block || !block.headerRow || !block.totalRow) {
    throw new Error('Проверка ЕНО: не найден блок отчёта для ' + title);
  }

  if (block.totalRow <= dataEndRow) {
    throw new Error(
      'Проверка ЕНО: строка ИТОГО стоит раньше конца данных. ' +
      'ИТОГО: ' + block.totalRow + ', последняя строка кампаний: ' + dataEndRow
    );
  }

  const dataValues = sheet
    .getRange(dataStartRow, 1, Math.max(dataEndRow - dataStartRow + 1, 1), lastCol)
    .getDisplayValues();

  for (let r = 0; r < dataValues.length; r++) {
    const hasTotalInData = dataValues[r].some(value => tdmeNormalize_(value) === 'итого');

    if (hasTotalInData) {
      throw new Error('Проверка ЕНО: слово ИТОГО попало в строку кампании ' + (dataStartRow + r));
    }
  }

  const totalValues = sheet.getRange(block.totalRow, 1, 1, lastCol).getDisplayValues()[0];
  const hasTotal = totalValues.some(value => tdmeNormalize_(value) === 'итого');

  if (!hasTotal) {
    // Самовосстановление: если после копирования метка исчезла, возвращаем её в строку ИТОГО.
    const header = tdmeHeaderMap_(sheet, block.headerRow);
    tdmeEnsureTotalLabel_(sheet, header, block.totalRow);

    const recheckValues = sheet.getRange(block.totalRow, 1, 1, lastCol).getDisplayValues()[0];
    const recheckHasTotal = recheckValues.some(value => tdmeNormalize_(value) === 'итого');

    if (!recheckHasTotal) {
      throw new Error('Проверка ЕНО: в строке ' + block.totalRow + ' нет метки ИТОГО.');
    }
  }

  const formulas = sheet.getRange(block.totalRow, 1, 1, lastCol).getFormulas()[0];

  formulas.forEach((formula, index) => {
    if (!formula) return;

    const ranges = String(formula).match(/(\$?[A-Z]{1,3})(\$?\d+):(\$?[A-Z]{1,3})(\$?\d+)/g) || [];

    ranges.forEach(rangeText => {
      const match = rangeText.match(/(\$?[A-Z]{1,3})(\$?\d+):(\$?[A-Z]{1,3})(\$?\d+)/);
      if (!match) return;

      const col1 = match[1].replace(/\$/g, '');
      const col2 = match[3].replace(/\$/g, '');
      const row1 = Number(match[2].replace(/\$/g, ''));
      const row2 = Number(match[4].replace(/\$/g, ''));

      if (col1 === col2 && row1 === dataStartRow && row2 !== dataEndRow) {
        throw new Error(
          'Проверка ЕНО: формула в строке ИТОГО, колонка ' + (index + 1) +
          ' тянется до ' + row2 + ', а должна до ' + dataEndRow +
          '. Формула: ' + formula
        );
      }
    });
  });

  Logger.log('Проверка ЕНО пройдена: ' + title);
}

function tdmeNotifyError_(title, error) {
  const message =
    title + '\n\n' +
    'Ошибка: ' + (error && error.message ? error.message : error) + '\n\n' +
    'Stack:\n' + (error && error.stack ? error.stack : '');

  Logger.log(message);

  try {
    const email = TDME_NOTIFY_EMAIL || Session.getActiveUser().getEmail();

    if (email) {
      MailApp.sendEmail({
        to: email,
        subject: 'Ошибка ЕНО — ТДМ',
        body: message
      });
    }
  } catch (mailError) {
    Logger.log('Не удалось отправить письмо об ошибке ЕНО: ' + mailError.message);
  }
}

// =====================
// ТЕСТ ДОСТУПА К ДИРЕКТУ
// =====================

function testTdmDirectManageAccess() {
  const clientLogin = TDME_CLIENT_LOGIN;

  let token = PropertiesService.getScriptProperties().getProperty('YANDEX_TOKEN');

  if (!token) {
    throw new Error('Не найден YANDEX_TOKEN в Script Properties.');
  }

  token = String(token)
    .replace(/^OAuth\s+/i, '')
    .replace(/^Bearer\s+/i, '')
    .trim();

  const url = 'https://api.direct.yandex.com/json/v5/campaigns';

  const body = {
    method: 'get',
    params: {
      SelectionCriteria: {},
      FieldNames: [
        'Id',
        'Name',
        'State',
        'Status'
      ]
    }
  };

  const options = {
    method: 'post',
    contentType: 'application/json',
    payload: JSON.stringify(body),
    headers: {
      Authorization: 'Bearer ' + token,
      'Client-Login': clientLogin,
      'Accept-Language': 'ru'
    },
    muteHttpExceptions: true
  };

  const response = UrlFetchApp.fetch(url, options);
  const code = response.getResponseCode();
  const text = response.getContentText();

  Logger.log('Код ответа: ' + code);
  Logger.log(text);

  if (code !== 200) {
    throw new Error('Тест доступа не прошёл. Код: ' + code + '\n' + text);
  }

  const data = JSON.parse(text);
  const campaigns = data.result && data.result.Campaigns ? data.result.Campaigns : [];

  Logger.log('Кампаний получено: ' + campaigns.length);

  SpreadsheetApp.getUi().alert(
    'Тест прошёл. Токен видит кампании через API. Кампаний найдено: ' + campaigns.length
  );
}


// =====================
// ИТОГО — безопасное исправление диапазонов формул
// =====================

/**
 * Исправляет диапазоны формул в строке ИТОГО под фактические строки кампаний.
 * Пример:
 * =ПРОМЕЖУТОЧНЫЕ.ИТОГИ(9;C594:C611)
 * станет:
 * =ПРОМЕЖУТОЧНЫЕ.ИТОГИ(9;C594:C614)
 */
function tdmeFixTotalRowFormulaRanges_(sheet, totalRow, dataStartRow, dataEndRow) {
  const lastCol = sheet.getLastColumn();
  const range = sheet.getRange(totalRow, 1, 1, lastCol);
  const formulas = range.getFormulas()[0];

  const fixed = formulas.map(formula => {
    if (!formula) return formula;
    return tdmeFixFormulaRangesToCurrentBlock_(formula, dataStartRow, dataEndRow);
  });

  range.setFormulas([fixed]);
}

/**
 * Меняет только диапазоны в одной колонке:
 * C594:C611 -> C594:C614.
 * Формулы вида G615/D615 не трогаем.
 */
function tdmeFixFormulaRangesToCurrentBlock_(formula, dataStartRow, dataEndRow) {
  return String(formula).replace(
    /(\$?[A-Z]{1,3})(\$?\d+):(\$?[A-Z]{1,3})(\$?\d+)/g,
    function(match, col1, row1, col2, row2) {
      const cleanCol1 = col1.replace(/\$/g, '');
      const cleanCol2 = col2.replace(/\$/g, '');

      // Не трогаем диапазоны через несколько колонок.
      if (cleanCol1 !== cleanCol2) return match;

      const newRow1 = row1.indexOf('$') === 0 ? '$' + dataStartRow : String(dataStartRow);
      const newRow2 = row2.indexOf('$') === 0 ? '$' + dataEndRow : String(dataEndRow);

      return col1 + newRow1 + ':' + col2 + newRow2;
    }
  );
}
// =====================
// ЕНО — безопасный поиск строки ИТОГО
// Встроено в основной код v20
// =====================
//
// Что исправляет:
// - ищет строку ИТОГО не только по первой колонке, а по всей строке;
// - если слово "ИТОГО" не найдено, ищет строку перед блоком комментариев;
// - добавляет подробную диагностику, если строка всё равно не найдена;
// - формулы и значения не перезаписывает.

function tdmeFindTotalRow_(sheet, headerRow) {
  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();
  const maxRow = Math.min(lastRow, headerRow + 180);

  // 1. Основной способ: ищем "ИТОГО" в любой ячейке строки.
  for (let row = headerRow + 1; row <= maxRow; row++) {
    const values = sheet.getRange(row, 1, 1, lastCol).getDisplayValues()[0];

    const hasTotal = values.some(value => {
      return tdmeNormalizeForTotalSearch_(value) === 'итого';
    });

    if (hasTotal) {
      return row;
    }
  }

  // 2. Запасной способ: ищем блок комментариев.
  // В ЕНО строка ИТОГО обычно находится перед комментариями.
  const commentRow = tdmeFindEnoCommentRowAfterHeader_(sheet, headerRow, maxRow, lastCol);

  if (commentRow) {
    for (let row = commentRow - 1; row > headerRow; row--) {
      const values = sheet.getRange(row, 1, 1, lastCol).getDisplayValues()[0];
      const nonEmptyCount = values.filter(value => String(value || '').trim() !== '').length;

      // Берём последнюю непустую строку перед комментариями.
      // Обычно это и есть ИТОГО.
      if (nonEmptyCount >= 2) {
        const firstCell = sheet.getRange(row, 1).getDisplayValue();

        // Если в первой ячейке почему-то нет "ИТОГО", аккуратно ставим.
        // Формулы не трогаем.
        if (tdmeNormalizeForTotalSearch_(firstCell) !== 'итого') {
          const cell = sheet.getRange(row, 1);

          if (!cell.getFormula()) {
            cell.setValue('ИТОГО');
          }
        }

        return row;
      }
    }
  }

  // 3. Ещё один запасной способ:
  // ищем последнюю строку с кампаниями после шапки, а следующую строку считаем ИТОГО.
  const calculatedTotalRow = tdmeGuessTotalRowByCampaignRows_(sheet, headerRow, maxRow, lastCol);

  if (calculatedTotalRow) {
    const cell = sheet.getRange(calculatedTotalRow, 1);

    if (!cell.getFormula() && tdmeNormalizeForTotalSearch_(cell.getDisplayValue()) !== 'итого') {
      cell.setValue('ИТОГО');
    }

    return calculatedTotalRow;
  }

  throw new Error(
    'Не нашла строку ИТОГО после строки ' + headerRow +
    '. Проверь, что в предыдущем недельном блоке есть строка ИТОГО и блок комментариев. ' +
    'Диапазон поиска: строки ' + (headerRow + 1) + '–' + maxRow + '.'
  );
}

function tdmeFindEnoCommentRowAfterHeader_(sheet, headerRow, maxRow, lastCol) {
  for (let row = headerRow + 1; row <= maxRow; row++) {
    const values = sheet.getRange(row, 1, 1, lastCol).getDisplayValues()[0];

    const hasComment = values.some(value => {
      const text = tdmeNormalizeForTotalSearch_(value);
      return text.indexOf('комментарии к отчету') !== -1 ||
        text.indexOf('комментарии к отчёту') !== -1 ||
        text.indexOf('комментарии') !== -1;
    });

    if (hasComment) {
      return row;
    }
  }

  return 0;
}

function tdmeGuessTotalRowByCampaignRows_(sheet, headerRow, maxRow, lastCol) {
  let lastDataRow = 0;

  for (let row = headerRow + 1; row <= maxRow; row++) {
    const values = sheet.getRange(row, 1, 1, lastCol).getDisplayValues()[0];
    const firstCell = tdmeNormalizeForTotalSearch_(values[0]);
    const nonEmptyCount = values.filter(value => String(value || '').trim() !== '').length;

    // Останавливаемся на комментариях / новом отчёте.
    const rowText = tdmeNormalizeForTotalSearch_(values.join(' '));

    if (
      rowText.indexOf('комментарии') !== -1 ||
      rowText.indexOf('отчет недельный') !== -1 ||
      rowText.indexOf('отчёт недельный') !== -1
    ) {
      break;
    }

    // Строки кампаний обычно содержат название РК и числовые данные.
    if (firstCell && firstCell !== 'итого' && nonEmptyCount >= 4) {
      lastDataRow = row;
    }

    // Если нашли пустую строку после данных — следующая логика уже не нужна.
    if (!firstCell && nonEmptyCount === 0 && lastDataRow) {
      break;
    }
  }

  if (!lastDataRow) {
    return 0;
  }

  return lastDataRow + 1;
}

function tdmeNormalizeForTotalSearch_(value) {
  return String(value || '')
    .toLowerCase()
    .replace(/ё/g, 'е')
    .replace(/[«»"']/g, '')
    .replace(/\u00a0/g, ' ')
    .replace(/\s+/g, ' ')
    .trim();
}

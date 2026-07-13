/*
============================================================
ЕНО — ЧИСТЫЙ ЗАЩИЩЁННЫЙ КОД + ПРАВИЛА
Версия: v20_geo_cpa + rules v1
============================================================

ПРАВИЛА, КОТОРЫЕ НУЖНО ПОМНИТЬ ПРИ ЛЮБЫХ ПРАВКАХ ЕНО:

ЕНО — ПРАВИЛО, ЧТОБЫ БЛОК НЕ ПАДАЛ И НЕ МОЛЧАЛ

Версия: 1.1

Актуальная схема v21:
- старые внешние цели не используются в активной логике;
- отчёт заполняется по Метрике AUTO, Ecommerce, Jivo и Callibri после подключения;
- лишние legacy-колонки из старых шаблонов не должны попадать в расчёт фактических лидов;
- колонка 'Факт сумма конверсий' сохраняется;
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
// v21 clean protected: цели отчёта = Метрика AUTO + Ecommerce + Jivo + Callibri после подключения.
// Старые внешние трекеры не используются в активной логике.
// Формулы не перезаписываем, отчёт валидируем после создания,
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
  view3Pages: ['453453800'],      // Просмотр 3х страниц — микроцель
  addToCart: ['504318736'],       // Ecommerce: добавление в корзину
  purchase: ['504318735'],        // Ecommerce: покупка — факт
  jivo: ['575188424'],            // Jivo-сайт: пользователь начал чат — факт
  callibriSpam: [],               // Callibri: Спам — источник пока не подключён
  callibriNonTarget: [],          // Callibri: Нецелевой_Лид — источник пока не подключён
  callibriLeadA: [],              // Callibri: Лид_Квал_A — источник пока не подключён
  callibriLeadC: []               // Callibri: Лид_Квал_C — источник пока не подключён
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

  // Сначала безопасно получаем Директ и Метрику. Callibri здесь не вызывается.
  const rows = tdmeLoadCampaignRows_(dateFrom, dateTo);
  const monthRows = tdmeLoadCampaignRows_(tdmeMonthStart_(dateTo), dateTo);

  // Один неизменяемый снимок Callibri за неделю используем во всей цепочке:
  // ДБ -> распределение по кампаниям ЕНО -> итоги -> сверка.
  const callibriSnapshot = tdmGptCallibriAggregateYandexCpcByCampaign_(dateFrom, dateTo);
  tdmApplyCallibriAggregateToDbPeriod_(callibriSnapshot, dateFrom, dateTo);
  tdmeApplyCallibriGoals_(rows, dateFrom, dateTo, callibriSnapshot);
  rows.sort((a, b) => b.cost - a.cost);

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

/**
 * REF-ENO-20260708
 * Безопасные утилиты для пакетной работы и API.
 */
function tdmeSafeLog_(context, error) {
  const message =
    '[ENO] ' + context + ': ' +
    (error && error.message ? error.message : String(error || 'unknown error'));

  Logger.log(message);

  if (error && error.stack) {
    Logger.log(error.stack);
  }
}

function tdmeNormalizeToken_(token) {
  return String(token || '')
    .replace(/^OAuth\s+/i, '')
    .replace(/^Bearer\s+/i, '')
    .trim();
}

function tdmeColumnLetter_(column) {
  let temp = Number(column || 0);
  let letter = '';

  while (temp > 0) {
    const modulo = (temp - 1) % 26;
    letter = String.fromCharCode(65 + modulo) + letter;
    temp = Math.floor((temp - modulo) / 26);
  }

  return letter;
}

function tdmeSafeJsonParse_(text, context) {
  try {
    return JSON.parse(String(text || '{}'));
  } catch (error) {
    throw new Error(
      context + ': не удалось распарсить JSON. Ответ: ' +
      String(text || '').slice(0, 1000)
    );
  }
}

function tdmeFetchWithRetry_(url, options, config) {
  const maxAttempts = Number(config && config.maxAttempts || 5);
  const sleepMs = Number(config && config.sleepMs || 3000);
  const retryCodes = (config && config.retryCodes) || [201, 202, 429, 500, 502, 503, 504];

  let lastError = null;

  for (let attempt = 1; attempt <= maxAttempts; attempt++) {
    try {
      const response = UrlFetchApp.fetch(url, options);
      const code = response.getResponseCode();
      const text = response.getContentText();

      if (code >= 200 && code < 300 && code !== 201 && code !== 202) {
        return {
          code: code,
          text: text,
          response: response
        };
      }

      if (retryCodes.indexOf(code) !== -1 && attempt < maxAttempts) {
        Logger.log(
          '[ENO] API retry ' + attempt + '/' + maxAttempts +
          '. Code=' + code +
          '. Body=' + String(text || '').slice(0, 500)
        );
        Utilities.sleep(sleepMs);
        continue;
      }

      throw new Error('HTTP ' + code + '. Ответ: ' + String(text || '').slice(0, 2000));
    } catch (error) {
      lastError = error;

      if (attempt < maxAttempts) {
        Logger.log(
          '[ENO] API exception retry ' + attempt + '/' + maxAttempts +
          ': ' + (error && error.message ? error.message : String(error))
        );
        Utilities.sleep(sleepMs);
        continue;
      }
    }
  }

  throw lastError || new Error('API request failed without response');
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
        'CampaignId',
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

function tdmeApplyCallibriGoals_(rows, dateFrom, dateTo, byDateCampaign) {
  if (!byDateCampaign || typeof byDateCampaign !== 'object') {
    throw new Error('ЕНО: снимок Callibri не передан. Повторный запрос API внутри распределения запрещён.');
  }

  const byCampaign = tdmCallibriRollupByCampaign_(byDateCampaign);
  const directById = {};
  const directByName = {};

  rows.forEach(function(row, index) {
    row.goals = row.goals || {};
    row.goals.callibriSpam = 0;
    row.goals.callibriNonTarget = 0;
    row.goals.callibriLeadA = 0;
    row.goals.callibriLeadB = 0;
    row.goals.callibriLeadC = 0;

    const campaignId = String(row.campaignId || '').trim();
    const campaignNameKey = tdmGptCallibriCampaignKey_(row.campaignName);

    if (campaignId) {
      if (!directById[campaignId]) directById[campaignId] = [];
      directById[campaignId].push(index);
    }
    if (!directByName[campaignNameKey]) directByName[campaignNameKey] = [];
    directByName[campaignNameKey].push(index);
  });

  Object.keys(byCampaign).forEach(function(key) {
    const callibri = byCampaign[key] || tdmCallibriEmptyReportRow_();
    const total = Number(callibri.callibriSpam || 0) +
      Number(callibri.callibriNonTarget || 0) +
      Number(callibri.callibriLeadA || 0) +
      Number(callibri.callibriLeadB || 0) +
      Number(callibri.callibriLeadC || 0);
    if (!total) return;

    const campaignId = String(callibri.campaignId || '').trim();
    const idMatches = campaignId && directById[campaignId] ? directById[campaignId] : [];
    const nameMatches = directByName[key] || [];
    const matches = idMatches.length ? idMatches : nameMatches;

    if (matches.length > 1) {
      Logger.log('[ENO] Callibri: найдено несколько строк Директа для ключа ' + key + '. Агрегат будет записан один раз в строку ' + (matches[0] + 1) + '.');
    }

    if (matches.length) {
      tdmeAddCallibriToCampaignRow_(rows[matches[0]], callibri);
      return;
    }

    // Добавляем строку только для реально нераспознанного utm_campaign из API.
    // Служебные строки "сверка", "дельта" и "не распределено" не создаются.
    const rawCampaign = String(callibri.utmCampaign || '').trim();
    const unmatchedName = rawCampaign || campaignId || 'Callibri без utm_campaign';
    const unmatchedRow = tdmeBuildUnmatchedCallibriRow_(unmatchedName, campaignId);
    tdmeAddCallibriToCampaignRow_(unmatchedRow, callibri);
    rows.push(unmatchedRow);
  });

  const apiTotalsForReconcile = tdmeCallibriTotalsFromRollup_(byCampaign);
  const enoTotals = tdmeCallibriTotalsFromRows_(rows);
  tdmeAssertCallibriTotalsEqual_('ЕНО: распределение Callibri по кампаниям', apiTotalsForReconcile, enoTotals);
  tdmeReconcileCallibriWithDb_(rows, dateFrom, dateTo, apiTotalsForReconcile);
}

function tdmeAddCallibriToCampaignRow_(row, callibri) {
  row.goals = row.goals || {};
  row.goals.callibriSpam = Number(row.goals.callibriSpam || 0) + Number(callibri.callibriSpam || 0);
  row.goals.callibriNonTarget = Number(row.goals.callibriNonTarget || 0) + Number(callibri.callibriNonTarget || 0);
  row.goals.callibriLeadA = Number(row.goals.callibriLeadA || 0) + Number(callibri.callibriLeadA || 0);
  row.goals.callibriLeadB = Number(row.goals.callibriLeadB || 0) + Number(callibri.callibriLeadB || 0);
  row.goals.callibriLeadC = Number(row.goals.callibriLeadC || 0) + Number(callibri.callibriLeadC || 0);
}

function tdmeBuildUnmatchedCallibriRow_(campaignName, campaignId) {
  return {
    campaignId: campaignId || '',
    campaignName: campaignName,
    campaignType: 'Callibri: не распознано в Директе',
    geo: 'н/а',
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
      callibriSpam: 0,
      callibriNonTarget: 0,
      addToCart: 0,
      purchase: 0,
      jivo: 0,
      callibriLeadA: 0,
      callibriLeadB: 0,
      callibriLeadC: 0
    }
  };
}

function tdmeCallibriTotalsFromRollup_(byCampaign) {
  const totals = { spam: 0, nonTarget: 0, leadA: 0, leadB: 0, leadC: 0 };
  Object.keys(byCampaign || {}).forEach(function(key) {
    const item = byCampaign[key] || {};
    totals.spam += Number(item.callibriSpam || 0);
    totals.nonTarget += Number(item.callibriNonTarget || 0);
    totals.leadA += Number(item.callibriLeadA || 0);
    totals.leadB += Number(item.callibriLeadB || 0);
    totals.leadC += Number(item.callibriLeadC || 0);
  });
  return totals;
}

function tdmeCallibriTotalsFromRows_(rows) {
  const totals = { spam: 0, nonTarget: 0, leadA: 0, leadB: 0, leadC: 0 };
  (rows || []).forEach(function(row) {
    const goals = row.goals || {};
    totals.spam += Number(goals.callibriSpam || 0);
    totals.nonTarget += Number(goals.callibriNonTarget || 0);
    totals.leadA += Number(goals.callibriLeadA || 0);
    totals.leadB += Number(goals.callibriLeadB || 0);
    totals.leadC += Number(goals.callibriLeadC || 0);
  });
  return totals;
}

function tdmeAssertCallibriTotalsEqual_(context, expected, actual) {
  const mismatches = [];
  ['spam', 'nonTarget', 'leadA', 'leadB', 'leadC'].forEach(function(key) {
    if (Number(expected[key] || 0) !== Number(actual[key] || 0)) {
      mismatches.push(key + ': expected=' + Number(expected[key] || 0) + ', actual=' + Number(actual[key] || 0));
    }
  });
  if (mismatches.length) throw new Error(context + ': ' + mismatches.join('; '));
}

function tdmeReconcileCallibriWithDb_(rows, dateFrom, dateTo, providedApiTotals) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const db = ss.getSheetByName('ДБ');
  if (!db) throw new Error('ЕНО: не найден лист ДБ для сверки Callibri.');

  const dbTotals = { spam: 0, nonTarget: 0, leadA: 0, leadB: 0, leadC: 0 };
  const apiTotals = providedApiTotals || { spam: 0, nonTarget: 0, leadA: 0, leadB: 0, leadC: 0 };
  const dates = tdmCallibriApiDates_(dateFrom, dateTo);
  let hasLeadBColumn = false;

  dates.forEach(function(apiDate) {
    const headerRow = TDM_DB.findHeaderRowForMonth_(db, apiDate);
    const header = TDM_DB.getHeaderMap_(db, headerRow);
    const sheetRow = TDM_DB.findDateRow_(db, header.date, headerRow, apiDate);
    if (!sheetRow) throw new Error('ЕНО: в ДБ не найдена дата ' + apiDate);

    const values = db.getRange(sheetRow, 1, 1, db.getLastColumn()).getValues()[0];
    const read = function(col) { return col > 0 ? tdmeToNumber_(values[col - 1]) : 0; };
    dbTotals.spam += read(header.callibriSpam);
    dbTotals.nonTarget += read(header.callibriNonTarget);
    dbTotals.leadA += read(header.callibriLeadA);
    dbTotals.leadB += read(header.callibriLeadB);
    dbTotals.leadC += read(header.callibriLeadC);
    hasLeadBColumn = hasLeadBColumn || header.callibriLeadB > 0;
  });

  if (Number(apiTotals.leadB || 0) > 0 && !hasLeadBColumn) {
    throw new Error('ЕНО: Callibri вернул Лид_Квал_B=' + apiTotals.leadB + ', но отдельной колонки B в ДБ нет. B не был прибавлен к A.');
  }

  tdmeAssertCallibriTotalsEqual_('ЕНО: Callibri API и ДБ не совпали', apiTotals, dbTotals);

  return;
}

function tdmeDirectRequest_(body) {
  try {
    const token = tdmeNormalizeToken_(
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

    const result = tdmeFetchWithRetry_(url, options, {
      maxAttempts: 8,
      sleepMs: 4000,
      retryCodes: [201, 202, 429, 500, 502, 503, 504]
    });

    return result.text;
  } catch (error) {
    tdmeSafeLog_('tdmeDirectRequest_', error);
    throw error;
  }
}

function tdmeApplyMetrikaAutomaticGoals_(rows, dateFrom, dateTo) {
  let metrika;

  try {
    metrika = tdmeFetchMetrikaAutomaticGoalsByCampaign_(dateFrom, dateTo);
  } catch (e) {
    Logger.log(
      'WARNING: ЕНО продолжен без дополнительного блока целей Метрики за ' +
      dateFrom + ' - ' + dateTo +
      '. Причина: ' + (e && e.message ? e.message : String(e))
    );

    // Важно: не роняем весь недельный отчёт.
    // Данные из Директа уже загружены, блок ЕНО должен собраться.
    return;
  }

  rows.forEach(row => {
    const key = tdmeCampaignKey_(row.campaignName);
    const goals = metrika.byCampaign[key] || tdmeEmptyMetrikaGoals_();

    row.goals.purchase = goals.purchase;
    row.goals.jivo = goals.jivo;
  });

  if (tdmeHasMetrikaGoals_(metrika.noCampaign)) {
    rows.push(tdmeBuildMetrikaNoCampaignRow_(metrika.noCampaign));
  }
}

function tdmeFetchMetrikaAutomaticGoalsByCampaign_(dateFrom, dateTo) {
  try {
    const token = tdmeNormalizeToken_(
      PropertiesService.getScriptProperties().getProperty('YANDEX_METRIKA_TOKEN') ||
      PropertiesService.getScriptProperties().getProperty('YANDEX_TOKEN')
    );

    if (!token) {
      throw new Error('Не найден YANDEX_METRIKA_TOKEN или YANDEX_TOKEN в Script Properties.');
    }

    if (!dateFrom || !dateTo) {
      throw new Error('Не задан период Метрики: dateFrom/dateTo.');
    }

    const params = {
      ids: TDME_METRIKA_COUNTER_ID,
      date1: dateFrom,
      date2: dateTo,
      dimensions: 'ym:s:automaticTrafficSource,ym:s:lastsignUTMCampaign',
      metrics: [
        'ym:s:goal504318735reaches',
        'ym:s:goal575188424reaches'
      ].join(','),
      accuracy: 'full',
      limit: 10000
    };

    const query = Object.keys(params)
      .map(key => encodeURIComponent(key) + '=' + encodeURIComponent(params[key]))
      .join('&');

    const result = tdmeFetchWithRetry_(
      'https://api-metrika.yandex.net/stat/v1/data?' + query,
      {
        method: 'get',
        headers: {
          Authorization: 'OAuth ' + token
        },
        muteHttpExceptions: true
      },
      {
        maxAttempts: 4,
        sleepMs: 3000,
        retryCodes: [429, 500, 502, 503, 504]
      }
    );

    const json = tdmeSafeJsonParse_(result.text, 'Метрика goals by campaign');

    const output = {
      byCampaign: {},
      noCampaign: tdmeEmptyMetrikaGoals_()
    };

    (json.data || []).forEach(item => {
      const dimensions = item && item.dimensions ? item.dimensions : [];
      const metrics = item && item.metrics ? item.metrics : [];
      const source = dimensions[0] || {};

      if (!tdmeIsMetrikaAdTraffic_(source)) {
        return;
      }

      const campaignName = dimensions[1] && dimensions[1].name;
      const goals = {
        purchase: tdmeToNumber_(metrics[0]),
        jivo: tdmeToNumber_(metrics[1])
      };

      if (!campaignName) {
        tdmeAddMetrikaGoals_(output.noCampaign, goals);
        return;
      }

      const key = tdmeCampaignKey_(campaignName);

      if (!output.byCampaign[key]) {
        output.byCampaign[key] = tdmeEmptyMetrikaGoals_();
      }

      tdmeAddMetrikaGoals_(output.byCampaign[key], goals);
    });

    return output;
  } catch (error) {
    tdmeSafeLog_('tdmeFetchMetrikaAutomaticGoalsByCampaign_', error);
    throw error;
  }
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
      callibriSpam: 0,
      callibriNonTarget: 0,
      addToCart: 0,
      purchase: goals.purchase,
      jivo: goals.jivo,
      callibriLeadA: 0,
      callibriLeadB: 0,
      callibriLeadC: 0
    }
  };
}

function tdmeEmptyMetrikaGoals_() {
  return {
    purchase: 0,
    jivo: 0
  };
}

function tdmeAddMetrikaGoals_(target, source) {
  target.purchase += Number(source.purchase || 0);
  target.jivo += Number(source.jivo || 0);
}

function tdmeHasMetrikaGoals_(goals) {
  return Number(goals.purchase || 0) +
    Number(goals.jivo || 0) > 0;
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
  const idxCampaignId = tdmeRequireColumn_(headers, 'CampaignId');
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
      campaignId: String(row[idxCampaignId] || '').trim(),
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
        callibriSpam: tdmeGoalSum_(row, goalIndexes, TDME_GOALS.callibriSpam),
        callibriNonTarget: tdmeGoalSum_(row, goalIndexes, TDME_GOALS.callibriNonTarget),
        addToCart: tdmeGoalSum_(row, goalIndexes, TDME_GOALS.addToCart),
        purchase: tdmeGoalSum_(row, goalIndexes, TDME_GOALS.purchase),
        jivo: tdmeGoalSum_(row, goalIndexes, TDME_GOALS.jivo),
        callibriLeadA: tdmeGoalSum_(row, goalIndexes, TDME_GOALS.callibriLeadA),
        callibriLeadB: 0,
        callibriLeadC: tdmeGoalSum_(row, goalIndexes, TDME_GOALS.callibriLeadC)
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

  // Ищем именно лишний legacy-столбец "Уникальный звонок".
  // Столбцы с историческими названиями не участвуют в расчёте фактических лидов.
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
    callibriSpam: tdmeContainsCol_(map, ['callibri: спам', 'callibri спам']),
    callibriNonTarget: tdmeContainsCol_(map, ['callibri: нецелевой_лид', 'callibri нецелевой лид', 'callibri нецелевой_лид']),
    addToCart: tdmeContainsCol_(map, ['ecommerce: добавление в корзину', 'ecommerce добавление в корзину']),
    purchase: tdmeContainsCol_(map, ['ecommerce: покупка', 'ecommerce покупка']),
    jivo: tdmeContainsCol_(map, ['jivo-чат', 'jivo чат', 'jivo']),
    callibriLeadA: tdmeContainsCol_(map, ['callibri: лид_квал_a', 'callibri лид квал a', 'callibri лид_квал_a']),
    callibriLeadB: tdmeContainsCol_(map, ['callibri: лид_квал_b', 'callibri лид квал b', 'callibri лид_квал_b']),
    callibriLeadC: tdmeContainsCol_(map, ['callibri: лид_квал_c', 'callibri лид квал c', 'callibri лид_квал_c'])
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


function tdmeLegacyEmailCol_(map) {
  return map['email'] || 0;
}

function tdmeLegacyCallsCol_(map) {
  return map['звонки'] || 0;
}

function tdmeApplyUpdatedGoalColumns_(sheet, headerRow, totalRow, header) {
  const lastCol = sheet.getLastColumn();

  // Новая схема целей: старые phone/email/messenger/legacy-трекеры не используются.
  if (lastCol >= 23) {
    sheet.getRange(headerRow, 16, 1, 8).setValues([[
      'Callibri: Спам',
      'Callibri: Нецелевой_Лид',
      'Ecommerce: добавление в корзину',
      'Ecommerce: покупка',
      'Jivo-чат',
      'Callibri: Лид_Квал_A',
      'Callibri: Лид_Квал_C',
      'Факт сумма конверсий'
    ]]);
  }
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
    header.callibriSpam,
    header.callibriNonTarget,
    header.addToCart,
    header.purchase,
    header.jivo,
    header.callibriLeadA,
    header.callibriLeadB,
    header.callibriLeadC
  ].filter(col => col > 0);
}

function tdmeClearDataInputCells_(sheet, header, startRow, rowsCount) {
  if (!sheet || !header || !startRow || !rowsCount || rowsCount <= 0) {
    Logger.log('[ENO] tdmeClearDataInputCells_: пустые параметры, очистка пропущена.');
    return;
  }

  const columns = tdmeInputColumns_(header);

  if (!columns.length) {
    Logger.log('[ENO] tdmeClearDataInputCells_: нет колонок для очистки.');
    return;
  }

  const ranges = columns.map(col => {
    const letter = tdmeColumnLetter_(col);
    return letter + startRow + ':' + letter + (startRow + rowsCount - 1);
  });

  sheet.getRangeList(ranges).clearContent();
}

function tdmeFillRows_(sheet, header, startRow, rows) {
  if (!sheet || !header || !startRow) {
    Logger.log('[ENO] tdmeFillRows_: пустые параметры, запись пропущена.');
    return;
  }

  if (!rows || !rows.length) {
    Logger.log('[ENO] tdmeFillRows_: rows пустой, запись строк кампаний пропущена.');
    return;
  }

  const columnWriters = [
    { col: header.rk, values: rows.map(item => item.campaignName) },
    { col: header.type, values: rows.map(item => item.campaignType) },
    { col: header.impressions, values: rows.map(item => item.impressions) },
    { col: header.clicks, values: rows.map(item => item.clicks) },
    { col: header.cost, values: rows.map(item => item.cost) },
    { col: header.avgBid, values: rows.map(item => item.avgBid) },
    { col: header.avgPosition, values: rows.map(item => item.avgPosition) },
    { col: header.avgTraffic, values: rows.map(item => item.avgTraffic) },
    { col: header.bounce, values: rows.map(item => item.bounce) },
    { col: header.depth, values: rows.map(item => item.depth) },
    { col: header.view3Pages, values: rows.map(item => tdmeGoalValue_((item.goals || {}).view3Pages)) },
    { col: header.callibriSpam, values: rows.map(item => tdmeGoalValue_((item.goals || {}).callibriSpam)) },
    { col: header.callibriNonTarget, values: rows.map(item => tdmeGoalValue_((item.goals || {}).callibriNonTarget)) },
    { col: header.addToCart, values: rows.map(item => tdmeGoalValue_((item.goals || {}).addToCart)) },
    { col: header.purchase, values: rows.map(item => tdmeGoalValue_((item.goals || {}).purchase)) },
    { col: header.jivo, values: rows.map(item => tdmeGoalValue_((item.goals || {}).jivo)) },
    { col: header.callibriLeadA, values: rows.map(item => tdmeGoalValue_((item.goals || {}).callibriLeadA)) },
    { col: header.callibriLeadB, values: rows.map(item => tdmeGoalValue_((item.goals || {}).callibriLeadB)) },
    { col: header.callibriLeadC, values: rows.map(item => tdmeGoalValue_((item.goals || {}).callibriLeadC)) }
  ];

  columnWriters.forEach(writer => {
    if (!writer.col || writer.col <= 0) return;

    const matrix = writer.values.map(value => [value]);
    sheet.getRange(startRow, writer.col, matrix.length, 1).setValues(matrix);
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
  if (!sheet || !header || !startRow || !rowsCount || rowsCount <= 0) {
    Logger.log('[ENO] tdmeHighlightGreenRows_: пустые параметры, форматирование пропущено.');
    return;
  }

  const lastCol = Math.max(
    header.callibriLeadC,
    header.callibriLeadB,
    header.callibriLeadA,
    header.jivo,
    header.purchase,
    header.addToCart,
    header.callibriNonTarget,
    header.callibriSpam,
    header.view3Pages,
    header.depth,
    header.cost,
    1
  );

  const backgrounds = [];
  const fontWeights = [];

  for (let i = 0; i < rowsCount; i++) {
    const item = rows && rows[i] ? rows[i] : { goals: {} };
    const goals = item.goals || {};

    const targetActions =
      Number(goals.addToCart || 0) +
      Number(goals.purchase || 0) +
      Number(goals.jivo || 0) +
      Number(goals.callibriLeadA || 0) +
      Number(goals.callibriLeadC || 0);

    const background = targetActions > 0 ? '#d9ead3' : '#ffffff';

    backgrounds.push(Array(lastCol).fill(background));
    fontWeights.push(Array(lastCol).fill('normal'));
  }

  sheet.getRange(startRow, 1, rowsCount, lastCol)
    .setBackgrounds(backgrounds)
    .setFontWeights(fontWeights);
}

// =====================
// КОММЕНТАРИИ
// =====================

function tdmeBuildComments_(weekRows, monthRows, dateFrom, dateTo) {
  const total = tdmeTotals_(weekRows);
  const monthTotal = tdmeTotals_(monthRows);

  // TDM 2026-07-06: бюджетный процент считаем как в ДБ — по плану месяца и рабочим дням.
  const monthPlan = tdmeMonthPlanForDate_(dateTo);
  const planToDate = tdmePlanToDate_(monthPlan, dateTo);
  const budgetPercent = planToDate > 0 ? monthTotal.cost / planToDate * 100 : 0;

  const byType = tdmeGroupByType_(weekRows);
  const byGeo = tdmeGroupByGeo_(weekRows);

  const rows = [];

  rows.push(['Комментарии к отчёту за прошлую неделю с ' + tdmeShortDate_(dateFrom) + ' по ' + tdmeDateRu_(dateTo), '']);
  rows.push(['', '']);

  rows.push(['Бюджет:', '']);
  rows.push([
    'Всего потрачено за неделю ' + tdmeRub_(total.cost) +
    '. По бюджету месяца на дату как в ДБ: план ' + tdmeRub_(planToDate) +
    ', факт ' + tdmeRub_(monthTotal.cost) +
    ', выполнение ' + tdmePercent_(budgetPercent),
    ''
  ]);
  rows.push(['', '']);

  rows.push(['Трафик:', '']);
  rows.push(['Показы — ' + tdmeInt_(total.impressions), '']);
  rows.push(['Клики — ' + tdmeInt_(total.clicks), '']);
  rows.push(['Средний CPC — ' + tdmeRub_(total.clicks > 0 ? total.cost / total.clicks : 0), '']);
  rows.push(['', '']);

  rows.push(['Достигнуты цели:', '']);
  rows.push(['Просмотр 3х страниц — ' + total.goals.view3Pages, '']);
  rows.push(['Callibri: Спам — ' + total.goals.callibriSpam, '']);
  rows.push(['Callibri: Нецелевой_Лид — ' + total.goals.callibriNonTarget, '']);
  rows.push(['Ecommerce: добавление в корзину — ' + total.goals.addToCart, '']);
  rows.push(['Ecommerce: покупка — ' + total.goals.purchase, '']);
  rows.push(['Jivo-чат — ' + total.goals.jivo, '']);
  rows.push(['Callibri: Лид_Квал_A — ' + total.goals.callibriLeadA, '']);
  rows.push(['Callibri: Лид_Квал_C — ' + total.goals.callibriLeadC, '']);
  rows.push(['Факт лидов — ' + tdmeTargetConversions_(total), '']);
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
      ', Jivo-чат — ' +
      item.goals.jivo +
      ', Callibri A — ' +
      item.goals.callibriLeadA +
      ', Callibri C — ' +
      item.goals.callibriLeadC,
      ''
    ]);
  });

  rows.push(['', '']);
  rows.push(['Вывод:', '']);
  rows.push(['Зелёным выделены кампании, которые принесли целевые действия. Кампании с расходом без целевых действий проверяем по запросам, площадкам, объявлениям и посадочным страницам.', '']);

  return rows;
}

function tdmeFindCommentRow_(sheet, totalRow) {
  if (!sheet || !totalRow) return 0;

  const maxRow = Math.min(sheet.getLastRow(), totalRow + 80);
  const rowsCount = Math.max(maxRow - totalRow, 0);

  if (!rowsCount) return 0;

  const values = sheet
    .getRange(totalRow + 1, 1, rowsCount, 1)
    .getDisplayValues();

  for (let i = 0; i < values.length; i++) {
    const text = tdmeNormalize_(values[i][0]);

    if (text.indexOf('комментар') !== -1) {
      return totalRow + 1 + i;
    }
  }

  return 0;
}

function tdmeWriteComments_(sheet, startRow, rows) {
  sheet.getRange(startRow, 1, 80, 8)
    .clearContent()
    .setBackground('#ffffff')
    // В скопированном шаблоне шрифт бывает белым: явно возвращаем видимый текст.
    .setFontColor('#000000')
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
  target.goals.callibriSpam += row.goals.callibriSpam;
  target.goals.callibriNonTarget += row.goals.callibriNonTarget;
  target.goals.addToCart += row.goals.addToCart;
  target.goals.purchase += row.goals.purchase;
  target.goals.jivo += row.goals.jivo;
  target.goals.callibriLeadA += row.goals.callibriLeadA;
  target.goals.callibriLeadB += Number(row.goals.callibriLeadB || 0);
  target.goals.callibriLeadC += row.goals.callibriLeadC;
}

function tdmeEmptyTotal_() {
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
      callibriLeadB: 0,
      callibriLeadC: 0
    }
  };
}

function tdmeTargetConversions_(item) {
  if (!item || !item.goals) return 0;

  // Факт лидов: Ecommerce purchase + Jivo + Callibri A + Callibri C.
  // Спам и нецелевые в факт не включаем.
  return Number(item.goals.purchase || 0) +
    Number(item.goals.jivo || 0) +
    Number(item.goals.callibriLeadA || 0) +
    Number(item.goals.callibriLeadC || 0);
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
  // TDM 2026-07-06: как в ДБ — план на дату считается по рабочим дням месяца.
  const date = tdmeParseDate_(dateTo);
  const firstDay = new Date(date.getFullYear(), date.getMonth(), 1);
  const lastDay = new Date(date.getFullYear(), date.getMonth() + 1, 0);
  const workingDays = tdmeWorkingDaysBetween_(firstDay, lastDay);
  const passedWorkingDays = tdmeWorkingDaysBetween_(firstDay, date);

  if (!workingDays) return 0;
  return monthPlan / workingDays * passedWorkingDays;
}

function tdmeMonthPlanForDate_(apiDate) {
  // Июль 2026 в ДБ: 280 000 ₽ с НДС. Старые месяцы оставляем на базовом плане ЕНО.
  if (String(apiDate) >= '2026-07-01') return 280000;
  return TDME_MONTH_PLAN;
}

function tdmeWorkingDaysBetween_(dateFrom, dateTo) {
  let count = 0;
  const current = new Date(dateFrom);
  const end = new Date(dateTo);

  while (current <= end) {
    const day = current.getDay();
    if (day !== 0 && day !== 6) count++;
    current.setDate(current.getDate() + 1);
  }

  return count;
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
      const firstCell = sheet.getRange(block.totalRow, 1);
      if (!firstCell.getFormula()) {
        firstCell.setValue('ИТОГО');
      }

      const finalValues = sheet.getRange(block.totalRow, 1, 1, lastCol).getDisplayValues()[0];
      const finalHasTotal = finalValues.some(value => tdmeNormalize_(value) === 'итого');

      if (!finalHasTotal) {
        throw new Error('Проверка ЕНО: в строке ' + block.totalRow + ' нет метки ИТОГО.');
      }
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
  let formulas = range.getFormulas()[0];

  if (!formulas.some(function(formula) { return Boolean(formula); })) {
    let templateFormulas = null;
    const firstRow = Math.max(1, totalRow - 250);

    for (let row = totalRow - 1; row >= firstRow; row--) {
      const values = sheet.getRange(row, 1, 1, lastCol).getDisplayValues()[0];
      const hasTotal = values.some(function(value) {
        return tdmeNormalizeForTotalSearch_(value) === 'итого';
      });
      if (!hasTotal) continue;

      const candidate = sheet.getRange(row, 1, 1, lastCol).getFormulas()[0];
      if (candidate.some(function(formula) { return Boolean(formula); })) {
        templateFormulas = candidate;
        break;
      }
    }

    if (!templateFormulas) {
      throw new Error('ЕНО: не удалось восстановить формулы строки ИТОГО ' + totalRow + ' — выше не найден шаблон с формулами.');
    }

    formulas = templateFormulas;
  }

  const fixed = formulas.map(function(formula) {
    if (!formula) return formula;
    return tdmeFixFormulaRangesToCurrentBlock_(formula, dataStartRow, dataEndRow);
  });

  range.setFormulas([fixed]);

  const verified = range.getFormulas()[0];
  if (!verified.some(function(formula) { return Boolean(formula); })) {
    throw new Error('ЕНО: формулы строки ИТОГО не восстановились после записи, строка ' + totalRow + '.');
  }
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
  if (!sheet || !headerRow) {
    throw new Error('tdmeFindTotalRow_: не передан sheet/headerRow.');
  }

  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();
  const maxRow = Math.min(lastRow, headerRow + 180);
  const rowsCount = Math.max(maxRow - headerRow, 0);

  if (!rowsCount) {
    throw new Error('tdmeFindTotalRow_: пустой диапазон поиска после строки ' + headerRow);
  }

  const values = sheet
    .getRange(headerRow + 1, 1, rowsCount, lastCol)
    .getDisplayValues();

  for (let i = 0; i < values.length; i++) {
    const hasTotal = values[i].some(value => {
      return tdmeNormalizeForTotalSearch_(value) === 'итого';
    });

    if (hasTotal) {
      return headerRow + 1 + i;
    }
  }

  const commentRow = tdmeFindEnoCommentRowAfterHeader_(sheet, headerRow, maxRow, lastCol);

  if (commentRow) {
    for (let row = commentRow - 1; row > headerRow; row--) {
      const rowValues = sheet.getRange(row, 1, 1, lastCol).getDisplayValues()[0];
      const nonEmptyCount = rowValues.filter(value => String(value || '').trim() !== '').length;

      if (nonEmptyCount >= 2) {
        const cell = sheet.getRange(row, 1);
        const firstCell = cell.getDisplayValue();

        if (tdmeNormalizeForTotalSearch_(firstCell) !== 'итого' && !cell.getFormula()) {
          cell.setValue('ИТОГО');
        }

        return row;
      }
    }
  }

  const calculatedTotalRow = tdmeGuessTotalRowByCampaignRows_(sheet, headerRow, maxRow, lastCol);

  if (calculatedTotalRow) {
    const cell = sheet.getRange(calculatedTotalRow, 1);

    if (!cell.getFormula() && tdmeNormalizeForTotalSearch_(cell.getDisplayValue()) !== 'итого') {
      cell.setValue('ИТОГО');
    }

    return calculatedTotalRow;
  }

  throw new Error(
    'Не найдена строка ИТОГО после строки ' + headerRow +
    '. Диапазон поиска: строки ' + (headerRow + 1) + '–' + maxRow + '.'
  );
}

function tdmeFindEnoCommentRowAfterHeader_(sheet, headerRow, maxRow, lastCol) {
  if (!sheet || !headerRow || !maxRow || !lastCol) return 0;

  const rowsCount = Math.max(maxRow - headerRow, 0);
  if (!rowsCount) return 0;

  const values = sheet
    .getRange(headerRow + 1, 1, rowsCount, lastCol)
    .getDisplayValues();

  for (let i = 0; i < values.length; i++) {
    const hasComment = values[i].some(value => {
      const text = tdmeNormalizeForTotalSearch_(value);
      return text.indexOf('комментарии к отчету') !== -1 ||
        text.indexOf('комментарии к отчёту') !== -1 ||
        text.indexOf('комментарии') !== -1;
    });

    if (hasComment) {
      return headerRow + 1 + i;
    }
  }

  return 0;
}

function tdmeGuessTotalRowByCampaignRows_(sheet, headerRow, maxRow, lastCol) {
  if (!sheet || !headerRow || !maxRow || !lastCol) return 0;

  const rowsCount = Math.max(maxRow - headerRow, 0);
  if (!rowsCount) return 0;

  const values = sheet
    .getRange(headerRow + 1, 1, rowsCount, lastCol)
    .getDisplayValues();

  let lastDataRow = 0;

  for (let i = 0; i < values.length; i++) {
    const rowValues = values[i];
    const sheetRow = headerRow + 1 + i;
    const firstCell = tdmeNormalizeForTotalSearch_(rowValues[0]);
    const nonEmptyCount = rowValues.filter(value => String(value || '').trim() !== '').length;
    const rowText = tdmeNormalizeForTotalSearch_(rowValues.join(' '));

    if (
      rowText.indexOf('комментарии') !== -1 ||
      rowText.indexOf('отчет недельный') !== -1 ||
      rowText.indexOf('отчёт недельный') !== -1
    ) {
      break;
    }

    if (firstCell && firstCell !== 'итого' && nonEmptyCount >= 4) {
      lastDataRow = sheetRow;
    }

    if (!firstCell && nonEmptyCount === 0 && lastDataRow) {
      break;
    }
  }

  return lastDataRow ? lastDataRow + 1 : 0;
}

/**
 * TDM 2026-07-06: аварийное восстановление ЕНО 29.06–05.07 под схему ДБ.
 */
function archived_tdmFixEnoCurrentWeekLikeDb20260706() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('ЕНО');
  if (!sheet) throw new Error('Не найден лист ЕНО');
  const topRow = 917, headerRow = 919, startRow = 920, totalRow = 939;

  sheet.getRange(headerRow, 16, 1, 10).setValues([[
    'Callibri: Спам','Callibri: Нецелевой_Лид','Ecommerce: добавление в корзину','Ecommerce: покупка','Jivo-чат','Callibri: Лид_Квал_A','Callibri: Лид_Квал_C','Факт сумма конверсий','CR','Стоимость факт CPA'
  ]]);

  const byCampaign = {
    'tdm_search_opt_rf': {spam:0,non:0,add:0,buy:0,jivo:1,a:0,c:0},
    'tdm_top-categ_dsa_spb': {spam:1,non:0,add:2,buy:0,jivo:1,a:0,c:0},
    'tdm_tovarnay_general_spb_ozk': {spam:0,non:0,add:1,buy:0,jivo:1,a:0,c:0},
    'tdm_search_opt_spb': {spam:0,non:0,add:1,buy:0,jivo:1,a:1,c:0},
    'tdm_top-categ_t-g_rf': {spam:0,non:0,add:0,buy:0,jivo:0,a:0,c:0},
    'tdm_tovarnay_general_rf_ozk': {spam:0,non:0,add:0,buy:0,jivo:1,a:1,c:1},
    'tdm_tovarnay_general_kz_ozk': {spam:0,non:0,add:3,buy:0,jivo:1,a:0,c:0},
    'tdm_search_categ_rf': {spam:0,non:0,add:0,buy:0,jivo:0,a:1,c:0},
    'tdm_search_vendor_rf_spb': {spam:0,non:0,add:0,buy:0,jivo:1,a:0,c:1},
    'tdm_smart-banners_rsya_spb-rf': {spam:0,non:0,add:2,buy:1,jivo:1,a:0,c:0},
    'tdm_search_opt_kz': {spam:0,non:0,add:0,buy:0,jivo:0,a:0,c:0},
    'tdm_top-categ_t-g_spb': {spam:0,non:0,add:1,buy:0,jivo:1,a:0,c:0},
    'tdm_search_proizv_spb-rf': {spam:0,non:0,add:0,buy:0,jivo:0,a:0,c:0},
    'tdm_search_categ_spb': {spam:0,non:0,add:0,buy:0,jivo:0,a:0,c:0},
    'tdm_top-categ_dsa_t-g_kz': {spam:0,non:0,add:0,buy:0,jivo:0,a:0,c:0},
    'tdm_top-categ_dsa_rf': {spam:0,non:0,add:0,buy:0,jivo:0,a:0,c:0},
    'tdm_search_new_spb': {spam:0,non:0,add:0,buy:0,jivo:0,a:1,c:0},
    'tdm_search_categ_kz': {spam:0,non:0,add:0,buy:0,jivo:0,a:0,c:0},
    'tdm_search_stroit_rf_spb': {spam:0,non:0,add:0,buy:0,jivo:0,a:0,c:0}
  };

  for (let row = startRow; row < totalRow; row++) {
    const name = String(sheet.getRange(row, 1).getDisplayValue() || '').trim();
    const item = byCampaign[name] || {spam:0,non:0,add:0,buy:0,jivo:0,a:0,c:0};
    const vals = [item.spam,item.non,item.add,item.buy,item.jivo,item.a,item.c].map(function(v){return Number(v||0)>0?Number(v||0):'-';});
    sheet.getRange(row,16,1,7).setValues([vals]);
    sheet.getRange(row,23).setFormula('=SUM(S'+row+':V'+row+')');
    sheet.getRange(row,24).setFormula('=IFERROR(W'+row+'/D'+row+';0)');
    sheet.getRange(row,25).setFormula('=IFERROR(G'+row+'/W'+row+';0)');
  }
  ['P','Q','R','S','T','U','V'].forEach(function(col,idx){sheet.getRange(totalRow,16+idx).setFormula('=SUM('+col+startRow+':'+col+(totalRow-1)+')');});
  sheet.getRange(totalRow,23).setFormula('=SUM(W'+startRow+':W'+(totalRow-1)+')');
  sheet.getRange(totalRow,24).setFormula('=IFERROR(W'+totalRow+'/D'+totalRow+';0)');
  sheet.getRange(totalRow,25).setFormula('=IFERROR(G'+totalRow+'/W'+totalRow+';0)');
  sheet.getRange(topRow,20,1,6).setValues([['Кол-во Лид','=W'+totalRow,'CR','=X'+totalRow,'CPA Лид','=Y'+totalRow]]);
  sheet.getRange(945,1).setValue('Всего потрачено за неделю 74 395,38 руб. По июлю на дату как в ДБ: план 36 521,74 руб., факт 40 840,22 руб., выполнение 111,82%.');

  sheet.getRange(headerRow,16,1,10).setFontColor('#000000').setFontWeight('bold').setHorizontalAlignment('center').setVerticalAlignment('middle').setWrap(true);
  sheet.getRange(headerRow,16).setBackground('#f4cccc');
  sheet.getRange(headerRow,17).setBackground('#cfe2f3');
  sheet.getRange(headerRow,18,1,2).setBackground('#d9ead3');
  sheet.getRange(headerRow,20).setBackground('#b6d7a8');
  sheet.getRange(headerRow,21).setBackground('#d9ead3');
  sheet.getRange(headerRow,22).setBackground('#d9d2e9');
  sheet.getRange(headerRow,23,1,3).setBackground('#fff2cc');
  sheet.getRange(startRow,11,totalRow-startRow+1,1).setNumberFormat('0.00%');
  sheet.getRange(startRow,23,totalRow-startRow+1,1).setNumberFormat('0');
  sheet.getRange(startRow,24,totalRow-startRow+1,1).setNumberFormat('0.00%');
  sheet.getRange(startRow,25,totalRow-startRow+1,1).setNumberFormat('"р."#,##0.00');
  SpreadsheetApp.flush();
  return {ok:true,sheet:'ЕНО',headerRow:headerRow,totalRow:totalRow};
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

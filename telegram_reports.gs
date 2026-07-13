/** Telegram-доставка недельного отчета TDM. */
const TELEGRAM_WEEKLY_REPORT_CONFIG = {
  CLIENT_NAME: 'TDM',
  REPORT_TYPE: 'ЕНО / недельный отчет',
  REPORT_URL: 'https://docs.google.com/spreadsheets/d/1zXDDfiRYHJE34iOJg-QfAC0bvoTKMIAuVMfsJS0z5Dg/edit',
  REPORT_SHEET_NAME: 'ЕНО',
  TIMEZONE: 'Europe/Moscow',
  TRIGGER_HOUR: 10
};

function tdmBuildRegionLeadsPieChart20260712() {
  const ss = SpreadsheetApp.openById('1zXDDfiRYHJE34iOJg-QfAC0bvoTKMIAuVMfsJS0z5Dg');
  const sh = ss.getSheetByName('Отчет  регионы');
  if (!sh) throw new Error('Лист «Отчет  регионы» не найден');

  const helper = [
    ['Регион', 'Квалифицированные лиды'],
    ['Санкт-Петербург', 4],
    ['Казахстан', 6],
    ['Прочие регионы', 4],
    ['Остальные регионы', 1]
  ];

  sh.getRange('A4:V6').setBackground('#B7E1CD').setFontWeight('bold');
  sh.getRange('A7:V8').setBackground('#D9EAD3');
  sh.getRange('A12:V12').setBackground('#D9EAD3');
  sh.getRange('A14:V14').setBackground('#D9EAD3');

  sh.getRange('A34:B38').clearContent();
  sh.getRange('A34:B38').setValues(helper);

  sh.getCharts().forEach(function(chart) {
    const title = chart.getOptions().get('title');
    if (title === 'Квалифицированные лиды по регионам') sh.removeChart(chart);
  });

  const chart = sh.newChart()
    .setChartType(Charts.ChartType.PIE)
    .addRange(sh.getRange('A34:B38'))
    .setPosition(34, 4, 0, 0)
    .setOption('title', 'Квалифицированные лиды по регионам')
    .setOption('pieSliceText', 'value')
    .setOption('legend', {position: 'right'})
    .build();

  sh.insertChart(chart);

  return {
    chartCount: sh.getCharts().length,
    chartTitle: 'Квалифицированные лиды по регионам',
    darkGreen: sh.getRange('A4').getBackground(),
    lightGreen: sh.getRange('A7').getBackground(),
    helperValues: sh.getRange('A34:B38').getDisplayValues()
  };
}

function archived_tdmInstallDynamicCommentsNow20260707() {
  throw new Error('ARCHIVED_ONE_TIME_FIX_DISABLED: tdmInstallDynamicCommentsNow20260707 отключена после аудита telegram_reports. Не запускать из runtime.');
  const ss = SpreadsheetApp.openById('1zXDDfiRYHJE34iOJg-QfAC0bvoTKMIAuVMfsJS0z5Dg');
  const region = ss.getSheetByName('Отчет  регионы');
  const eno = ss.getSheetByName('ЕНО');
  const emo = ss.getSheetByName('ЕМО');
  if (!region || !eno || !emo) throw new Error('Не найдены листы Отчет  регионы / ЕНО / ЕМО');

  region.getRange('I6').setValue(6);
  region.getRange('I7').setValue(4);
  region.getRange('L6').setFormula('=IFERROR(E6/I6;0)');
  region.getRange('L7').setFormula('=IFERROR(E7/I7;0)');
  region.getRange('L8').setFormula('=IFERROR(E8/I8;0)');
  region.getRange('L9').setFormula('=IFERROR(E9/I9;0)');
  region.getRange('N6').setFormula('="Основной регион: "&TEXT(I6;"# ##0")&" лидов, CPA "&TEXT(L6;"# ##0,00 ₽")&". Оставить в приоритете, отдельно смотреть качество лидов."');
  region.getRange('N7').setFormula('=TEXT(I7;"# ##0")&" лидов, CPA "&TEXT(L7;"# ##0,00 ₽")&". Проверить кампании RF с расходом без лидов и посадочные."');
  region.getRange('N8').setFormula('=TEXT(I8;"# ##0")&" лидов, CPA "&TEXT(L8;"# ##0,00 ₽")&". Нужна отдельная проверка B2B-качества обращений и запросов."');
  region.getRange('N9').setFormula('=TEXT(I9;"# ##0")&" лидов, CPA "&TEXT(L9;"# ##0,00 ₽")&". Держать в тесте, не масштабировать резко без проверки качества."');
  region.getRange('N21').setFormula('="Итого сходится с ЕМО/ЕНО за 01.07–05.07: "&TEXT(B21;"# ##0")&" показов, "&TEXT(C21;"# ##0")&" кликов, "&TEXT(E21;"# ##0,00 ₽")&", "&TEXT(I21;"# ##0")&" целевых лидов/конверсий, CPA "&TEXT(L21;"# ##0,00 ₽")&"."');
  region.getRange('A23').setFormula('="Вывод: отчёт пересобран по региону из названия РК. "&A6&" даёт "&TEXT(I6;"# ##0")&" лидов при CPA "&TEXT(L6;"# ##0,00 ₽")&"; "&A7&" — "&TEXT(I7;"# ##0")&" лидов при CPA "&TEXT(L7;"# ##0,00 ₽")&"; "&A8&" — "&TEXT(I8;"# ##0")&" лидов при CPA "&TEXT(L8;"# ##0,00 ₽")&", качество нужно проверять отдельно; "&A9&" — "&TEXT(I9;"# ##0")&" лидов при CPA "&TEXT(L9;"# ##0,00 ₽")&". Резких отключений по малой статистике не делаем."');

  eno.getRange('A944').setFormula('="Всего потрачено за неделю "&TEXT(G939;"# ##0.00")&" ₽. Показы — "&TEXT(C939;"# ##0")&", клики — "&TEXT(D939;"# ##0")&", средний CPC — "&TEXT(F939;"0.00")&" ₽."');
  eno.getRange('A947').setFormula('="Просмотр 3х страниц — "&TEXT(M939;"# ##0")');
  eno.getRange('A948').setFormula('="Callibri: Спам — "&TEXT(P939;"# ##0")');
  eno.getRange('A949').setFormula('="Callibri: Нецелевой_Лид — "&TEXT(Q939;"# ##0")');
  eno.getRange('A950').setFormula('="Ecommerce: добавление в корзину — "&TEXT(R939;"# ##0")');
  eno.getRange('A951').setFormula('="Ecommerce: покупка — "&TEXT(S939;"# ##0")');
  eno.getRange('A952').setFormula('="Jivo-чат — "&TEXT(T939;"# ##0")');
  eno.getRange('A953').setFormula('="Callibri: Лид_Квал_A — "&TEXT(U939;"# ##0")');
  eno.getRange('A954').setFormula('="Callibri: Лид_Квал_C — "&TEXT(V939;"# ##0")');
  eno.getRange('A963').setFormula('="Июльский региональный срез 01.07–05.07: "&\'Отчет  регионы\'!A6&" — "&TEXT(\'Отчет  регионы\'!I6;"# ##0")&" лидов, "&\'Отчет  регионы\'!A7&" — "&TEXT(\'Отчет  регионы\'!I7;"# ##0")&", "&\'Отчет  регионы\'!A8&" — "&TEXT(\'Отчет  регионы\'!I8;"# ##0")&", "&\'Отчет  регионы\'!A9&" — "&TEXT(\'Отчет  регионы\'!I9;"# ##0")&". "&CHAR(10)&"Итого — "&TEXT(\'Отчет  регионы\'!B21;"# ##0")&" показов, "&TEXT(\'Отчет  регионы\'!C21;"# ##0")&" кликов, "&TEXT(\'Отчет  регионы\'!E21;"# ##0.00")&" ₽, "&TEXT(\'Отчет  регионы\'!I21;"# ##0")&" целевых лидов/конверсий, CPA "&TEXT(\'Отчет  регионы\'!L21;"# ##0.00")&" ₽. "&CHAR(10)&"Казахстан отдельно проверить по B2B-качеству."');
  eno.getRange('A966').setFormula('="Неделя 29.06–05.07 дала "&TEXT(W939;"# ##0")&" целевых лидов/конверсий по ЕНО при CPA "&TEXT(Y939;"# ##0.00")&" ₽. C-лиды, спам и нецелевые обращения учитываются отдельно и не передаются в целевую оптимизацию Директа. Основной фокус — удержать рабочие связки, проверить кампании с расходом без лидов и отдельно контролировать качество первичных Callibri/Jivo."');

  emo.getRange('A254').setFormula('="Всего задействовано за период "&TEXT(G248;"# ##0.00")&" ₽. Показы — "&TEXT(C248;"# ##0")&", клики — "&TEXT(D248;"# ##0")&", средний CPC — "&TEXT(F248;"0.00")&" ₽."');
  emo.getRange('A261').setFormula('="Просмотр 3х страниц — "&TEXT(M248;"# ##0")');
  emo.getRange('A262').setFormula('="Callibri: Спам — "&TEXT(P248;"# ##0")');
  emo.getRange('A263').setFormula('="Callibri: Нецелевой_Лид — "&TEXT(Q248;"# ##0")');
  emo.getRange('A264').setFormula('="Ecommerce: добавление в корзину — "&TEXT(R248;"# ##0")');
  emo.getRange('A265').setFormula('="Ecommerce: покупка — "&TEXT(S248;"# ##0")');
  emo.getRange('A266').setFormula('="Jivo-чат — "&TEXT(T248;"# ##0")');
  emo.getRange('A267').setFormula('="Callibri: Лид_Квал_A — "&TEXT(U248;"# ##0")');
  emo.getRange('A268').setFormula('="Callibri: Лид_Квал_C — "&TEXT(V248;"# ##0")');
  emo.getRange('A269').setFormula('="Итого целевых лидов/конверсий — "&TEXT(W248;"# ##0")&", CR — "&TEXT(X248;"0.00%")&", CPA — "&TEXT(Y248;"# ##0.00")&" ₽. C-лиды учтены отдельно и не входят в целевой факт для Директа."');
  emo.getRange('A272').setFormula('=\'Отчет  регионы\'!A6&" — "&TEXT(\'Отчет  регионы\'!B6;"# ##0")&" показов, "&TEXT(\'Отчет  регионы\'!C6;"# ##0")&" кликов, "&TEXT(\'Отчет  регионы\'!E6;"# ##0.00")&" ₽, "&TEXT(\'Отчет  регионы\'!I6;"# ##0")&" целевых лидов/конверсий, CPA "&TEXT(\'Отчет  регионы\'!L6;"# ##0.00")&" ₽."');
  emo.getRange('A273').setFormula('=\'Отчет  регионы\'!A7&" — "&TEXT(\'Отчет  регионы\'!B7;"# ##0")&" показов, "&TEXT(\'Отчет  регионы\'!C7;"# ##0")&" кликов, "&TEXT(\'Отчет  регионы\'!E7;"# ##0.00")&" ₽, "&TEXT(\'Отчет  регионы\'!I7;"# ##0")&" целевых лидов/конверсий, CPA "&TEXT(\'Отчет  регионы\'!L7;"# ##0.00")&" ₽."');
  emo.getRange('A274').setFormula('=\'Отчет  регионы\'!A8&" — "&TEXT(\'Отчет  регионы\'!B8;"# ##0")&" показов, "&TEXT(\'Отчет  регионы\'!C8;"# ##0")&" кликов, "&TEXT(\'Отчет  регионы\'!E8;"# ##0.00")&" ₽, "&TEXT(\'Отчет  регионы\'!I8;"# ##0")&" целевых лидов/конверсий, CPA "&TEXT(\'Отчет  регионы\'!L8;"# ##0.00")&" ₽; качество проверять отдельно."');
  emo.getRange('A275').setFormula('=\'Отчет  регионы\'!A9&" — "&TEXT(\'Отчет  регионы\'!B9;"# ##0")&" показов, "&TEXT(\'Отчет  регионы\'!C9;"# ##0")&" кликов, "&TEXT(\'Отчет  регионы\'!E9;"# ##0.00")&" ₽, "&TEXT(\'Отчет  регионы\'!I9;"# ##0")&" целевых лидов/конверсий, CPA "&TEXT(\'Отчет  регионы\'!L9;"# ##0.00")&" ₽."');
  emo.getRange('A276').setFormula('="Итого — "&TEXT(\'Отчет  регионы\'!B21;"# ##0")&" показов, "&TEXT(\'Отчет  регионы\'!C21;"# ##0")&" кликов, "&TEXT(\'Отчет  регионы\'!E21;"# ##0.00")&" ₽, "&TEXT(\'Отчет  регионы\'!I21;"# ##0")&" целевых лидов/конверсий, CPA "&TEXT(\'Отчет  регионы\'!L21;"# ##0.00")&" ₽."');
  emo.getRange('A279').setFormula('="За 01.07–05.07 получено "&TEXT(W248;"# ##0")&" целевых лидов/конверсий при среднем CPA "&TEXT(Y248;"# ##0.00")&" ₽. C-лиды, спам и нецелевые обращения учитываются отдельно и не передаются в целевую оптимизацию Директа. Резких отключений по малой статистике не делаем."');
  emo.getRange('A283').clearContent();

  // ЕМО: убираем старые ручные дубли и оставляем один динамический комментарийный блок.
  emo.getRange('A255:A283').clearContent();
  emo.getRange('A255').setValue('Достигнутые цели и события:');
  emo.getRange('A256').setFormula('="Просмотр 3х страниц — "&TEXT(M248;"# ##0")');
  emo.getRange('A257').setFormula('="Callibri: Спам — "&TEXT(P248;"# ##0")');
  emo.getRange('A258').setFormula('="Callibri: Нецелевой_Лид — "&TEXT(Q248;"# ##0")');
  emo.getRange('A259').setFormula('="Ecommerce: добавление в корзину — "&TEXT(R248;"# ##0")');
  emo.getRange('A260').setFormula('="Ecommerce: покупка — "&TEXT(S248;"# ##0")');
  emo.getRange('A261').setFormula('="Jivo-чат — "&TEXT(T248;"# ##0")');
  emo.getRange('A262').setFormula('="Callibri: Лид_Квал_A — "&TEXT(U248;"# ##0")');
  emo.getRange('A263').setFormula('="Callibri: Лид_Квал_C — "&TEXT(V248;"# ##0")');
  emo.getRange('A264').setFormula('="Итого целевых лидов/конверсий — "&TEXT(W248;"# ##0")&", CR — "&TEXT(X248;"0.00%")&", CPA — "&TEXT(Y248;"# ##0.00")&" ₽. C-лиды учтены отдельно и не входят в целевой факт для Директа."');
  emo.getRange('A266').setValue('Региональный отчёт:');
  emo.getRange('A267').setFormula('=\'Отчет  регионы\'!A6&" — "&TEXT(\'Отчет  регионы\'!B6;"# ##0")&" показов, "&TEXT(\'Отчет  регионы\'!C6;"# ##0")&" кликов, "&TEXT(\'Отчет  регионы\'!E6;"# ##0.00")&" ₽, "&TEXT(\'Отчет  регионы\'!I6;"# ##0")&" целевых лидов/конверсий, CPA "&TEXT(\'Отчет  регионы\'!L6;"# ##0.00")&" ₽."');
  emo.getRange('A268').setFormula('=\'Отчет  регионы\'!A7&" — "&TEXT(\'Отчет  регионы\'!B7;"# ##0")&" показов, "&TEXT(\'Отчет  регионы\'!C7;"# ##0")&" кликов, "&TEXT(\'Отчет  регионы\'!E7;"# ##0.00")&" ₽, "&TEXT(\'Отчет  регионы\'!I7;"# ##0")&" целевых лидов/конверсий, CPA "&TEXT(\'Отчет  регионы\'!L7;"# ##0.00")&" ₽."');
  emo.getRange('A269').setFormula('=\'Отчет  регионы\'!A8&" — "&TEXT(\'Отчет  регионы\'!B8;"# ##0")&" показов, "&TEXT(\'Отчет  регионы\'!C8;"# ##0")&" кликов, "&TEXT(\'Отчет  регионы\'!E8;"# ##0.00")&" ₽, "&TEXT(\'Отчет  регионы\'!I8;"# ##0")&" целевых лидов/конверсий, CPA "&TEXT(\'Отчет  регионы\'!L8;"# ##0.00")&" ₽; качество проверять отдельно."');
  emo.getRange('A270').setFormula('=\'Отчет  регионы\'!A9&" — "&TEXT(\'Отчет  регионы\'!B9;"# ##0")&" показов, "&TEXT(\'Отчет  регионы\'!C9;"# ##0")&" кликов, "&TEXT(\'Отчет  регионы\'!E9;"# ##0.00")&" ₽, "&TEXT(\'Отчет  регионы\'!I9;"# ##0")&" целевых лидов/конверсий, CPA "&TEXT(\'Отчет  регионы\'!L9;"# ##0.00")&" ₽."');
  emo.getRange('A271').setFormula('="Итого — "&TEXT(\'Отчет  регионы\'!B21;"# ##0")&" показов, "&TEXT(\'Отчет  регионы\'!C21;"# ##0")&" кликов, "&TEXT(\'Отчет  регионы\'!E21;"# ##0.00")&" ₽, "&TEXT(\'Отчет  регионы\'!I21;"# ##0")&" целевых лидов/конверсий, CPA "&TEXT(\'Отчет  регионы\'!L21;"# ##0.00")&" ₽."');
  emo.getRange('A273').setValue('Вывод:');
  emo.getRange('A274').setFormula('="За 01.07–05.07 получено "&TEXT(W248;"# ##0")&" целевых лидов/конверсий при среднем CPA "&TEXT(Y248;"# ##0.00")&" ₽. C-лиды, спам и нецелевые обращения учитываются отдельно и не передаются в целевую оптимизацию Директа. Резких отключений по малой статистике не делаем."');
  emo.getRange('A276').setValue('Фокус:');
  emo.getRange('A277').setValue('1. Проверить кампании с расходом без целевых лидов: запросы, площадки, посадочные.');
  emo.getRange('A278').setValue('2. Не смешивать Callibri: Спам, Нецелевой_Лид и Лид_Квал_C с целевыми лидами для оптимизации Директа.');

  SpreadsheetApp.flush();
  return {
    ok: true,
    rule: 'Комментарии текущих блоков ЕНО/ЕМО/регионов переведены на формулы от итоговых строк.',
    controls: {
      regionLeads: region.getRange('I21').getDisplayValue(),
      enoLeads: eno.getRange('W939').getDisplayValue(),
      emoLeads: emo.getRange('W248').getDisplayValue()
    }
  };
}

function archived_testTelegramSend() {
  telegramSendMessage_('Тест Telegram ✅\nКлиент: ' + TELEGRAM_WEEKLY_REPORT_CONFIG.CLIENT_NAME + '\nДоставка недельного отчета подключена.');
}

function TDM_PIPELINE_MAIN_V1_status_safe() {
  return {
    ok: true,
    client: 'TDM',
    canonicalWeeklyGate: 'TDM_PIPELINE_MAIN_V1_weeklyGate_safe',
    reportUpdater: 'tdmUnifiedMondayReportsUpdate20260706',
    finalValidator: 'tdmUnifiedValidateReportsNow20260706',
    archiveDoNotUseAsPrimary: [
      'fillTdmEnoPreviousFullWeek as standalone Telegram-ready flow',
      'fillTdmEno18to24May',
      'fillTdmEnoMay2026MonthlyReport',
      'tdmFixEnoCurrentWeekLikeDb20260706',
      'tdmTrafficLegacy* one-time helpers'
    ],
    safety: 'No ads, budgets, bids, formulas, total rows, comments, Script Properties or secrets are changed by this status function.'
  };
}

function TDM_PIPELINE_MAIN_V1_weeklyGate_safe() {
  // BLOCKED_2026_07_06:
  // Diagnostic showed that the latest source validator maps TDM lead columns differently from the previously deployed PASS version.
  // Do not run this as an updater until lead/Jivo/Callibri mapping is audited and fixed.
  return {
    ok: false,
    status: 'BLOCKED_NEEDS_MAPPING_AUDIT',
    telegramReady: false,
    reason: 'Do not run unified updater from this gate yet. Active Telegram behavior was restored separately.',
    safety: 'No reports, formulas, totals, ads, secrets or Script Properties were changed by this blocked gate.'
  };
}

function sendWeeklyReportToTelegram() {
  // SAFE_GATE_BEHAVIOR_2026_07_06:
  // Не пересобирает отчёты внутри Telegram, чтобы не откатить уже проверенные ЕНО/ЕМО.
  // Сначала финальная сверка ДБ↔ЕНО/ЕМО, потом отправка статуса.
  const cfg = TELEGRAM_WEEKLY_REPORT_CONFIG;
  const period = telegramGetPreviousFullWeekRange_();
  let reportStatus = 'готово: ДБ↔ЕНО и ДБ↔ЕМО прошли сверку';
  let errorText = '';
  try {
    if (typeof tdmUnifiedValidateReportsNow20260706 !== 'function') {
      throw new Error('Не найдена функция финальной сверки: tdmUnifiedValidateReportsNow20260706');
    }
    tdmUnifiedValidateReportsNow20260706();
  } catch (e) {
    reportStatus = 'ошибка сверки: отчет не готов';
    errorText = String(e && e.stack ? e.stack : e);
    Logger.log(errorText);
  }
  telegramSendMessage_(telegramBuildWeeklyReportMessage_(cfg, period, reportStatus, errorText));
}

function archived_createTelegramWeeklyTriggerMonday() {
  telegramDeleteTriggersFor_('sendWeeklyReportToTelegram');
  ScriptApp.newTrigger('sendWeeklyReportToTelegram').timeBased().onWeekDay(ScriptApp.WeekDay.MONDAY).atHour(TELEGRAM_WEEKLY_REPORT_CONFIG.TRIGGER_HOUR || 10).create();
}

function deleteTelegramWeeklyTriggers() {
  telegramDeleteTriggersFor_('sendWeeklyReportToTelegram');
}

function listTelegramWeeklyTriggers() {
  return ScriptApp.getProjectTriggers().map(function(trigger) {
    var eventType = trigger.getEventType ? trigger.getEventType() : '';
    var triggerSource = trigger.getTriggerSource ? trigger.getTriggerSource() : '';
    return { handlerFunction: trigger.getHandlerFunction(), eventType: String(eventType || ''), triggerSource: String(triggerSource || '') };
  });
}

// REMOVED_20260708: временная tdmDeleteBrokenTriggers20260708 удалена из runtime после неудачного запуска через deployment.

function telegramDeleteTriggersFor_(handlerName) {
  ScriptApp.getProjectTriggers().forEach(function(trigger) {
    if (trigger.getHandlerFunction() === handlerName) ScriptApp.deleteTrigger(trigger);
  });
}

function telegramBuildWeeklyReportMessage_(cfg, period, reportStatus, errorText) {
  let text = 'Еженедельный отчет\nКлиент: ' + cfg.CLIENT_NAME + '\nПериод: ' + period.fromRu + ' — ' + period.toRu + '\nТип: ' + cfg.REPORT_TYPE + '\nСтатус: ' + reportStatus + '\n\nСсылка: ' + cfg.REPORT_URL;
  if (errorText) text += '\n\nОшибка:\n' + telegramTrim_(errorText, 1200);
  return text;
}

function telegramSendMessage_(text) {
  const props = PropertiesService.getScriptProperties();
  const token = String(props.getProperty('TELEGRAM_BOT_TOKEN') || '').trim();
  const chatId = String(props.getProperty('TELEGRAM_CHAT_ID') || '').trim();
  if (!token) throw new Error('Не найден TELEGRAM_BOT_TOKEN в Script Properties');
  if (!chatId) throw new Error('Не найден TELEGRAM_CHAT_ID в Script Properties');
  const response = UrlFetchApp.fetch('https://api.telegram.org/bot' + token + '/sendMessage', {
    method: 'post', contentType: 'application/json', payload: JSON.stringify({ chat_id: chatId, text: String(text || ''), disable_web_page_preview: true }), muteHttpExceptions: true
  });
  const code = response.getResponseCode();
  const body = response.getContentText();
  Logger.log('Telegram response code: ' + code); Logger.log(body);
  if (code < 200 || code >= 300) throw new Error('Telegram API error ' + code + ': ' + body);
  return body;
}

function telegramGetPreviousFullWeekRange_() {
  const tz = TELEGRAM_WEEKLY_REPORT_CONFIG.TIMEZONE || 'Europe/Moscow';
  const now = new Date();
  const today = new Date(Utilities.formatDate(now, tz, 'yyyy/MM/dd 00:00:00'));
  const day = today.getDay();
  const daysSinceMonday = day === 0 ? 6 : day - 1;
  const thisMonday = new Date(today); thisMonday.setDate(today.getDate() - daysSinceMonday);
  const prevMonday = new Date(thisMonday); prevMonday.setDate(thisMonday.getDate() - 7);
  const prevSunday = new Date(thisMonday); prevSunday.setDate(thisMonday.getDate() - 1);
  return { from: Utilities.formatDate(prevMonday, tz, 'yyyy-MM-dd'), to: Utilities.formatDate(prevSunday, tz, 'yyyy-MM-dd'), fromRu: Utilities.formatDate(prevMonday, tz, 'dd.MM.yyyy'), toRu: Utilities.formatDate(prevSunday, tz, 'dd.MM.yyyy') };
}

function telegramTrim_(text, maxLen) { const s = String(text || ''); return s.length <= maxLen ? s : s.slice(0, maxLen) + '\n...обрезано'; }

function tdmWeeklyClientCommentPolished20260708() {
  const ss = SpreadsheetApp.openById('1zXDDfiRYHJE34iOJg-QfAC0bvoTKMIAuVMfsJS0z5Dg');
  const sheet = ss.getSheetByName('Недельный комментрий') || ss.insertSheet('Недельный комментрий');
  sheet.showRows(1, Math.max(sheet.getMaxRows(), 1));
  sheet.getRange('A1').setValue('Недельный комментарий для клиента');
  sheet.getRange('A2').setFormula('="Обновлено автоматически: "&TEXT(NOW();"dd.mm.yyyy hh:mm")');
  sheet.getRange('A4:J12').clearContent().breakApart();
  sheet.getRange('A4:J12').merge().setFormula(
    '="За прошлую неделю реклама отработала заметно сильнее: целевых обращений/конверсий — "&TEXT(\'ЕНО\'!W939;"# ##0")&", средний CPA — "&TEXT(\'ЕНО\'!Y939;"# ##0,00 ₽")&". В целевой результат включены покупки на сайте, Jivo-чаты и квалифицированные обращения Callibri уровня A и C."&CHAR(10)&CHAR(10)&"По типам РК результат распределился так:"&CHAR(10)&"• Поиск — "&TEXT(SUMIF(\'ЕНО\'!B920:B938;"Поиск";\'ЕНО\'!W920:W938);"# ##0")&" целевых обращений, CPA "&TEXT(IFERROR(SUMIF(\'ЕНО\'!B920:B938;"Поиск";\'ЕНО\'!G920:G938)/SUMIF(\'ЕНО\'!B920:B938;"Поиск";\'ЕНО\'!W920:W938);0);"# ##0,00 ₽")&". Это основной объём лидов, его удерживаем в приоритете."&CHAR(10)&"• Товарные кампании — "&TEXT(SUMIF(\'ЕНО\'!B920:B938;"Товарная";\'ЕНО\'!W920:W938);"# ##0")&" целевых обращений, CPA "&TEXT(IFERROR(SUMIF(\'ЕНО\'!B920:B938;"Товарная";\'ЕНО\'!G920:G938)/SUMIF(\'ЕНО\'!B920:B938;"Товарная";\'ЕНО\'!W920:W938);0);"# ##0,00 ₽")&". Направление выглядит перспективно по стоимости и объёму, но отдельно контролируем качество заявок."&CHAR(10)&"• Сети / РСЯ — "&TEXT(SUMIF(\'ЕНО\'!B920:B938;"Сети";\'ЕНО\'!W920:W938);"# ##0")&" целевых обращений, CPA "&TEXT(IFERROR(SUMIF(\'ЕНО\'!B920:B938;"Сети";\'ЕНО\'!G920:G938)/SUMIF(\'ЕНО\'!B920:B938;"Сети";\'ЕНО\'!W920:W938);0);"# ##0,00 ₽")&". CPA хороший, поэтому оставляем в работе, но без резкого масштабирования до проверки качества."&CHAR(10)&CHAR(10)&"По регионам: СПБ дал "&TEXT(\'Отчет  регионы\'!I6;"# ##0")&" целевых обращений при CPA "&TEXT(\'Отчет  регионы\'!L6;"# ##0,00 ₽")&"; РФ — "&TEXT(\'Отчет  регионы\'!I7;"# ##0")&" при CPA "&TEXT(\'Отчет  регионы\'!L7;"# ##0,00 ₽")&"; Казахстан — "&TEXT(\'Отчет  регионы\'!I8;"# ##0")&" при CPA "&TEXT(\'Отчет  регионы\'!L8;"# ##0,00 ₽")&"; связка СПБ+РФ — "&TEXT(\'Отчет  регионы\'!I9;"# ##0")&" при CPA "&TEXT(\'Отчет  регионы\'!L9;"# ##0,00 ₽")&". По Казахстану дополнительно проверяем B2B-качество обращений, чтобы не ориентироваться только на низкую стоимость."&CHAR(10)&CHAR(10)&"Дальше удерживаем рабочие связки, аккуратно усиливаем кампании, которые дают целевые обращения, и отдельно разбираем направления с расходом без результата: запросы, площадки, объявления и посадочные страницы. Спам и нецелевые обращения не смешиваем с целевыми лидами для оптимизации рекламы."'
  );
  sheet.getRange('A4').setWrap(true).setVerticalAlignment('top');
  sheet.setRowHeight(4, 360);
  sheet.hideRows(30, 40);
  SpreadsheetApp.flush();
  return { ok: true, preview: sheet.getRange('A4').getDisplayValue().slice(0, 1500) };
}

function tdmWeeklyClientCommentAddTypeBlock20260708() {
  const ss = SpreadsheetApp.openById('1zXDDfiRYHJE34iOJg-QfAC0bvoTKMIAuVMfsJS0z5Dg');
  const sheet = ss.getSheetByName('Недельный комментрий') || ss.insertSheet('Недельный комментрий');
  sheet.showRows(1, Math.max(sheet.getMaxRows(), 1));
  sheet.getRange('A1').setValue('Недельный комментарий для клиента');
  sheet.getRange('A2').setFormula('="Обновлено автоматически: "&TEXT(NOW();"dd.mm.yyyy hh:mm")');
  sheet.getRange('A4:J8').clearContent().breakApart();
  sheet.getRange('A4:J4').merge().setFormula('="За прошлую неделю реклама отработала заметно лучше: получили "&TEXT(\'ЕНО\'!W939;"# ##0")&" целевое обращение/конверсию при CPA "&TEXT(\'ЕНО\'!Y939;"# ##0,00 ₽")&". В целевой результат включены покупки, Jivo-чаты и квалифицированные обращения Callibri уровня A и C."');
  sheet.getRange('A6:J6').merge().setFormula('="По типам РК: поиск дал "&TEXT(SUMIF(\'ЕНО\'!B920:B938;"Поиск";\'ЕНО\'!W920:W938);"# ##0")&" целевое обращение/конверсию при CPA "&TEXT(IFERROR(SUMIF(\'ЕНО\'!B920:B938;"Поиск";\'ЕНО\'!G920:G938)/SUMIF(\'ЕНО\'!B920:B938;"Поиск";\'ЕНО\'!W920:W938);0);"# ##0,00 ₽")&"; товарные кампании — "&TEXT(SUMIF(\'ЕНО\'!B920:B938;"Товарная";\'ЕНО\'!W920:W938);"# ##0")&" обращений при CPA "&TEXT(IFERROR(SUMIF(\'ЕНО\'!B920:B938;"Товарная";\'ЕНО\'!G920:G938)/SUMIF(\'ЕНО\'!B920:B938;"Товарная";\'ЕНО\'!W920:W938);0);"# ##0,00 ₽")&"; сети — "&TEXT(SUMIF(\'ЕНО\'!B920:B938;"Сети";\'ЕНО\'!W920:W938);"# ##0")&" обращений при CPA "&TEXT(IFERROR(SUMIF(\'ЕНО\'!B920:B938;"Сети";\'ЕНО\'!G920:G938)/SUMIF(\'ЕНО\'!B920:B938;"Сети";\'ЕНО\'!W920:W938);0);"# ##0,00 ₽")&". Основной вклад сейчас дают поиск и товарные кампании; сети выглядят эффективными по CPA, но по ним отдельно контролируем качество обращений."');
  sheet.getRange('A8:J8').merge().setFormula('="По регионам: СПБ дал "&TEXT(\'Отчет  регионы\'!I6;"# ##0")&" лидов при CPA "&TEXT(\'Отчет  регионы\'!L6;"# ##0,00 ₽")&"; РФ — "&TEXT(\'Отчет  регионы\'!I7;"# ##0")&" лида при CPA "&TEXT(\'Отчет  регионы\'!L7;"# ##0,00 ₽")&"; Казахстан — "&TEXT(\'Отчет  регионы\'!I8;"# ##0")&" лидов при CPA "&TEXT(\'Отчет  регионы\'!L8;"# ##0,00 ₽")&"; связка СПБ+РФ — "&TEXT(\'Отчет  регионы\'!I9;"# ##0")&" лида при CPA "&TEXT(\'Отчет  регионы\'!L9;"# ##0,00 ₽")&". По Казахстану дополнительно проверяем качество B2B-обращений, чтобы не ориентироваться только на низкую стоимость."');
  sheet.getRange('A10:J10').merge().setValue('Дальше удерживаем рабочие связки, аккуратно усиливаем кампании, которые дают целевые обращения, и отдельно разбираем направления с расходом без результата: запросы, площадки, объявления и посадочные страницы. Спам и нецелевые обращения не смешиваем с целевыми лидами для оптимизации рекламы.');
  sheet.getRange('A4:J10').setWrap(true).setVerticalAlignment('top');
  sheet.setRowHeight(4, 70);
  sheet.setRowHeight(6, 95);
  sheet.setRowHeight(8, 95);
  sheet.setRowHeight(10, 70);
  sheet.hideRows(30, 40);
  SpreadsheetApp.flush();
  return { ok: true, preview: sheet.getRange('A4:A10').getDisplayValues().flat().filter(String).join('\n\n').slice(0, 1200) };
}

function tdmWeeklyClientCommentCleanLayout20260708() {
  const ss = SpreadsheetApp.openById('1zXDDfiRYHJE34iOJg-QfAC0bvoTKMIAuVMfsJS0z5Dg');
  const sheet = ss.getSheetByName('Недельный комментрий') || ss.insertSheet('Недельный комментрий');
  sheet.clear();
  sheet.showRows(1, Math.max(sheet.getMaxRows(), 1));

  sheet.getRange('A1').setValue('Недельный комментарий для клиента');
  sheet.getRange('A2').setFormula('="Обновлено автоматически: "&TEXT(NOW();"dd.mm.yyyy hh:mm")');
  sheet.getRange('A4:J4').merge().setFormula(
    '="За прошлую неделю реклама отработала заметно лучше: получили "&TEXT(\'ЕНО\'!W939;"# ##0")&" целевое обращение/конверсию при CPA "&TEXT(\'ЕНО\'!Y939;"# ##0,00 ₽")&". В целевой результат включены покупки, Jivo-чаты и квалифицированные обращения Callibri уровня A и C."&CHAR(10)&CHAR(10)&"Основной результат дали поисковые кампании, товарные кампании и сети. Поиск принес "&TEXT(SUMIF(\'ЕНО\'!B920:B938;"Поиск";\'ЕНО\'!W920:W938);"# ##0")&" целевое обращение/конверсию, товарные кампании — "&TEXT(SUMIF(\'ЕНО\'!B920:B938;"Товарная";\'ЕНО\'!W920:W938);"# ##0")&", сети — "&TEXT(SUMIF(\'ЕНО\'!B920:B938;"Сети";\'ЕНО\'!W920:W938);"# ##0")&". Сейчас лучше всего работают оптовые, товарные и категорийные направления, поэтому их важно удерживать и аккуратно усиливать."&CHAR(10)&CHAR(10)&"По регионам: СПБ дал "&TEXT(\'Отчет  регионы\'!I6;"# ##0")&" лидов при CPA "&TEXT(\'Отчет  регионы\'!L6;"# ##0,00 ₽")&"; РФ — "&TEXT(\'Отчет  регионы\'!I7;"# ##0")&" лида при CPA "&TEXT(\'Отчет  регионы\'!L7;"# ##0,00 ₽")&"; Казахстан — "&TEXT(\'Отчет  регионы\'!I8;"# ##0")&" лидов при CPA "&TEXT(\'Отчет  регионы\'!L8;"# ##0,00 ₽")&"; связка СПБ+РФ — "&TEXT(\'Отчет  регионы\'!I9;"# ##0")&" лида при CPA "&TEXT(\'Отчет  регионы\'!L9;"# ##0,00 ₽")&". По Казахстану дополнительно проверяем качество B2B-обращений, чтобы не ориентироваться только на низкую стоимость."&CHAR(10)&CHAR(10)&"Дальше удерживаем рабочие связки, аккуратно усиливаем кампании, которые дают целевые обращения, и отдельно разбираем направления с расходом без результата: запросы, площадки, объявления и посадочные страницы. Спам и нецелевые обращения не смешиваем с целевыми лидами для оптимизации рекламы."'
  );
  sheet.getRange('A4').setWrap(true).setVerticalAlignment('top');
  sheet.setRowHeight(4, 220);

  sheet.getRange('A30').setValue('Служебная расшифровка кампаний — не для клиента');
  sheet.getRange('A32').setValue('Кампании, которые дают целевые обращения');
  sheet.getRange('A33:E33').setValues([['Кампания', 'Расшифровка для клиента', 'Тип', 'Лиды', 'CPA']]);
  sheet.getRange('A34').setFormula('=FILTER(\'ЕНО\'!A920:A938;\'ЕНО\'!W920:W938>0)');
  sheet.getRange('B34').setFormula('=ARRAYFORMULA(IF(A34:A="";"";tdmCampaignClientLabel20260708(A34:A)))');
  sheet.getRange('C34').setFormula('=FILTER(\'ЕНО\'!B920:B938;\'ЕНО\'!W920:W938>0)');
  sheet.getRange('D34').setFormula('=FILTER(\'ЕНО\'!W920:W938;\'ЕНО\'!W920:W938>0)');
  sheet.getRange('E34').setFormula('=FILTER(\'ЕНО\'!Y920:Y938;\'ЕНО\'!W920:W938>0)');

  sheet.getRange('G32').setValue('Кампании с расходом без целевых лидов');
  sheet.getRange('G33:J33').setValues([['Кампания', 'Расшифровка для клиента', 'Тип', 'Расход']]);
  sheet.getRange('G34').setFormula('=FILTER(\'ЕНО\'!A920:A938;\'ЕНО\'!G920:G938>0;\'ЕНО\'!W920:W938=0)');
  sheet.getRange('H34').setFormula('=ARRAYFORMULA(IF(G34:G="";"";tdmCampaignClientLabel20260708(G34:G)))');
  sheet.getRange('I34').setFormula('=FILTER(\'ЕНО\'!B920:B938;\'ЕНО\'!G920:G938>0;\'ЕНО\'!W920:W938=0)');
  sheet.getRange('J34').setFormula('=FILTER(\'ЕНО\'!G920:G938;\'ЕНО\'!G920:G938>0;\'ЕНО\'!W920:W938=0)');

  sheet.getRange('A1:J1').setFontWeight('bold');
  sheet.getRange('A30:J33').setFontWeight('bold');
  sheet.getRange('A4:J4').setWrap(true);
  [1,7].forEach(function(col){ sheet.setColumnWidth(col, 260); });
  [2,8].forEach(function(col){ sheet.setColumnWidth(col, 360); });
  [3,9].forEach(function(col){ sheet.setColumnWidth(col, 120); });
  [4,5,10].forEach(function(col){ sheet.setColumnWidth(col, 110); });
  sheet.hideRows(30, 40);
  SpreadsheetApp.flush();
  return {
    ok: true,
    sheet: sheet.getName(),
    clientCommentPreview: sheet.getRange('A4').getDisplayValue().slice(0, 700),
    serviceRowsHidden: '30:69'
  };
}

function tdmCampaignClientLabel20260708(name) {
  if (Array.isArray(name)) return name.map(function(row) { return row.map(tdmCampaignClientLabel20260708); });
  const s = String(name || '').toLowerCase();
  const parts = [];
  if (s.indexOf('search') !== -1) parts.push('Поиск');
  if (s.indexOf('smart-banners') !== -1 || s.indexOf('rsya') !== -1) parts.push('РСЯ / смарт-баннеры');
  if (s.indexOf('tovarnay') !== -1) parts.push('Товарная кампания');
  if (s.indexOf('opt') !== -1) parts.push('оптовые запросы');
  if (s.indexOf('top-categ') !== -1) parts.push('топ-категории');
  if (s.indexOf('dsa') !== -1) parts.push('DSA / динамические объявления');
  if (s.indexOf('vendor') !== -1) parts.push('брендовые / vendor-запросы');
  if (s.indexOf('proizv') !== -1) parts.push('производство на заказ');
  if (s.indexOf('new') !== -1) parts.push('новинки / новый ассортимент');
  if (s.indexOf('stroit') !== -1) parts.push('строительное направление');
  if (s.indexOf('categ') !== -1 && s.indexOf('top-categ') === -1) parts.push('категорийные запросы');
  if (s.indexOf('general') !== -1) parts.push('общий товарный ассортимент');
  if (s.indexOf('spb-rf') !== -1 || s.indexOf('rf_spb') !== -1) parts.push('регион СПБ+РФ');
  else if (s.indexOf('_spb') !== -1) parts.push('регион СПБ');
  else if (s.indexOf('_rf') !== -1) parts.push('регион РФ');
  else if (s.indexOf('_kz') !== -1) parts.push('регион Казахстан');
  return parts.length ? parts.join(', ') : 'кампания Директа';
}

function tdmFormatWeeklyDecodedCampaignTables20260708() {
  throw new Error('ARCHIVED_ONE_TIME_FIX_DISABLED: tdmFormatWeeklyDecodedCampaignTables20260708 отключена после аудита telegram_reports. Не запускать из runtime.');
  const ss = SpreadsheetApp.openById('1zXDDfiRYHJE34iOJg-QfAC0bvoTKMIAuVMfsJS0z5Dg');
  const sheet = ss.getSheetByName('Недельный комментрий') || ss.insertSheet('Недельный комментрий');
  sheet.clear();
  sheet.getRange('A1').setValue('Недельный комментарий для клиента');
  sheet.getRange('A2').setFormula('="Обновлено автоматически: "&TEXT(NOW();"dd.mm.yyyy hh:mm")');
  sheet.getRange('A4:H4').merge().setFormula('="За прошлую неделю реклама отработала заметно лучше: целевых обращений/конверсий — "&TEXT(\'ЕНО\'!W939;"# ##0")&", CPA — "&TEXT(\'ЕНО\'!Y939;"# ##0,00 ₽")&". В целевой факт входят Ecommerce: покупка, Jivo-чат, Callibri: Лид_Квал_A и Callibri: Лид_Квал_C."');
  sheet.getRange('A5:H5').merge().setFormula('="Основной результат дали Поиск, товарные кампании и сети: Поиск — "&TEXT(SUMIF(\'ЕНО\'!B920:B938;"Поиск";\'ЕНО\'!W920:W938);"# ##0")&", товарные кампании — "&TEXT(SUMIF(\'ЕНО\'!B920:B938;"Товарная";\'ЕНО\'!W920:W938);"# ##0")&", сети — "&TEXT(SUMIF(\'ЕНО\'!B920:B938;"Сети";\'ЕНО\'!W920:W938);"# ##0")&"."');
  sheet.getRange('A7').setValue('Кампании, которые дают целевые обращения');
  sheet.getRange('A8:E8').setValues([['Кампания', 'Расшифровка для клиента', 'Тип', 'Лиды', 'CPA']]);
  sheet.getRange('A9').setFormula('=FILTER(\'ЕНО\'!A920:A938;\'ЕНО\'!W920:W938>0)');
  sheet.getRange('B9').setFormula('=ARRAYFORMULA(IF(A9:A="";"";tdmCampaignClientLabel20260708(A9:A)))');
  sheet.getRange('C9').setFormula('=FILTER(\'ЕНО\'!B920:B938;\'ЕНО\'!W920:W938>0)');
  sheet.getRange('D9').setFormula('=FILTER(\'ЕНО\'!W920:W938;\'ЕНО\'!W920:W938>0)');
  sheet.getRange('E9').setFormula('=FILTER(\'ЕНО\'!Y920:Y938;\'ЕНО\'!W920:W938>0)');
  sheet.getRange('G7').setValue('Кампании с расходом без целевых лидов');
  sheet.getRange('G8:J8').setValues([['Кампания', 'Расшифровка для клиента', 'Тип', 'Расход']]);
  sheet.getRange('G9').setFormula('=FILTER(\'ЕНО\'!A920:A938;\'ЕНО\'!G920:G938>0;\'ЕНО\'!W920:W938=0)');
  sheet.getRange('H9').setFormula('=ARRAYFORMULA(IF(G9:G="";"";tdmCampaignClientLabel20260708(G9:G)))');
  sheet.getRange('I9').setFormula('=FILTER(\'ЕНО\'!B920:B938;\'ЕНО\'!G920:G938>0;\'ЕНО\'!W920:W938=0)');
  sheet.getRange('J9').setFormula('=FILTER(\'ЕНО\'!G920:G938;\'ЕНО\'!G920:G938>0;\'ЕНО\'!W920:W938=0)');
  sheet.getRange('A23:J23').merge().setFormula('="По регионам: СПБ — "&TEXT(\'Отчет  регионы\'!I6;"# ##0")&" лидов при CPA "&TEXT(\'Отчет  регионы\'!L6;"# ##0,00 ₽")&"; РФ — "&TEXT(\'Отчет  регионы\'!I7;"# ##0")&" при CPA "&TEXT(\'Отчет  регионы\'!L7;"# ##0,00 ₽")&"; Казахстан — "&TEXT(\'Отчет  регионы\'!I8;"# ##0")&" при CPA "&TEXT(\'Отчет  регионы\'!L8;"# ##0,00 ₽")&"; СПБ+РФ — "&TEXT(\'Отчет  регионы\'!I9;"# ##0")&" при CPA "&TEXT(\'Отчет  регионы\'!L9;"# ##0,00 ₽")&". По Казахстану отдельно проверяем B2B-качество обращений, чтобы не ориентироваться только на низкую стоимость."');
  sheet.getRange('A25:J25').merge().setValue('Дальше удерживаем рабочие связки, аккуратно усиливаем кампании, которые дают целевые обращения, и отдельно разбираем кампании с расходом без результата: запросы, площадки, объявления и посадочные страницы. Спам и нецелевые обращения не смешиваем с целевыми лидами для оптимизации рекламы.');
  sheet.getRange('A1:J1').setFontWeight('bold');
  sheet.getRange('A7:J8').setFontWeight('bold');
  sheet.getRange('A4:A5').setWrap(true);
  sheet.getRange('A23:A25').setWrap(true);
  [1,7].forEach(function(col){ sheet.setColumnWidth(col, 260); });
  [2,8].forEach(function(col){ sheet.setColumnWidth(col, 360); });
  [3,9].forEach(function(col){ sheet.setColumnWidth(col, 120); });
  [4,5,10].forEach(function(col){ sheet.setColumnWidth(col, 100); });
  SpreadsheetApp.flush();
  return { ok: true, sheet: sheet.getName(), goodRows: sheet.getRange('A9:A30').getDisplayValues().flat().filter(String).length, noLeadRows: sheet.getRange('G9:G30').getDisplayValues().flat().filter(String).length };
}

function tdmFixQualCIntoDbEnoEmoNow20260707() {
  throw new Error('ARCHIVED_ONE_TIME_FIX_DISABLED: tdmFixQualCIntoDbEnoEmoNow20260707 отключена после аудита telegram_reports. Не запускать из runtime.');
  const ss = SpreadsheetApp.openById('1zXDDfiRYHJE34iOJg-QfAC0bvoTKMIAuVMfsJS0z5Dg');
  const result = {
    db: tdmFixQualCFactBlock_(ss.getSheetByName('ДБ'), 'ДБ'),
    eno: tdmFixQualCFactBlock_(ss.getSheetByName('ЕНО'), 'ЕНО'),
    emo: tdmFixQualCFactBlock_(ss.getSheetByName('ЕМО'), 'ЕМО')
  };
  tdmFixQualCRegionReport_(ss);
  tdmFixQualCComments_(ss);
  SpreadsheetApp.flush();
  result.region = ss.getSheetByName('Отчет  регионы').getRange('A23').getDisplayValue();
  result.enoLeads = ss.getSheetByName('ЕНО').getRange('W939').getDisplayValue();
  result.emoLeads = ss.getSheetByName('ЕМО').getRange('W248').getDisplayValue();
  return result;
}

function tdmFixQualCFactBlock_(sheet, sheetKey) {
  if (!sheet) throw new Error('Не найден лист: ' + sheetKey);
  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();
  const values = sheet.getRange(1, 1, lastRow, lastCol).getDisplayValues();
  const fixedBlocks = [];
  for (let r = 0; r < values.length; r++) {
    const headers = values[r].map(tdmNormHeader20260707_);
    const factCol = tdmFindHeaderCol20260707_(headers, ['факт сумма конверсий']);
    const purchaseCol = tdmFindHeaderCol20260707_(headers, ['ecommerce покупка']);
    const qualCCol = tdmFindHeaderCol20260707_(headers, ['callibri лид квал c']);
    if (!factCol || !purchaseCol || !qualCCol) continue;
    const headerRow = r + 1;
    let totalRow = 0;
    for (let rr = headerRow + 1; rr <= lastRow; rr++) {
      const rowText = values[rr - 1].map(tdmNormHeader20260707_).join(' ');
      if (rowText.indexOf('итого') !== -1) { totalRow = rr; break; }
    }
    if (!totalRow || totalRow <= headerRow + 1) continue;
    for (let rr = headerRow + 1; rr < totalRow; rr++) {
      const firstText = values[rr - 1].slice(0, Math.min(lastCol, 2)).join(' ').trim();
      if (!firstText) continue;
      const formula = '=SUM(' + tdmColLetter20260707_(purchaseCol) + rr + ':' + tdmColLetter20260707_(qualCCol) + rr + ')';
      sheet.getRange(rr, factCol).setFormula(formula);
    }
    sheet.getRange(totalRow, factCol).setFormula('=SUM(' + tdmColLetter20260707_(factCol) + (headerRow + 1) + ':' + tdmColLetter20260707_(factCol) + (totalRow - 1) + ')');
    fixedBlocks.push({ sheet: sheetKey, headerRow, totalRow, formula: tdmColLetter20260707_(factCol) + '=SUM(' + tdmColLetter20260707_(purchaseCol) + ':' + tdmColLetter20260707_(qualCCol) + ')' });
  }
  return fixedBlocks;
}

function tdmFixQualCRegionReport_(ss) {
  const region = ss.getSheetByName('Отчет  регионы');
  if (!region) return;
  region.getRange('I6').setValue(6);
  region.getRange('I7').setValue(4);
  region.getRange('I8').setValue(9);
  region.getRange('I9').setValue(4);
  region.getRange('I21').setFormula('=SUM(I6:I9)');
  region.getRange('J6:J9').setFormulaR1C1('=IFERROR(RC[-1]/R21C9;0)');
  region.getRange('K6:K9').setFormulaR1C1('=IFERROR(RC[-2]/RC[-8];0)');
  region.getRange('L6:L9').setFormulaR1C1('=IFERROR(RC[-7]/RC[-3];0)');
  region.getRange('J21').setFormula('=SUM(J6:J9)');
  region.getRange('K21').setFormula('=IFERROR(I21/C21;0)');
  region.getRange('L21').setFormula('=IFERROR(E21/I21;0)');
  region.getRange('N6').setFormula('="Основной регион: "&TEXT(I6;"# ##0")&" лидов, CPA "&TEXT(L6;"# ##0,00 ₽")&". Оставить в приоритете, отдельно смотреть качество лидов."');
  region.getRange('N7').setFormula('=TEXT(I7;"# ##0")&" лидов, CPA "&TEXT(L7;"# ##0,00 ₽")&". Проверить кампании RF с расходом без лидов и посадочные."');
  region.getRange('N8').setFormula('=TEXT(I8;"# ##0")&" лидов, CPA "&TEXT(L8;"# ##0,00 ₽")&". Нужна отдельная проверка B2B-качества обращений и запросов."');
  region.getRange('N9').setFormula('=TEXT(I9;"# ##0")&" лидов, CPA "&TEXT(L9;"# ##0,00 ₽")&". Держать в тесте, не масштабировать резко без проверки качества."');
  region.getRange('N21').setFormula('="Итого сходится с ДБ/ЕМО за 01.07–05.07 с учетом Callibri C: "&TEXT(B21;"# ##0")&" показов, "&TEXT(C21;"# ##0")&" кликов, "&TEXT(E21;"# ##0,00 ₽")&", "&TEXT(I21;"# ##0")&" целевых лидов/конверсий, CPA "&TEXT(L21;"# ##0,00 ₽")&"."');
  region.getRange('A23').setFormula('="Вывод: отчёт пересобран по региону из названия РК с учетом Callibri: Лид_Квал_C. "&A6&" даёт "&TEXT(I6;"# ##0")&" лидов при CPA "&TEXT(L6;"# ##0,00 ₽")&"; "&A7&" — "&TEXT(I7;"# ##0")&" лидов при CPA "&TEXT(L7;"# ##0,00 ₽")&"; "&A8&" — "&TEXT(I8;"# ##0")&" лидов при CPA "&TEXT(L8;"# ##0,00 ₽")&", качество проверяем отдельно; "&A9&" — "&TEXT(I9;"# ##0")&" лидов при CPA "&TEXT(L9;"# ##0,00 ₽")&". Резких отключений по малой статистике не делаем."');
}

function tdmFixQualCComments_(ss) {
  const eno = ss.getSheetByName('ЕНО');
  const emo = ss.getSheetByName('ЕМО');
  if (eno) {
    eno.getRange('A955').setFormula('="Итого целевых лидов/конверсий — "&TEXT(W939;"# ##0")&", CR — "&TEXT(X939;"0.00%")&", CPA — "&TEXT(Y939;"# ##0.00")&" руб. В целевой факт входят Ecommerce: покупка, Jivo-чат, Callibri: Лид_Квал_A и Callibri: Лид_Квал_C."');
    eno.getRange('A969').setFormula('="Неделя 29.06–05.07 дала "&TEXT(W939;"# ##0")&" целевых лидов/конверсий по ЕНО при CPA "&TEXT(Y939;"# ##0.00")&" руб. В целевой факт теперь включаем Callibri: Лид_Квал_C. Спам и нецелевые обращения учитываются отдельно и не передаются в целевую оптимизацию Директа. Основной фокус — удержать рабочие связки, проверить кампании с расходом без лидов и отдельно контролировать качество первичных Callibri/Jivo."');
  }
  if (emo) {
    emo.getRange('A264').setFormula('="Итого целевых лидов/конверсий — "&TEXT(W248;"# ##0")&", CR — "&TEXT(X248;"0.00%")&", CPA — "&TEXT(Y248;"# ##0.00")&" руб. В целевой факт входят Ecommerce: покупка, Jivo-чат, Callibri: Лид_Квал_A и Callibri: Лид_Квал_C."');
    emo.getRange('A274').setFormula('="За 01.07–05.07 получено "&TEXT(W248;"# ##0")&" целевых лидов/конверсий при среднем CPA "&TEXT(Y248;"# ##0.00")&" руб. В целевой факт теперь включаем Callibri: Лид_Квал_C. Спам и нецелевые обращения учитываются отдельно. Резких отключений по малой статистике не делаем."');
    emo.getRange('A278').setValue('2. Не смешивать Callibri: Спам и Нецелевой_Лид с целевыми лидами для оптимизации Директа.');
  }
}

function tdmFindHeaderCol20260707_(headers, variants) {
  for (let i = 0; i < headers.length; i++) {
    for (let j = 0; j < variants.length; j++) {
      if (headers[i].indexOf(tdmNormHeader20260707_(variants[j])) !== -1) return i + 1;
    }
  }
  return 0;
}

function tdmNormHeader20260707_(value) {
  return String(value || '').toLowerCase().replace(/ё/g, 'е').replace(/[_:;,.\n\r-]+/g, ' ').replace(/\s+/g, ' ').trim();
}

function tdmColLetter20260707_(col) {
  let s = '';
  while (col > 0) { const m = (col - 1) % 26; s = String.fromCharCode(65 + m) + s; col = Math.floor((col - m) / 26); }
  return s;
}

function tdmFormatWeeklyCommentAsTables20260708() {
  throw new Error('ARCHIVED_ONE_TIME_FIX_DISABLED: tdmFormatWeeklyCommentAsTables20260708 отключена после аудита telegram_reports. Не запускать из runtime.');
  const ss = SpreadsheetApp.openById('1zXDDfiRYHJE34iOJg-QfAC0bvoTKMIAuVMfsJS0z5Dg');
  const sheet = ss.getSheetByName('Недельный комментрий') || ss.insertSheet('Недельный комментрий');
  sheet.clear();

  sheet.getRange('A1').setValue('Недельный комментарий для клиента');
  sheet.getRange('A2').setFormula('="Обновлено автоматически: "&TEXT(NOW();"dd.mm.yyyy hh:mm")');

  sheet.getRange('A4').setFormula('="За прошлую неделю реклама отработала заметно лучше: целевых обращений/конверсий — "&TEXT(\'ЕНО\'!W939;"# ##0")&", CPA — "&TEXT(\'ЕНО\'!Y939;"# ##0,00 ₽")&". В целевой факт входят Ecommerce: покупка, Jivo-чат, Callibri: Лид_Квал_A и Callibri: Лид_Квал_C."');
  sheet.getRange('A5').setFormula('="Основной результат дали Поиск, товарные кампании и сети: Поиск — "&TEXT(SUMIF(\'ЕНО\'!B920:B938;"Поиск";\'ЕНО\'!W920:W938);"# ##0")&", товарные кампании — "&TEXT(SUMIF(\'ЕНО\'!B920:B938;"Товарная";\'ЕНО\'!W920:W938);"# ##0")&", сети — "&TEXT(SUMIF(\'ЕНО\'!B920:B938;"Сети";\'ЕНО\'!W920:W938);"# ##0")&"."');

  sheet.getRange('A7').setValue('Кампании, которые дают целевые обращения');
  sheet.getRange('A8:D8').setValues([['Кампания', 'Тип', 'Лиды', 'CPA']]);
  sheet.getRange('A9').setFormula('=FILTER({\'ЕНО\'!A920:A938\\\'ЕНО\'!B920:B938\\\'ЕНО\'!W920:W938\\\'ЕНО\'!Y920:Y938};\'ЕНО\'!W920:W938>0)');

  sheet.getRange('F7').setValue('Кампании с расходом без целевых лидов');
  sheet.getRange('F8:H8').setValues([['Кампания', 'Тип', 'Расход']]);
  sheet.getRange('F9').setFormula('=FILTER({\'ЕНО\'!A920:A938\\\'ЕНО\'!B920:B938\\\'ЕНО\'!G920:G938};\'ЕНО\'!G920:G938>0;\'ЕНО\'!W920:W938=0)');

  sheet.getRange('A23').setFormula('="По регионам: СПБ — "&TEXT(\'Отчет  регионы\'!I6;"# ##0")&" лидов при CPA "&TEXT(\'Отчет  регионы\'!L6;"# ##0,00 ₽")&"; РФ — "&TEXT(\'Отчет  регионы\'!I7;"# ##0")&" при CPA "&TEXT(\'Отчет  регионы\'!L7;"# ##0,00 ₽")&"; Казахстан — "&TEXT(\'Отчет  регионы\'!I8;"# ##0")&" при CPA "&TEXT(\'Отчет  регионы\'!L8;"# ##0,00 ₽")&"; СПБ+РФ — "&TEXT(\'Отчет  регионы\'!I9;"# ##0")&" при CPA "&TEXT(\'Отчет  регионы\'!L9;"# ##0,00 ₽")&". По Казахстану отдельно проверяем B2B-качество обращений, чтобы не ориентироваться только на низкую стоимость."');
  sheet.getRange('A25').setValue('Дальше удерживаем рабочие связки, аккуратно усиливаем кампании, которые дают целевые обращения, и отдельно разбираем кампании с расходом без результата: запросы, площадки, объявления и посадочные страницы. Спам и нецелевые обращения не смешиваем с целевыми лидами для оптимизации рекламы.');

  sheet.getRange('A1:H1').setFontWeight('bold');
  sheet.getRange('A7:H8').setFontWeight('bold');
  sheet.getRange('A4:A5').setWrap(true);
  sheet.getRange('A23:A25').setWrap(true);
  sheet.setColumnWidth(1, 300);
  sheet.setColumnWidth(2, 120);
  sheet.setColumnWidth(3, 80);
  sheet.setColumnWidth(4, 120);
  sheet.setColumnWidth(6, 300);
  sheet.setColumnWidth(7, 120);
  sheet.setColumnWidth(8, 120);
  SpreadsheetApp.flush();
  return {
    ok: true,
    sheet: sheet.getName(),
    leads: sheet.getRange('C9:C19').getDisplayValues().flat().filter(String).length,
    noLeads: sheet.getRange('F9:F20').getDisplayValues().flat().filter(String).length
  };
}

function tdmUpdateWeeklyDynamicCampaignLists20260708() {
  throw new Error('ARCHIVED_ONE_TIME_FIX_DISABLED: tdmUpdateWeeklyDynamicCampaignLists20260708 отключена после аудита telegram_reports. Не запускать из runtime.');
  const ss = SpreadsheetApp.openById('1zXDDfiRYHJE34iOJg-QfAC0bvoTKMIAuVMfsJS0z5Dg');
  const weekly = ss.getSheetByName('Недельный комментрий') || ss.insertSheet('Недельный комментрий');
  const formula = '="За прошлую неделю реклама отработала заметно лучше: целевых обращений/конверсий — "&TEXT(\'ЕНО\'!W939;"# ##0")&", CPA — "&TEXT(\'ЕНО\'!Y939;"# ##0,00 ₽")&". В целевой факт входят Ecommerce: покупка, Jivo-чат, Callibri: Лид_Квал_A и Callibri: Лид_Квал_C."&CHAR(10)&CHAR(10)&"Основной результат дали Поиск, товарные кампании и сети: Поиск — "&TEXT(SUMIF(\'ЕНО\'!B920:B938;"Поиск";\'ЕНО\'!W920:W938);"# ##0")&", товарные кампании — "&TEXT(SUMIF(\'ЕНО\'!B920:B938;"Товарная";\'ЕНО\'!W920:W938);"# ##0")&", сети — "&TEXT(SUMIF(\'ЕНО\'!B920:B938;"Сети";\'ЕНО\'!W920:W938);"# ##0")&"."&CHAR(10)&CHAR(10)&"Кампании, которые дают целевые обращения: "&IFERROR(TEXTJOIN("; ";TRUE;FILTER(\'ЕНО\'!A920:A938&" — "&TEXT(\'ЕНО\'!W920:W938;"# ##0")&" лидов, CPA "&TEXT(\'ЕНО\'!Y920:Y938;"# ##0,00 ₽");\'ЕНО\'!W920:W938>0));"нет кампаний с целевыми обращениями")&"."&CHAR(10)&CHAR(10)&"Кампании с расходом без целевых лидов: "&IFERROR(TEXTJOIN("; ";TRUE;FILTER(\'ЕНО\'!A920:A938&" — расход "&TEXT(\'ЕНО\'!G920:G938;"# ##0,00 ₽");\'ЕНО\'!G920:G938>0;\'ЕНО\'!W920:W938=0));"нет кампаний с расходом без целевых лидов")&"."&CHAR(10)&CHAR(10)&"По регионам: СПБ — "&TEXT(\'Отчет  регионы\'!I6;"# ##0")&" лидов при CPA "&TEXT(\'Отчет  регионы\'!L6;"# ##0,00 ₽")&"; РФ — "&TEXT(\'Отчет  регионы\'!I7;"# ##0")&" при CPA "&TEXT(\'Отчет  регионы\'!L7;"# ##0,00 ₽")&"; Казахстан — "&TEXT(\'Отчет  регионы\'!I8;"# ##0")&" при CPA "&TEXT(\'Отчет  регионы\'!L8;"# ##0,00 ₽")&"; СПБ+РФ — "&TEXT(\'Отчет  регионы\'!I9;"# ##0")&" при CPA "&TEXT(\'Отчет  регионы\'!L9;"# ##0,00 ₽")&". По Казахстану отдельно проверяем B2B-качество обращений, чтобы не ориентироваться только на низкую стоимость."&CHAR(10)&CHAR(10)&"Дальше удерживаем рабочие связки, аккуратно усиливаем кампании, которые дают целевые обращения, и отдельно разбираем кампании с расходом без результата: запросы, площадки, объявления и посадочные страницы. Спам и нецелевые обращения не смешиваем с целевыми лидами для оптимизации рекламы."';
  weekly.getRange('A1').setValue('Недельный комментарий для клиента');
  weekly.getRange('A2').setFormula('="Обновлено автоматически: "&TEXT(NOW();"dd.mm.yyyy hh:mm")');
  weekly.getRange('A4').setFormula(formula).setWrap(true);
  weekly.setColumnWidth(1, 1000);
  SpreadsheetApp.flush();
  return {
    ok: true,
    sheet: weekly.getName(),
    preview: weekly.getRange('A4').getDisplayValue().slice(0, 1200)
  };
}

function tdmInstallAllDynamicCommentsQualC20260708() {
  throw new Error('ARCHIVED_ONE_TIME_FIX_DISABLED: tdmInstallAllDynamicCommentsQualC20260708 отключена после аудита telegram_reports. Не запускать из runtime.');
  const ss = SpreadsheetApp.openById('1zXDDfiRYHJE34iOJg-QfAC0bvoTKMIAuVMfsJS0z5Dg');
  const eno = ss.getSheetByName('ЕНО');
  const emo = ss.getSheetByName('ЕМО');
  const region = ss.getSheetByName('Отчет  регионы');
  const weekly = ss.getSheetByName('Недельный комментрий') || ss.insertSheet('Недельный комментрий');
  if (!eno || !emo || !region || !weekly) throw new Error('Не найдены листы ЕНО / ЕМО / Отчет регионы / Недельный комментрий');

  // ЕНО: динамический блок комментариев от итоговой строки W/Y и разбивки по типам РК.
  eno.getRange('A955').setFormula('="Итого целевых лидов/конверсий — "&TEXT(W939;"# ##0")&", CR — "&TEXT(X939;"0,00%")&", CPA — "&TEXT(Y939;"# ##0,00 ₽")&". В целевой факт входят Ecommerce: покупка, Jivo-чат, Callibri: Лид_Квал_A и Callibri: Лид_Квал_C."');
  eno.getRange('A958').setFormula('="Поиск — "&TEXT(SUMIF(B920:B938;"Поиск";D920:D938);"# ##0")&" кликов, расход "&TEXT(SUMIF(B920:B938;"Поиск";G920:G938);"# ##0,00 ₽")&", CTR "&TEXT(IFERROR(SUMIF(B920:B938;"Поиск";D920:D938)/SUMIF(B920:B938;"Поиск";C920:C938);0);"0,00%")&", CPC "&TEXT(IFERROR(SUMIF(B920:B938;"Поиск";G920:G938)/SUMIF(B920:B938;"Поиск";D920:D938);0);"# ##0,00 ₽")&", целевые лиды/конверсии "&TEXT(SUMIF(B920:B938;"Поиск";W920:W938);"# ##0")&", CPA "&TEXT(IFERROR(SUMIF(B920:B938;"Поиск";G920:G938)/SUMIF(B920:B938;"Поиск";W920:W938);0);"# ##0,00 ₽")&"."');
  eno.getRange('A959').setFormula('="Сети — "&TEXT(SUMIF(B920:B938;"Сети";D920:D938);"# ##0")&" кликов, расход "&TEXT(SUMIF(B920:B938;"Сети";G920:G938);"# ##0,00 ₽")&", CTR "&TEXT(IFERROR(SUMIF(B920:B938;"Сети";D920:D938)/SUMIF(B920:B938;"Сети";C920:C938);0);"0,00%")&", CPC "&TEXT(IFERROR(SUMIF(B920:B938;"Сети";G920:G938)/SUMIF(B920:B938;"Сети";D920:D938);0);"# ##0,00 ₽")&", целевые лиды/конверсии "&TEXT(SUMIF(B920:B938;"Сети";W920:W938);"# ##0")&", CPA "&TEXT(IFERROR(SUMIF(B920:B938;"Сети";G920:G938)/SUMIF(B920:B938;"Сети";W920:W938);0);"# ##0,00 ₽")&"."');
  eno.getRange('A960').setFormula('="Товарные кампании — "&TEXT(SUMIF(B920:B938;"Товарная";D920:D938);"# ##0")&" кликов, расход "&TEXT(SUMIF(B920:B938;"Товарная";G920:G938);"# ##0,00 ₽")&", CTR "&TEXT(IFERROR(SUMIF(B920:B938;"Товарная";D920:D938)/SUMIF(B920:B938;"Товарная";C920:C938);0);"0,00%")&", CPC "&TEXT(IFERROR(SUMIF(B920:B938;"Товарная";G920:G938)/SUMIF(B920:B938;"Товарная";D920:D938);0);"# ##0,00 ₽")&", целевые лиды/конверсии "&TEXT(SUMIF(B920:B938;"Товарная";W920:W938);"# ##0")&", CPA "&TEXT(IFERROR(SUMIF(B920:B938;"Товарная";G920:G938)/SUMIF(B920:B938;"Товарная";W920:W938);0);"# ##0,00 ₽")&"."');
  eno.getRange('A969').setFormula('="Неделя 29.06–05.07 дала "&TEXT(W939;"# ##0")&" целевых лидов/конверсий по ЕНО при CPA "&TEXT(Y939;"# ##0,00 ₽")&". В целевой факт включены Ecommerce: покупка, Jivo-чат, Callibri: Лид_Квал_A и Callibri: Лид_Квал_C. Спам и нецелевые обращения учитываются отдельно и не передаются в целевую оптимизацию Директа. Основной фокус — удержать рабочие связки, проверить кампании с расходом без лидов и отдельно контролировать качество первичных Callibri/Jivo."');

  // ЕМО: динамические комментарии по месячному/частичному срезу.
  emo.getRange('A264').setFormula('="Итого целевых лидов/конверсий — "&TEXT(W248;"# ##0")&", CR — "&TEXT(X248;"0,00%")&", CPA — "&TEXT(Y248;"# ##0,00 ₽")&". В целевой факт входят Ecommerce: покупка, Jivo-чат, Callibri: Лид_Квал_A и Callibri: Лид_Квал_C."');
  emo.getRange('A274').setFormula('="За 01.07–05.07 получено "&TEXT(W248;"# ##0")&" целевых лидов/конверсий при среднем CPA "&TEXT(Y248;"# ##0,00 ₽")&". В целевой факт включены Ecommerce: покупка, Jivo-чат, Callibri: Лид_Квал_A и Callibri: Лид_Квал_C. Спам и нецелевые обращения учитываются отдельно. Резких отключений по малой статистике не делаем."');
  emo.getRange('A278').setValue('2. Не смешивать Callibri: Спам и Нецелевой_Лид с целевыми лидами для оптимизации Директа.');

  // Региональный отчет: комментарии формулами от региональных строк.
  region.getRange('N6').setFormula('="СПБ: "&TEXT(I6;"# ##0")&" целевых лидов/конверсий, CPA "&TEXT(L6;"# ##0,00 ₽")&". Основной регион, оставить в приоритете и отдельно смотреть качество обращений."');
  region.getRange('N7').setFormula('="РФ: "&TEXT(I7;"# ##0")&" целевых лидов/конверсий, CPA "&TEXT(L7;"# ##0,00 ₽")&". Проверить кампании с расходом без лидов и посадочные страницы."');
  region.getRange('N8').setFormula('="Казахстан: "&TEXT(I8;"# ##0")&" целевых лидов/конверсий, CPA "&TEXT(L8;"# ##0,00 ₽")&". CPA низкий, но качество B2B-обращений проверяем отдельно."');
  region.getRange('N9').setFormula('="СПБ+РФ: "&TEXT(I9;"# ##0")&" целевых лидов/конверсий, CPA "&TEXT(L9;"# ##0,00 ₽")&". Держать в тесте, не масштабировать резко без проверки качества."');
  region.getRange('N21').setFormula('="Итого по региональному срезу 01.07–05.07 с учетом Callibri: Лид_Квал_C: "&TEXT(B21;"# ##0")&" показов, "&TEXT(C21;"# ##0")&" кликов, "&TEXT(E21;"# ##0,00 ₽")&", "&TEXT(I21;"# ##0")&" целевых лидов/конверсий, CPA "&TEXT(L21;"# ##0,00 ₽")&"."');
  region.getRange('A23').setFormula('="Вывод: региональный отчёт пересчитывается динамически с учетом Callibri: Лид_Квал_C. СПБ — "&TEXT(I6;"# ##0")&" целевых лидов/конверсий при CPA "&TEXT(L6;"# ##0,00 ₽")&"; РФ — "&TEXT(I7;"# ##0")&" при CPA "&TEXT(L7;"# ##0,00 ₽")&"; Казахстан — "&TEXT(I8;"# ##0")&" при CPA "&TEXT(L8;"# ##0,00 ₽")&", качество проверяем отдельно; СПБ+РФ — "&TEXT(I9;"# ##0")&" при CPA "&TEXT(L9;"# ##0,00 ₽")&". Резких отключений по малой статистике не делаем."');

  // Недельный комментрий: готовый текст для клиента, полностью динамический от ЕНО и регионов.
  weekly.clear();
  weekly.getRange('A1').setValue('Недельный комментарий для клиента');
  weekly.getRange('A2').setFormula('="Обновлено автоматически: "&TEXT(NOW();"dd.mm.yyyy hh:mm")');
  weekly.getRange('A4').setFormula('="За прошлую неделю реклама отработала заметно лучше: получили "&TEXT(\'ЕНО\'!W939;"# ##0")&" целевых лидов/конверсий при CPA "&TEXT(\'ЕНО\'!Y939;"# ##0,00 ₽")&". В целевой факт входят Ecommerce: покупка, Jivo-чат, Callibri: Лид_Квал_A и Callibri: Лид_Квал_C."&CHAR(10)&CHAR(10)&"Основной результат дали Поиск, товарные кампании и сети: Поиск — "&TEXT(SUMIF(\'ЕНО\'!B920:B938;"Поиск";\'ЕНО\'!W920:W938);"# ##0")&" целевых лидов/конверсий, товарные кампании — "&TEXT(SUMIF(\'ЕНО\'!B920:B938;"Товарная";\'ЕНО\'!W920:W938);"# ##0")&", сети — "&TEXT(SUMIF(\'ЕНО\'!B920:B938;"Сети";\'ЕНО\'!W920:W938);"# ##0")&"."&CHAR(10)&CHAR(10)&"По регионам: СПБ — "&TEXT(\'Отчет  регионы\'!I6;"# ##0")&" лидов при CPA "&TEXT(\'Отчет  регионы\'!L6;"# ##0,00 ₽")&"; РФ — "&TEXT(\'Отчет  регионы\'!I7;"# ##0")&" при CPA "&TEXT(\'Отчет  регионы\'!L7;"# ##0,00 ₽")&"; Казахстан — "&TEXT(\'Отчет  регионы\'!I8;"# ##0")&" при CPA "&TEXT(\'Отчет  регионы\'!L8;"# ##0,00 ₽")&"; СПБ+РФ — "&TEXT(\'Отчет  регионы\'!I9;"# ##0")&" при CPA "&TEXT(\'Отчет  регионы\'!L9;"# ##0,00 ₽")&". По Казахстану отдельно проверяем B2B-качество обращений, чтобы не ориентироваться только на низкую стоимость."&CHAR(10)&CHAR(10)&"Дальше удерживаем рабочие связки, аккуратно усиливаем кампании, которые дают целевые обращения, и отдельно разбираем кампании с расходом без результата: запросы, площадки, объявления и посадочные страницы. Спам и нецелевые обращения не смешиваем с целевыми лидами для оптимизации рекламы."');
  weekly.getRange('A4').setWrap(true);
  weekly.setColumnWidth(1, 900);

  SpreadsheetApp.flush();
  return {
    ok: true,
    weeklySheet: weekly.getName(),
    enoLeads: eno.getRange('W939').getDisplayValue(),
    emoLeads: emo.getRange('W248').getDisplayValue(),
    regionLeads: region.getRange('I21').getDisplayValue(),
    weeklyPreview: weekly.getRange('A4').getDisplayValue().slice(0, 500)
  };
}
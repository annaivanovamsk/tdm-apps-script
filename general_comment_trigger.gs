/**
 * ТДМ — отдельное обновление листа «Общий комментарий».
 * Читает только уже сформированные листы отчётов и не вызывает Директ,
 * Метрику или Callibri API.
 */
const TDM_GENERAL_COMMENT_TRIGGER_20260727 = {
  spreadsheetId: '1zXDDfiRYHJE34iOJg-QfAC0bvoTKMIAuVMfsJS0z5Dg',
  handler: 'tdmRunGeneralCommentMonday20260727',
  timezone: 'Europe/Moscow'
};

/**
 * Понедельничный запуск после основного контура отчётов.
 * Берёт период прошлой полной недели и читает готовый блок ЕНО,
 * предыдущий блок ЕНО, Callibri_Сверка и актуальный «Отчет  регионы».
 */
function tdmRunGeneralCommentMonday20260727() {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);

  try {
    const week = tdmUnifiedPreviousFullWeek_();
    return tdmBuildGeneralCommentForPeriod20260727_(
      week.dateFrom,
      week.dateTo,
      null
    );
  } catch (error) {
    tdmeNotifyError_('ТДМ: общий комментарий не обновился', error);
    throw error;
  } finally {
    lock.releaseLock();
  }
}

/**
 * Оставляет ровно один триггер: понедельник около 10:30 по Москве.
 * Основные отчёты запускаются раньше; комментарий строится после них.
 */
function tdmInstallGeneralCommentMondayTrigger20260727() {
  const config = TDM_GENERAL_COMMENT_TRIGGER_20260727;
  let deleted = 0;

  ScriptApp.getProjectTriggers().forEach(function(trigger) {
    if (trigger.getHandlerFunction() !== config.handler) return;
    ScriptApp.deleteTrigger(trigger);
    deleted++;
  });

  ScriptApp.newTrigger(config.handler)
    .timeBased()
    .onWeekDay(ScriptApp.WeekDay.MONDAY)
    .atHour(10)
    .nearMinute(30)
    .inTimezone(config.timezone)
    .create();

  const count = ScriptApp.getProjectTriggers().filter(function(trigger) {
    return trigger.getHandlerFunction() === config.handler;
  }).length;

  if (count !== 1) {
    throw new Error('Триггер общего комментария установлен неоднозначно: ' + count);
  }

  const result = {
    ok: true,
    deleted: deleted,
    handler: config.handler,
    schedule: 'MONDAY 10:30 Europe/Moscow',
    count: count
  };

  const ss = SpreadsheetApp.openById(config.spreadsheetId);
  const audit = ss.getSheetByName('_GPT Weekly Audit');
  if (audit) {
    audit.appendRow([
      Utilities.formatDate(new Date(), config.timezone, 'dd.MM.yyyy HH:mm:ss'),
      '',
      '',
      'TDM_2026',
      'ТДМ',
      'OK',
      'Apps Script trigger',
      'Общий комментарий: установлен понедельничный триггер 10:30 Europe/Moscow',
      'Источник только сформированные листы ЕНО, Callibri_Сверка и Отчет регионы',
      JSON.stringify(result),
      JSON.stringify(result)
    ]);
  }

  return result;
}

/**
 * telegram_reports.gs
 * Telegram-доставка недельного отчета TDM.
 * Секреты хранить только в Script Properties:
 * - TELEGRAM_BOT_TOKEN
 * - TELEGRAM_CHAT_ID
 */

const TELEGRAM_WEEKLY_REPORT_CONFIG = {
  CLIENT_NAME: 'TDM',
  REPORT_TYPE: 'ЕНО / недельный отчет',
  REPORT_URL: 'https://docs.google.com/spreadsheets/d/1zXDDfiRYHJE34iOJg-QfAC0bvoTKMIAuVMfsJS0z5Dg/edit',
  REPORT_SHEET_NAME: 'ЕНО',
  WEEKLY_REPORT_FUNCTION: 'fillTdmEnoPreviousFullWeek',
  TIMEZONE: 'Europe/Moscow',
  TRIGGER_HOUR: 9
};

function testTelegramSend() {
  telegramSendMessage_(
    'Тест Telegram ✅\n' +
    'Клиент: ' + TELEGRAM_WEEKLY_REPORT_CONFIG.CLIENT_NAME + '\n' +
    'Доставка отчетов подключена.'
  );
}

function sendWeeklyReportToTelegram() {
  const cfg = TELEGRAM_WEEKLY_REPORT_CONFIG;
  const period = telegramGetPreviousFullWeekRange_();
  let reportStatus = 'отчет обновлен';
  let errorText = '';

  try {
    if (cfg.WEEKLY_REPORT_FUNCTION) {
      const fn = globalThis[cfg.WEEKLY_REPORT_FUNCTION];
      if (typeof fn !== 'function') {
        throw new Error('Не найдена функция недельного отчета: ' + cfg.WEEKLY_REPORT_FUNCTION);
      }
      fn();
    } else {
      reportStatus = 'функция недельного отчета не указана';
    }
  } catch (e) {
    reportStatus = 'ошибка при обновлении отчета';
    errorText = String(e && e.stack ? e.stack : e);
    Logger.log(errorText);
  }

  telegramSendMessage_(telegramBuildWeeklyReportMessage_(cfg, period, reportStatus, errorText));
}

function createTelegramWeeklyTriggerMonday() {
  telegramDeleteTriggersFor_('sendWeeklyReportToTelegram');

  ScriptApp.newTrigger('sendWeeklyReportToTelegram')
    .timeBased()
    .onWeekDay(ScriptApp.WeekDay.MONDAY)
    .atHour(TELEGRAM_WEEKLY_REPORT_CONFIG.TRIGGER_HOUR || 10)
    .create();

  Logger.log('Создан Telegram-триггер sendWeeklyReportToTelegram');
}

function deleteTelegramWeeklyTriggers() {
  telegramDeleteTriggersFor_('sendWeeklyReportToTelegram');
}

function telegramDeleteTriggersFor_(handlerName) {
  ScriptApp.getProjectTriggers().forEach(function(trigger) {
    if (trigger.getHandlerFunction() === handlerName) {
      ScriptApp.deleteTrigger(trigger);
    }
  });
}

function telegramBuildWeeklyReportMessage_(cfg, period, reportStatus, errorText) {
  let text =
    'Еженедельный отчет\n' +
    'Клиент: ' + cfg.CLIENT_NAME + '\n' +
    'Период: ' + period.fromRu + ' — ' + period.toRu + '\n' +
    'Тип: ' + cfg.REPORT_TYPE + '\n' +
    'Статус: ' + reportStatus + '\n';

  if (cfg.REPORT_URL) {
    text += '\nСсылка: ' + cfg.REPORT_URL + '\n';
  }

  if (errorText) {
    text += '\nОшибка:\n' + telegramTrim_(errorText, 1200);
  }

  return text;
}

function telegramSendMessage_(text) {
  const props = PropertiesService.getScriptProperties();
  const token = String(props.getProperty('TELEGRAM_BOT_TOKEN') || '').trim();
  const chatId = String(props.getProperty('TELEGRAM_CHAT_ID') || '').trim();

  if (!token) throw new Error('Не найден TELEGRAM_BOT_TOKEN в Script Properties');
  if (!chatId) throw new Error('Не найден TELEGRAM_CHAT_ID в Script Properties');

  const response = UrlFetchApp.fetch('https://api.telegram.org/bot' + token + '/sendMessage', {
    method: 'post',
    contentType: 'application/json',
    payload: JSON.stringify({ chat_id: chatId, text: String(text || ''), disable_web_page_preview: true }),
    muteHttpExceptions: true
  });

  const code = response.getResponseCode();
  const body = response.getContentText();
  Logger.log('Telegram response code: ' + code);
  Logger.log(body);

  if (code < 200 || code >= 300) {
    throw new Error('Telegram API error ' + code + ': ' + body);
  }

  return body;
}

function telegramGetPreviousFullWeekRange_() {
  const tz = TELEGRAM_WEEKLY_REPORT_CONFIG.TIMEZONE || 'Europe/Moscow';
  const now = new Date();
  const today = new Date(Utilities.formatDate(now, tz, 'yyyy/MM/dd 00:00:00'));
  const day = today.getDay();
  const daysSinceMonday = day === 0 ? 6 : day - 1;
  const thisMonday = new Date(today);
  thisMonday.setDate(today.getDate() - daysSinceMonday);
  const prevMonday = new Date(thisMonday);
  prevMonday.setDate(thisMonday.getDate() - 7);
  const prevSunday = new Date(thisMonday);
  prevSunday.setDate(thisMonday.getDate() - 1);

  return {
    from: Utilities.formatDate(prevMonday, tz, 'yyyy-MM-dd'),
    to: Utilities.formatDate(prevSunday, tz, 'yyyy-MM-dd'),
    fromRu: Utilities.formatDate(prevMonday, tz, 'dd.MM.yyyy'),
    toRu: Utilities.formatDate(prevSunday, tz, 'dd.MM.yyyy')
  };
}

function telegramTrim_(text, maxLen) {
  const s = String(text || '');
  return s.length <= maxLen ? s : s.slice(0, maxLen) + '\n...обрезано';
}

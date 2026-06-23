/**
 * GPT Web App Runner
 * Универсальный вход для Custom GPT Action.
 *
 * Важно:
 * - клиент не выбирается автоматически;
 * - запускаются только действия из whitelist GPT_ACTIONS;
 * - проверка статуса читает таблицу без записи;
 * - fillPreviousFullWeek запускает существующую функцию ЕНО: fillTdmEnoPreviousFullWeek().
 */

const GPT_CLIENT_ALIASES = {
  'TDM_2026': 'TDM_2026',
  'ТДМ_2026': 'TDM_2026'
};

const GPT_CLIENTS = {
  TDM_2026: {
    name: 'ТДМ_2026',
    spreadsheetId: '1zXDDfiRYHJE34iOJg-QfAC0bvoTKMIAuVMfsJS0z5Dg'
  }
};

const GPT_ACTIONS = {
  checkReportStatus: function(clientConfig, args) {
    return gptCheckReportStatus_(clientConfig, args);
  },

  buildWeeklyPlan: function() {
    tdmBuildWeeklyPlanReal();
    return {
      message: 'План на неделю сформирован.'
    };
  },

  sendWeeklyPlanEmail: function() {
    tdmSendTasksEmail();
    return {
      message: 'Письмо с задачами отправлено.'
    };
  },

  createDailyAgentTrigger: function() {
    tdmCreateDailyTrigger();
    return {
      message: 'Ежедневный триггер агента создан.'
    };
  },

  fillDbYesterday: function() {
    fillTdmDbYesterday();
    return {
      message: 'ДБ заполнен за вчера.'
    };
  },

  fillDbToYesterday: function() {
    fillTdmDbToYesterday();
    return {
      message: 'ДБ заполнен до вчерашнего дня.'
    };
  },

  createDbTrigger8am: function() {
    createTdmDbTrigger8am();
    return {
      message: 'Триггер ДБ на 08:00 создан.'
    };
  },

  fillPreviousFullWeek: function() {
    fillTdmEnoPreviousFullWeek();
    return {
      message: 'ЕНО за прошлую полную неделю сформирован.'
    };
  },

  createEnoWeeklyTriggerMonday9am: function() {
    createTdmEnoWeeklyTriggerMonday9am();
    return {
      message: 'Еженедельный триггер ЕНО на понедельник 09:00 создан.'
    };
  }
};

function doPost(e) {
  return gptHandleRequest_(e);
}

function doGet(e) {
  return gptHandleRequest_(e);
}

function gptHandleRequest_(e) {
  const startedAt = new Date();

  try {
    const payload = gptReadPayload_(e);

    const rawClient = String(payload.client || '').trim();
    const client = GPT_CLIENT_ALIASES[rawClient] || rawClient;
    const action = String(payload.action || '').trim();
    const args = payload.args || {};

    if (!client) {
      return gptJson_({
        ok: false,
        error: 'Missing client',
        allowedClients: Object.keys(GPT_CLIENTS),
        allowedAliases: Object.keys(GPT_CLIENT_ALIASES)
      });
    }

    if (!GPT_CLIENTS[client]) {
      return gptJson_({
        ok: false,
        error: 'Unknown client',
        client: rawClient,
        normalizedClient: client,
        allowedClients: Object.keys(GPT_CLIENTS),
        allowedAliases: Object.keys(GPT_CLIENT_ALIASES)
      });
    }

    if (!action) {
      return gptJson_({
        ok: false,
        client: client,
        error: 'Missing action',
        allowedActions: Object.keys(GPT_ACTIONS)
      });
    }

    if (!GPT_ACTIONS[action]) {
      return gptJson_({
        ok: false,
        client: client,
        action: action,
        error: 'Unknown action',
        allowedActions: Object.keys(GPT_ACTIONS)
      });
    }

    const result = GPT_ACTIONS[action](GPT_CLIENTS[client], args);

    return gptJson_({
      ok: true,
      client: client,
      clientName: GPT_CLIENTS[client].name,
      action: action,
      startedAt: startedAt.toISOString(),
      finishedAt: new Date().toISOString(),
      result: result || {}
    });

  } catch (err) {
    return gptJson_({
      ok: false,
      error: err && err.message ? err.message : String(err),
      stack: err && err.stack ? err.stack : '',
      startedAt: startedAt.toISOString(),
      finishedAt: new Date().toISOString()
    });
  }
}

function gptReadPayload_(e) {
  if (e && e.postData && e.postData.contents) {
    return JSON.parse(e.postData.contents || '{}');
  }

  if (e && e.parameter) {
    return {
      client: e.parameter.client,
      action: e.parameter.action,
      args: e.parameter.args ? JSON.parse(e.parameter.args) : {}
    };
  }

  return {};
}

function gptCheckReportStatus_(clientConfig, args) {
  const ss = SpreadsheetApp.openById(clientConfig.spreadsheetId);
  const sheets = ss.getSheets();

  return {
    spreadsheetName: ss.getName(),
    spreadsheetUrl: ss.getUrl(),
    sheetsCount: sheets.length,
    sheets: sheets.map(function(sheet) {
      return {
        name: sheet.getName(),
        lastRow: sheet.getLastRow(),
        lastColumn: sheet.getLastColumn()
      };
    }),
    checkedAt: new Date().toISOString()
  };
}

function gptJson_(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

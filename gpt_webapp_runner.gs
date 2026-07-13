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

  checkCallibriAggregate: function(clientConfig, args) {
    return tdmGptCheckCallibriAggregate_(args || {});
  },

  buildWeeklyPlan: function() {
    tdmBuildWeeklyPlanReal();
    return {
      message: 'План на неделю сформирован.'
    };
  },

  sendWeeklyPlanEmail: function() {
    throw new Error('BLOCKED_ARCHIVED_ACTION: отправка weekly plan email отключена из GPT runner. Нужна отдельная проверка актуальной функции отправки.');
  },

  createDailyAgentTrigger: function() {
    throw new Error('BLOCKED_NEEDS_CALLIBRI_MAPPING_AUDIT: ежедневный agent-trigger не создаём, пока отчётный pipeline не прошёл QA.');
  },

  fillDbYesterday: function() {
    throw new Error('BLOCKED_NEEDS_CALLIBRI_MAPPING_AUDIT: ДБ не запускаем через GPT runner, пока Callibri не подключён безопасно.');
  },

  fillDbToYesterday: function() {
    throw new Error('BLOCKED_NEEDS_CALLIBRI_MAPPING_AUDIT: ДБ не запускаем через GPT runner, пока Callibri не подключён безопасно.');
  },

  createDbTrigger8am: function() {
    throw new Error('BLOCKED_NEEDS_CALLIBRI_MAPPING_AUDIT: ежедневный триггер ДБ не создаём, пока Callibri не подключён безопасно.');
  },

  fillPreviousFullWeek: function() {
    throw new Error('BLOCKED_NEEDS_CALLIBRI_MAPPING_AUDIT: ЕНО не запускаем через GPT runner, пока Callibri не размечен по кампаниям.');
  },

  createEnoWeeklyTriggerMonday9am: function() {
    throw new Error('BLOCKED_NEEDS_CALLIBRI_MAPPING_AUDIT: недельный триггер ЕНО не создаём, пока Callibri не размечен по кампаниям.');
  }
};

function tdmGptCheckCallibriAggregate_(args) {
  const dateFrom = String(args.dateFrom || '').trim();
  const dateTo = String(args.dateTo || '').trim();
  if (!/^\d{4}-\d{2}-\d{2}$/.test(dateFrom) || !/^\d{4}-\d{2}-\d{2}$/.test(dateTo)) {
    throw new Error('Нужны args.dateFrom и args.dateTo в формате YYYY-MM-DD');
  }

  const aggregate = args.callibriSnapshot || tdmGptCallibriAggregateYandexCpcByCampaign_(dateFrom, dateTo);
  const rows = Object.keys(aggregate).sort().map(function(key) { return aggregate[key]; });
  const totals = rows.reduce(function(acc, row) {
    acc.totalYandexCpc += row.totalYandexCpc;
    acc.callibriLeadA += row.callibriLeadA;
    acc.callibriLeadB += row.callibriLeadB || 0;
    acc.callibriLeadC += row.callibriLeadC;
    acc.callibriSpam += row.callibriSpam;
    acc.callibriNonTarget += row.callibriNonTarget;
    acc.unknownClass += row.unknownClass;
    return acc;
  }, { totalYandexCpc: 0, callibriLeadA: 0, callibriLeadB: 0, callibriLeadC: 0, callibriSpam: 0, callibriNonTarget: 0, unknownClass: 0 });

  return { dateFrom: dateFrom, dateTo: dateTo, rowsCount: rows.length, totals: totals, rows: rows.slice(0, 200) };
}

function tdmGptCallibriAggregateYandexCpcByCampaign_(dateFrom, dateTo) {
  const days = tdmGptCallibriDateSpanDays_(dateFrom, dateTo);
  if (days < 1 || days > 7) {
    throw new Error('Callibri: один снимок можно получать только за период от 1 до 7 дней. Получено дней: ' + days);
  }

  // Один вызов API на один отчётный период. Снимок дальше только читается:
  // он используется и ДБ, и ЕНО, и контрольной сверкой.
  const json = tdmGptCallibriFetchStatistics_(dateFrom, dateTo);
  const items = tdmGptCallibriExtractSafeItems_(json);
  const aggregate = {};
  const seenInteractions = {};

  items.forEach(function(item) {
    const source = tdmGptNorm_(item.utm_source);
    const medium = tdmGptNorm_(item.utm_medium);
    if ((source !== 'yandex' && source !== 'ya') || medium !== 'cpc') return;

    const apiDate = tdmGptCallibriDateToApi_(item.date);
    if (!apiDate || apiDate < dateFrom || apiDate > dateTo) return;
    if (!tdmGptCallibriIsPrimary_(item)) return;

    // Стабильный ID защищает от повторов между каналами ответа API
    // и от повторного появления того же обращения при повторной загрузке.
    const interactionId = String(item.interactionId || '').trim();
    if (!interactionId) throw new Error('Callibri: обращение без стабильного ID после нормализации.');
    if (seenInteractions[interactionId]) return;
    seenInteractions[interactionId] = true;

    const rawUtmCampaign = String(item.utm_campaign || '').trim();
    const campaignId = String(item.campaign_id || '').trim() || (/^\d+$/.test(rawUtmCampaign) ? rawUtmCampaign : '');
    const campaignKey = tdmGptCallibriCampaignKey_(rawUtmCampaign || campaignId);
    const key = apiDate + '|' + campaignKey;

    if (!aggregate[key]) {
      aggregate[key] = {
        date: apiDate,
        campaignKey: campaignKey,
        campaignId: campaignId,
        utmCampaign: rawUtmCampaign,
        totalYandexCpc: 0,
        callibriLeadA: 0,
        callibriLeadB: 0,
        callibriLeadC: 0,
        callibriSpam: 0,
        callibriNonTarget: 0,
        unknownClass: 0
      };
    }

    aggregate[key].totalYandexCpc += 1;
    const cls = tdmGptCallibriClassKey_(item.status);
    if (cls === 'leadA') aggregate[key].callibriLeadA += 1;
    else if (cls === 'leadB') aggregate[key].callibriLeadB += 1;
    else if (cls === 'leadC') aggregate[key].callibriLeadC += 1;
    else if (cls === 'spam') aggregate[key].callibriSpam += 1;
    else if (cls === 'nonTarget') aggregate[key].callibriNonTarget += 1;
    else aggregate[key].unknownClass += 1;
  });

  return tdmGptCallibriDeepFreeze_(aggregate);
}

function tdmGptCallibriFetchStatistics_(dateFrom, dateTo) {
  const props = PropertiesService.getScriptProperties();
  const baseUrl = String(props.getProperty('CALLIBRI_API_BASE_URL') || 'https://api.callibri.ru').replace(/\/$/, '');
  const siteId = props.getProperty('CALLIBRI_SITE_ID');
  if (!siteId) throw new Error('CALLIBRI_SITE_ID не найден в Script Properties');
  const params = { site_id: siteId, date1: tdmGptCallibriApiToRuDate_(dateFrom), date2: tdmGptCallibriApiToRuDate_(dateTo) };
  const userToken = props.getProperty('CALLIBRI_USER_TOKEN');
  const userEmail = props.getProperty('CALLIBRI_USER_EMAIL');
  const legacyToken = props.getProperty('CALLIBRI_TOKEN');
  const legacyEmail = props.getProperty('CALLIBRI_EMAIL');

  if (userToken && userEmail) {
    params.user_token = userToken;
    params.user_email = userEmail;
  } else {
    if (legacyToken) params.token = legacyToken;
    if (legacyEmail) params.email = legacyEmail;
  }
  const extraRaw = props.getProperty('CALLIBRI_AUTH_QUERY_JSON');
  if (extraRaw) {
    const extra = JSON.parse(extraRaw);
    Object.keys(extra).forEach(function(key) { params[key] = extra[key]; });
  }
  const response = UrlFetchApp.fetch(baseUrl + '/site_get_statistics?' + tdmGptToQueryString_(params), { method: 'get', muteHttpExceptions: true });
  if (response.getResponseCode() < 200 || response.getResponseCode() >= 300) throw new Error('Callibri API error ' + response.getResponseCode() + ': ' + response.getContentText());
  return JSON.parse(response.getContentText());
}

function tdmGptCallibriExtractSafeItems_(json) {
  const result = [];
  const channels = json && json.channels_statistics ? json.channels_statistics : [];
  // REPORT_RULE_2026_07_07_FIXED:
  // Для ДБ/ЕНО/ЕМО считаем только первичные Callibri-обращения по всем статусам: calls + feedbacks + emails.
  // Каждый статус хранится отдельно; в целевые для Директа попадают только первичные Лид_Квал_A/B.
  ['calls', 'feedbacks', 'emails'].forEach(function(arrayName) {
    channels.forEach(function(channel) {
      const rows = channel && channel[arrayName] ? channel[arrayName] : [];
      rows.forEach(function(row) {
        result.push({
          type: arrayName,
          interactionId: tdmGptCallibriStableId_(row, arrayName),
          date: row.date || '',
          status: row.status || '',
          is_lid: row.is_lid || '',
          client_type: row.client_type || row.clientType || row.client_status || row.clientStatus || row.client_kind || row.clientKind || row.customer_type || row.customerType || row.visitor_type || row.visitorType || '',
          lead_number: row.number || row.num || row.n || row.order || row.lead_number || row.leadNumber || '',
          is_first: row.is_first || row.isFirst || row.first || row.is_primary || row.isPrimary || row.primary || '',
          source: row.source || '',
          utm_source: row.utm_source || row.source || '',
          utm_medium: row.utm_medium || '',
          utm_campaign: row.utm_campaign || '',
          campaign_id: row.campaign_id || row.campaignId || row.ya_campaign_id || row.yaCampaignId || ''
        });
      });
    });
  });
  return result;
}

function tdmGptCallibriStableId_(row, arrayName) {
  row = row || {};
  const type = String(arrayName || 'interaction').toLowerCase();
  const idCandidates = [
    row.uuid,
    row.call_id,
    row.callId,
    row.feedback_id,
    row.feedbackId,
    row.email_id,
    row.emailId,
    row.request_id,
    row.requestId,
    row.lead_id,
    row.leadId,
    row.id
  ];

  for (let i = 0; i < idCandidates.length; i++) {
    const value = String(idCandidates[i] == null ? '' : idCandidates[i]).trim();
    if (value) return type + ':id:' + value;
  }

  // У старых ответов Callibri ID может отсутствовать. Тогда строим стабильный
  // SHA-256 отпечаток из полей обращения. Телефон/email наружу не возвращаются.
  const fingerprint = [
    type,
    row.date || '',
    row.status || '',
    row.is_lid || '',
    row.number || row.num || row.n || row.order || row.lead_number || row.leadNumber || '',
    row.utm_source || row.source || '',
    row.utm_medium || '',
    row.utm_campaign || '',
    row.campaign_id || row.campaignId || '',
    row.phone || row.client_phone || row.clientPhone || '',
    row.email || row.client_email || row.clientEmail || '',
    row.message || row.comment || row.text || ''
  ].map(function(value) { return String(value == null ? '' : value).trim(); }).join('|');

  const digest = Utilities.computeDigest(
    Utilities.DigestAlgorithm.SHA_256,
    fingerprint,
    Utilities.Charset.UTF_8
  );
  const hex = digest.map(function(byte) {
    const normalized = byte < 0 ? byte + 256 : byte;
    return ('0' + normalized.toString(16)).slice(-2);
  }).join('');

  return type + ':sha256:' + hex;
}

function tdmGptCallibriDateSpanDays_(dateFrom, dateTo) {
  const start = tdmGptParseApiDate_(dateFrom);
  const end = tdmGptParseApiDate_(dateTo);
  if (isNaN(start.getTime()) || isNaN(end.getTime()) || end < start) return 0;
  return Math.floor((end.getTime() - start.getTime()) / 86400000) + 1;
}

function tdmGptCallibriDeepFreeze_(value) {
  if (!value || typeof value !== 'object' || Object.isFrozen(value)) return value;
  Object.keys(value).forEach(function(key) {
    tdmGptCallibriDeepFreeze_(value[key]);
  });
  return Object.freeze(value);
}

function tdmGptCallibriClassKey_(status) {
  const s = tdmGptNorm_(status).replace(/[\s\-]+/g, '_');
  if (!s) return 'unknown';
  if (s.indexOf('спам') !== -1 || s.indexOf('spam') !== -1) return 'spam';
  if (s.indexOf('нецелев') !== -1 || s.indexOf('non_target') !== -1 || s.indexOf('nontarget') !== -1) return 'nonTarget';
  if (s.indexOf('лид_квал_a') !== -1 || s === 'a' || s.indexOf('qual_a') !== -1) return 'leadA';
  if (s.indexOf('лид_квал_b') !== -1 || s === 'b' || s.indexOf('qual_b') !== -1) return 'leadB';
  if (s.indexOf('лид_квал_c') !== -1 || s === 'c' || s.indexOf('qual_c') !== -1) return 'leadC';
  return 'unknown';
}

function tdmGptCallibriIsPrimary_(item) {
  item = item || {};
  const text = tdmGptNorm_(item.client_type || item.is_first || '');
  const leadNumber = String(item.lead_number || '').trim();

  if (text.indexOf('повтор') !== -1 || text.indexOf('repeat') !== -1 || text.indexOf('secondary') !== -1) return false;
  if (text.indexOf('первич') !== -1 || text.indexOf('primary') !== -1 || text === 'true' || text === '1' || text === 'yes') return true;
  if (leadNumber) return leadNumber === '1';

  // Если Callibri API не отдал признак первичности, не теряем обращение, но статус остаётся раздельным.
  return true;
}

function tdmGptCallibriCampaignKey_(value) { return tdmGptNorm_(value).replace(/[{}]/g, '').replace(/\s+/g, '_') || '__no_campaign__'; }
function tdmGptCallibriDateToApi_(value) { const text = String(value || '').trim(); let m = text.match(/^(\d{1,2})\.(\d{1,2})\.(\d{4})/); if (m) return [m[3], String(m[2]).padStart(2, '0'), String(m[1]).padStart(2, '0')].join('-'); m = text.match(/^(\d{4})-(\d{1,2})-(\d{1,2})/); if (m) return [m[1], String(m[2]).padStart(2, '0'), String(m[3]).padStart(2, '0')].join('-'); return ''; }
function tdmGptCallibriApiToRuDate_(apiDate) { const p = String(apiDate).split('-'); return [p[2], p[1], p[0]].join('.'); }
function tdmGptCallibriSplitByMaxDays_(dateFrom, dateTo, maxDays) { const result = []; let start = tdmGptParseApiDate_(dateFrom); const end = tdmGptParseApiDate_(dateTo); while (start <= end) { const chunkEnd = new Date(start); chunkEnd.setDate(chunkEnd.getDate() + maxDays - 1); if (chunkEnd > end) chunkEnd.setTime(end.getTime()); result.push({ dateFrom: tdmGptFormatApiDate_(start), dateTo: tdmGptFormatApiDate_(chunkEnd) }); start = new Date(chunkEnd); start.setDate(start.getDate() + 1); } return result; }
function tdmGptParseApiDate_(apiDate) { const p = String(apiDate).split('-').map(Number); return new Date(p[0], p[1] - 1, p[2]); }
function tdmGptFormatApiDate_(date) { return Utilities.formatDate(date, Session.getScriptTimeZone(), 'yyyy-MM-dd'); }
function tdmGptToQueryString_(params) { return Object.keys(params).filter(function(key) { return params[key] !== '' && params[key] !== null && params[key] !== undefined; }).map(function(key) { return encodeURIComponent(key) + '=' + encodeURIComponent(params[key]); }).join('&'); }
function tdmGptNorm_(value) { return String(value || '').toLowerCase().replace(/ё/g, 'е').replace(/[«»"']/g, '').trim(); }

function archived_tdmGptCheckCallibriAggregate20260629_20260705() {
  return tdmGptCheckCallibriAggregate_({
    dateFrom: '2026-06-29',
    dateTo: '2026-07-05'
  });
}

function archived_tdmGptInspectCallibriSafe20260629_20260705() {
  const json = tdmGptCallibriFetchStatistics_('2026-06-29', '2026-07-05');
  const items = tdmGptCallibriExtractSafeItems_(json).filter(function(item) {
    return tdmGptNorm_(item.utm_source) === 'yandex' && tdmGptNorm_(item.utm_medium) === 'cpc';
  });
  const byDate = {};
  const byType = {};
  const byIsLid = {};
  const byStatus = {};
  items.forEach(function(item) {
    const date = tdmGptCallibriDateToApi_(item.date) || '__no_date__';
    const type = item.type || '__no_type__';
    const isLid = String(item.is_lid || '__empty__');
    const status = String(item.status || '__empty__');
    byDate[date] = (byDate[date] || 0) + 1;
    byType[type] = (byType[type] || 0) + 1;
    byIsLid[isLid] = (byIsLid[isLid] || 0) + 1;
    byStatus[status] = (byStatus[status] || 0) + 1;
  });
  return { total: items.length, byDate: byDate, byType: byType, byIsLid: byIsLid, byStatus: byStatus };
}

function archived_tdmGptCheckCallibriPropsSafe() {
  const props = PropertiesService.getScriptProperties();
  const all = props.getProperties();
  const keys = [
    'CALLIBRI_API_BASE_URL',
    'CALLIBRI_SITE_ID',
    'CALLIBRI_TOKEN',
    'CALLIBRI_EMAIL',
    'CALLIBRI_AUTH_QUERY_JSON',
    'CALLIBRI_API_TOKEN',
    'CALLIBRI_API_KEY',
    'CALLIBRI_KEY',
    'CALLIBRI_LOGIN',
    'CALLIBRI_USER_EMAIL',
    'CALLIBRI_USER_TOKEN'
  ];
  const result = {};
  keys.forEach(function(key) {
    const value = props.getProperty(key);
    result[key] = {
      exists: !!value,
      length: value ? String(value).length : 0
    };
  });
  result.detectedCallibriLikeKeys = Object.keys(all).filter(function(key) {
    const k = String(key || '').toLowerCase();
    return k.indexOf('callibri') !== -1;
  }).sort();
  result.detectedAuthLikeKeys = Object.keys(all).filter(function(key) {
    const k = String(key || '').toLowerCase();
    return k.indexOf('token') !== -1 || k.indexOf('api') !== -1 || k.indexOf('key') !== -1 || k.indexOf('email') !== -1 || k.indexOf('login') !== -1;
  }).sort();
  return result;
}

function archived_tdmGptProbeCallibriAuthSafe() {
  const props = PropertiesService.getScriptProperties();
  const siteId = props.getProperty('CALLIBRI_SITE_ID');
  const email = props.getProperty('CALLIBRI_EMAIL') || props.getProperty('CALLIBRI_USER_EMAIL');
  const token = props.getProperty('CALLIBRI_TOKEN') || props.getProperty('CALLIBRI_USER_TOKEN');
  const bases = ['https://api.callibri.ru', 'https://callibri.ru', 'https://callibri.ru/api'];
  const authModes = [
    { name: 'user_email_user_token', params: { user_email: email, user_token: token } },
    { name: 'email_token', params: { email: email, token: token } },
    { name: 'user_email_token', params: { user_email: email, token: token } },
    { name: 'email_user_token', params: { email: email, user_token: token } }
  ];
  const result = [];
  bases.forEach(function(base) {
    authModes.forEach(function(mode) {
      const params = Object.assign({ site_id: siteId, date1: '29.06.2026', date2: '29.06.2026' }, mode.params);
      const url = base + '/site_get_statistics?' + tdmGptToQueryString_(params);
      try {
        const response = UrlFetchApp.fetch(url, { method: 'get', muteHttpExceptions: true });
        result.push({ base: base, authMode: mode.name, code: response.getResponseCode(), bodyStart: String(response.getContentText() || '').slice(0, 160) });
      } catch (e) {
        result.push({ base: base, authMode: mode.name, error: String(e && e.message ? e.message : e) });
      }
      Utilities.sleep(1100);
    });
  });
  return result;
}

function archived_doPost(e) {
  return gptHandleRequest_(e);
}

function archived_doGet(e) {
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

// =====================
// TDM Callibri report helpers
// SAFETY_2026_07_07:
// Общие helper-функции для ЕНО/ЕМО.
// Не запускают отчёты, не пишут в таблицу, не меняют Директ/Метрику.
// Целевые Callibri-лиды для Директа = только первичные Лид_Квал_A + Лид_Квал_B.
// Лид_Квал_C, Спам, Нецелевой_Лид, повторные и unknownClass не входят в целевые лиды.
// =====================

function tdmCallibriRollupByCampaign_(callibriByDateCampaign) {
  const result = {};

  Object.keys(callibriByDateCampaign || {}).forEach(function(key) {
    const row = callibriByDateCampaign[key] || {};
    const campaignKey = row.campaignKey || '__no_campaign__';

    if (!result[campaignKey]) {
      result[campaignKey] = {
        campaignKey: campaignKey,
        campaignId: String(row.campaignId || ''),
        utmCampaign: String(row.utmCampaign || ''),
        totalYandexCpc: 0,
        callibriLeadA: 0,
        callibriLeadB: 0,
        callibriLeadC: 0,
        callibriSpam: 0,
        callibriNonTarget: 0,
        unknownClass: 0
      };
    }

    result[campaignKey].totalYandexCpc += Number(row.totalYandexCpc || 0);
    result[campaignKey].callibriLeadA += Number(row.callibriLeadA || 0);
    result[campaignKey].callibriLeadB += Number(row.callibriLeadB || 0);
    result[campaignKey].callibriLeadC += Number(row.callibriLeadC || 0);
    result[campaignKey].callibriSpam += Number(row.callibriSpam || 0);
    result[campaignKey].callibriNonTarget += Number(row.callibriNonTarget || 0);
    result[campaignKey].unknownClass += Number(row.unknownClass || 0);

    if (!result[campaignKey].campaignId && row.campaignId) {
      result[campaignKey].campaignId = String(row.campaignId);
    }
    if (!result[campaignKey].utmCampaign && row.utmCampaign) {
      result[campaignKey].utmCampaign = String(row.utmCampaign);
    }
  });

  return result;
}

function tdmCallibriEmptyReportRow_() {
  return {
    totalYandexCpc: 0,
    callibriLeadA: 0,
    callibriLeadB: 0,
    callibriLeadC: 0,
    callibriSpam: 0,
    callibriNonTarget: 0,
    unknownClass: 0
  };
}

function tdmCallibriTargetLeads_(row) {
  row = row || {};
  return Number(row.callibriLeadA || 0) + Number(row.callibriLeadB || 0);
}

function tdmCallibriGetByCampaign_(callibriByCampaign, campaignName) {
  const campaignKey = tdmGptCallibriCampaignKey_(campaignName);
  return (callibriByCampaign || {})[campaignKey] || tdmCallibriEmptyReportRow_();
}

/**
 * ONE_TIME_FIX_2026_07_07
 * Исправляет сверку ДБ ↔ ЕНО для переходной недели 29.06.2026–05.07.2026.
 * Не трогает расход, клики, показы, CTR/CPC/CPA и строку ИТОГО напрямую.
 * Правит только входные ячейки целевых действий в строках дат, чтобы формулы T/W/ИТОГО пересчитались сами.
 */
function tdmFixDbEnoBridgeWeek20260707() {
  throw new Error('ARCHIVED_ONE_TIME_FIX_DISABLED: tdmFixDbEnoBridgeWeek20260707 отключена после рефакторинга. Не запускать из runtime.');
  const ss = SpreadsheetApp.openById('1zXDDfiRYHJE34iOJg-QfAC0bvoTKMIAuVMfsJS0z5Dg');
  const db = ss.getSheetByName('ДБ');
  if (!db) throw new Error('Лист ДБ не найден');

  const rows = {
    '2026-06-29': 264,
    '2026-06-30': 265,
    '2026-07-01': 290,
    '2026-07-02': 291,
    '2026-07-03': 292,
    '2026-07-04': 293,
    '2026-07-05': 294
  };

  // Июньский legacy-блок: Q=чат/TalkMe-compatible, R=leadA-compatible, S=leadC-compatible.
  // Добавляем недостающие Jivo/LeadA и убираем legacy-переизбыток LeadC.
  db.getRange('Q264').setValue(4);
  db.getRange('R264').setValue(2);
  db.getRange('S264').setValue(0);

  db.getRange('Q265').setValue(4);
  db.getRange('R265').setValue(1);
  db.getRange('S265').setValue(0);

  // Июльский Callibri-блок: S292 был лишним C-лидом для сверки с ЕНО.
  db.getRange('S292').setValue(0);

  SpreadsheetApp.flush();

  const control = tdmFixDbEnoBridgeWeekControl20260707_(db, rows);
  if (control.fact !== 24 || control.jivo !== 17 || control.leadA !== 4 || control.leadC !== 2) {
    throw new Error('После фикса сверка всё ещё не сошлась: ' + JSON.stringify(control));
  }

  return {
    ok: true,
    period: '2026-06-29—2026-07-05',
    changedCells: ['Q264', 'R264', 'S264', 'Q265', 'R265', 'S265', 'S292'],
    control: control
  };
}

function tdmFixDbEnoBridgeWeekControl20260707_(db, rows) {
  const n = function(a1) {
    const v = db.getRange(a1).getDisplayValue();
    const x = String(v || '').replace(/[^0-9,.-]/g, '').replace(',', '.');
    return x === '' || x === '-' ? 0 : Number(x) || 0;
  };

  const result = { jivo: 0, leadA: 0, leadC: 0, fact: 0 };

  // 29–30 июня: legacy-колонки приводим к новой логике сверки.
  [264, 265].forEach(function(r) {
    result.jivo += n('Q' + r);
    result.leadA += n('R' + r);
    result.leadC += n('S' + r);
    result.fact += n('T' + r);
  });

  // 01–05 июля: текущие Callibri/Jivo-колонки.
  [290, 291, 292, 293, 294].forEach(function(r) {
    result.jivo += n('Q' + r);
    result.leadA += n('R' + r);
    result.leadC += n('S' + r);
    result.fact += n('T' + r);
  });

  return result;
}

/**
 * ONE_TIME_FIX_2026_07_07_EMO
 * Приводит ЕМО 01.07–05.07 к исправленному ДБ после Callibri/Jivo-маппинга.
 * Меняет только входной C-лид в строке кампании и видимые summary/comment values этого блока.
 */
function tdmFixEmoJuly0105AfterDbBridge20260707() {
  throw new Error('ARCHIVED_ONE_TIME_FIX_DISABLED: tdmFixEmoJuly0105AfterDbBridge20260707 отключена после рефакторинга. Не запускать из runtime.');
  const ss = SpreadsheetApp.openById('1zXDDfiRYHJE34iOJg-QfAC0bvoTKMIAuVMfsJS0z5Dg');
  const emo = ss.getSheetByName('ЕМО');
  if (!emo) throw new Error('Лист ЕМО не найден');

  // Блок ЕМО: 01.07–05.07.2026, header row 487, data rows 488:505, total row 506.
  // Убираем лишний Callibri: Лид_Квал_C из первой строки блока: W488 1 -> 0.
  emo.getRange('W488').setValue(0);

  SpreadsheetApp.flush();

  const clicks = Number(emo.getRange('D506').getValue() || 0);
  const cost = Number(emo.getRange('G506').getValue() || 0);
  const leadC = Number(emo.getRange('W506').getValue() || 0);
  const fact = Number(emo.getRange('X506').getValue() || 0);
  const cr = clicks ? fact / clicks : 0;
  const cpa = fact ? cost / fact : 0;

  if (leadC !== 2 || fact !== 12) {
    throw new Error('ЕМО после фикса не сошёлся: ' + JSON.stringify({ leadC: leadC, fact: fact, clicks: clicks, cost: cost }));
  }

  // Верхняя карточка блока.
  emo.getRange('Q485').setValue(12);
  emo.getRange('S485').setValue(cr);
  emo.getRange('U485').setValue(cpa);

  // Текстовые комментарии блока.
  emo.getRange('A529').setValue('Callibri: Лид_Квал_C — 2');
  emo.getRange('A530').setValue('Итого получено: 12 лидов, CR — 0,70 %, CPA — 3 403,35 руб.');
  emo.getRange('A536').setValue('Поиск — 376 кликов, расход 29 511,43 руб., средний CPC — 78,49 руб., получено — 7 лидов, CR — 1,86 %, CPA — 4 215,92 руб.');
  emo.getRange('A545').setValue('За июль 01.07–05.07 получено 12 лидов при среднем CPA 3 403,35 руб. Основной вклад по лидам даёт Поиск: 7 лидов. Сети и товарные кампании тоже дают лиды, поэтому их не отключаем резко — проверяем качество обращений, запросы, площадки и посадочные.');

  SpreadsheetApp.flush();

  return {
    ok: true,
    period: '2026-07-01—2026-07-05',
    changedCells: ['W488', 'Q485', 'S485', 'U485', 'A529', 'A530', 'A536', 'A545'],
    control: { leadC: leadC, fact: fact, cr: cr, cpa: cpa }
  };
}

/**
 * STRICT_VALIDATOR_2026_07_07
 * ДБ ↔ ЕНО ↔ ЕМО: явная сверка всех текущих целей Callibri/Jivo/Ecommerce.
 * Ничего не пишет в таблицу. Только читает итоговые строки и возвращает diff по каждой цели.
 */
function tdmStrictValidateDbEnoEmoGoals20260707() {
  const ss = SpreadsheetApp.openById('1zXDDfiRYHJE34iOJg-QfAC0bvoTKMIAuVMfsJS0z5Dg');
  const db = ss.getSheetByName('ДБ');
  const eno = ss.getSheetByName('ЕНО');
  const emo = ss.getSheetByName('ЕМО');
  if (!db || !eno || !emo) throw new Error('Не найдены листы ДБ/ЕНО/ЕМО');

  const enoDb = tdmStrictDbTotals20260707_(db, 264, 265, 290, 294);
  const emoDb = tdmStrictDbTotals20260707_(db, 290, 294);
  const enoReport = tdmStrictReportTotals20260707_(eno, 'Отчет недельный по РК  29.06.2026 - 05.07.2026', 3, 4, 7, 16, 17, 18, 19, 20, 21, 22, 23);
  const emoReport = tdmStrictReportTotals20260707_(emo, 'Отчет месячный по РК 1.07. - 5.07.2026', 3, 4, 7, 16, 17, 19, 20, 21, 22, 23, 24);

  const enoCompare = tdmStrictCompareTotals20260707_(enoDb, enoReport);
  const emoCompare = tdmStrictCompareTotals20260707_(emoDb, emoReport);

  return {
    ok: enoCompare.ok && emoCompare.ok,
    checkedGoals: ['impressions', 'clicks', 'cost', 'spam', 'nonTarget', 'addToCart', 'purchase', 'jivo', 'leadA', 'leadC', 'fact'],
    eno: { sheet: 'ЕНО', dateFrom: '2026-06-29', dateTo: '2026-07-05', ok: enoCompare.ok, db: enoDb, report: enoReport, diff: enoCompare.diff },
    emo: { sheet: 'ЕМО', dateFrom: '2026-07-01', dateTo: '2026-07-05', ok: emoCompare.ok, db: emoDb, report: emoReport, diff: emoCompare.diff }
  };
}

function tdmStrictDbTotals20260707_(db, startRow1, endRow1, startRow2, endRow2) {
  const result = tdmStrictEmptyTotals20260707_();
  [[startRow1, endRow1], [startRow2, endRow2]].forEach(function(pair) {
    if (!pair[0] || !pair[1]) return;
    const rows = db.getRange(pair[0], 1, pair[1] - pair[0] + 1, 23).getDisplayValues();
    rows.forEach(function(row) {
      result.impressions += tdmStrictNum20260707_(row[2]);
      result.clicks += tdmStrictNum20260707_(row[3]);
      result.cost += tdmStrictNum20260707_(row[5]);
      result.spam += tdmStrictNum20260707_(row[12]);
      result.nonTarget += tdmStrictNum20260707_(row[13]);
      result.addToCart += tdmStrictNum20260707_(row[14]);
      result.purchase += tdmStrictNum20260707_(row[15]);
      result.jivo += tdmStrictNum20260707_(row[16]);
      result.leadA += tdmStrictNum20260707_(row[17]);
      result.leadC += tdmStrictNum20260707_(row[18]);
      result.fact += tdmStrictNum20260707_(row[19]);
    });
  });
  result.cost = Math.round(result.cost * 100) / 100;
  return result;
}

function tdmStrictReportTotals20260707_(sheet, title, impressionsCol, clicksCol, costCol, spamCol, nonTargetCol, addToCartCol, purchaseCol, jivoCol, leadACol, leadCCol, factCol) {
  const values = sheet.getDataRange().getDisplayValues();
  let titleRow = -1;
  for (let i = 0; i < values.length; i++) {
    if (String(values[i][0] || '').trim() === title) titleRow = i + 1;
  }
  if (titleRow < 0) throw new Error('Не найден блок: ' + title);

  let totalRow = -1;
  for (let r = titleRow + 1; r <= values.length; r++) {
    if (String(values[r - 1][0] || '').trim().toUpperCase() === 'ИТОГО') {
      totalRow = r;
      break;
    }
  }
  if (totalRow < 0) throw new Error('Не найдена строка ИТОГО для блока: ' + title);

  const row = values[totalRow - 1];
  return {
    impressions: tdmStrictNum20260707_(row[impressionsCol - 1]),
    clicks: tdmStrictNum20260707_(row[clicksCol - 1]),
    cost: Math.round(tdmStrictNum20260707_(row[costCol - 1]) * 100) / 100,
    spam: tdmStrictNum20260707_(row[spamCol - 1]),
    nonTarget: tdmStrictNum20260707_(row[nonTargetCol - 1]),
    addToCart: tdmStrictNum20260707_(row[addToCartCol - 1]),
    purchase: tdmStrictNum20260707_(row[purchaseCol - 1]),
    jivo: tdmStrictNum20260707_(row[jivoCol - 1]),
    leadA: tdmStrictNum20260707_(row[leadACol - 1]),
    leadC: tdmStrictNum20260707_(row[leadCCol - 1]),
    fact: tdmStrictNum20260707_(row[factCol - 1])
  };
}

function tdmStrictCompareTotals20260707_(db, report) {
  const diff = {};
  let ok = true;
  Object.keys(tdmStrictEmptyTotals20260707_()).forEach(function(key) {
    const tolerance = key === 'cost' ? 1 : 0;
    const delta = Math.round((Number(report[key] || 0) - Number(db[key] || 0)) * 100) / 100;
    diff[key] = { db: db[key], report: report[key], delta: delta, ok: Math.abs(delta) <= tolerance };
    if (!diff[key].ok) ok = false;
  });
  return { ok: ok, diff: diff };
}

function tdmStrictEmptyTotals20260707_() {
  return { impressions: 0, clicks: 0, cost: 0, spam: 0, nonTarget: 0, addToCart: 0, purchase: 0, jivo: 0, leadA: 0, leadC: 0, fact: 0 };
}

function tdmStrictNum20260707_(value) {
  const s = String(value || '').replace(/\s/g, '').replace(/[^0-9,.-]/g, '').replace(',', '.');
  if (!s || s === '-') return 0;
  return Number(s) || 0;
}

function tdmFixEnoEmoComments20260707() {
  throw new Error('ARCHIVED_ONE_TIME_FIX_DISABLED: tdmFixEnoEmoComments20260707 отключена после рефакторинга. Не запускать из runtime.');
  const ss = SpreadsheetApp.openById('1zXDDfiRYHJE34iOJg-QfAC0bvoTKMIAuVMfsJS0z5Dg');
  const eno = ss.getSheetByName('ЕНО');
  const emo = ss.getSheetByName('ЕМО');
  if (!eno || !emo) throw new Error('Не найдены листы ЕНО/ЕМО');

  eno.getRange('A948').setValue('Callibri: Спам — 2');
  eno.getRange('A949').setValue('Callibri: Нецелевой_Лид — 4');

  const emoComments = [
    ['Комментарии к отчёту за июль с 1.07. по 5.07.2026'],
    [''],
    ['Бюджет:'],
    ['Всего задействовано за период 40 840,22 руб.'],
    [''],
    ['Трафик:'],
    ['Показы — 67 846'],
    ['Клики — 1 720'],
    ['Средний CPC — 23,74 руб.'],
    [''],
    ['Достигнутые цели и события:'],
    ['Просмотр 3х страниц — 129'],
    ['Callibri: Спам — 0'],
    ['Callibri: Нецелевой_Лид — 4'],
    ['Ecommerce: добавление в корзину — 6'],
    ['Ecommerce: покупка — 0'],
    ['Jivo-чат — 9'],
    ['Callibri: Лид_Квал_A — 1'],
    ['Callibri: Лид_Квал_C — 2'],
    ['Итого целевых лидов/конверсий — 12, CR — 0,70 %, CPA — 3 403,35 руб.'],
    [''],
    ['По типам РК:'],
    ['Поиск — 376 кликов, расход 29 511,43 руб., получено 7 лидов, CR — 1,86 %, CPA — 4 215,92 руб.'],
    ['Сети — 218 кликов, расход 3 978,29 руб., получено 2 лида, CR — 0,92 %, CPA — 1 989,15 руб.'],
    ['Поиск/Сети — 1 126 кликов, расход 7 350,50 руб., получено 3 лида, CR — 0,27 %, CPA — 2 450,17 руб.'],
    [''],
    ['Казахстан:'],
    ['Казахстанские кампании держим отдельным контролем: проверяем B2B-смысл обращений и не делаем резких выводов по малой статистике.'],
    [''],
    ['Вывод:'],
    ['За 01.07–05.07 получено 12 целевых лидов/конверсий при среднем CPA 3 403,35 руб. Основной вклад даёт Поиск; сетевые и товарные кампании не отключаем резко — проверяем качество обращений, запросы, площадки и посадочные.'],
    [''],
    ['Фокус:'],
    ['1. Проверить кампании с расходом без целевых лидов: запросы, площадки, посадочные.'],
    ['2. Не смешивать Callibri: Спам и Нецелевой_Лид с целевыми лидами для оптимизации.'],
    ['3. По Казахстану отдельно проверить B2B-смысл и качество обращений.']
  ];

  emo.getRange('A250:A290').clearContent();
  emo.getRange(250, 1, emoComments.length, 1).setValues(emoComments);
  SpreadsheetApp.flush();

  const enoCheck = eno.getRange('A948:A949').getDisplayValues().map(function(r) { return r[0]; });
  const emoCheck = emo.getRange('A250:A285').getDisplayValues().map(function(r) { return r[0]; });
  if (enoCheck[0] !== 'Callibri: Спам — 2' || enoCheck[1] !== 'Callibri: Нецелевой_Лид — 4') {
    throw new Error('ЕНО комментарии не обновились: ' + JSON.stringify(enoCheck));
  }
  if (emoCheck.indexOf('Итого целевых лидов/конверсий — 12, CR — 0,70 %, CPA — 3 403,35 руб.') === -1) {
    throw new Error('ЕМО комментарии не обновились');
  }

  return {
    ok: true,
    changedRanges: ['ЕНО!A948:A949', 'ЕМО!A250:A290'],
    eno: enoCheck,
    emoRows: emoComments.length
  };
}

function tdmFixEmoRow227ThreePages20260707() {
  throw new Error('ARCHIVED_ONE_TIME_FIX_DISABLED: tdmFixEmoRow227ThreePages20260707 отключена после рефакторинга. Не запускать из runtime.');
  const ss = SpreadsheetApp.openById('1zXDDfiRYHJE34iOJg-QfAC0bvoTKMIAuVMfsJS0z5Dg');
  const emo = ss.getSheetByName('ЕМО');
  if (!emo) throw new Error('Лист ЕМО не найден');
  emo.getRange('M233').setValue(18);
  SpreadsheetApp.flush();
  const total = String(emo.getRange('M248').getDisplayValue() || '').replace(/\s/g, '');
  if (total !== '129') throw new Error('M248 не стал 129: ' + total);
  return { ok: true, changedCells: ['ЕМО!M233'], control: { threePages: 129 } };
}

/**
 * ONE_TIME_FIX_2026_07_07_EMO_ROW_227_GOALS
 * Исправляет дубль ЕМО с 227 строки: текущие цели Callibri/Jivo/Ecommerce.
 */
function tdmFixEmoRow227Goals20260707() {
  throw new Error('ARCHIVED_ONE_TIME_FIX_DISABLED: tdmFixEmoRow227Goals20260707 отключена после рефакторинга. Не запускать из runtime.');
  const ss = SpreadsheetApp.openById('1zXDDfiRYHJE34iOJg-QfAC0bvoTKMIAuVMfsJS0z5Dg');
  const emo = ss.getSheetByName('ЕМО');
  if (!emo) throw new Error('Лист ЕМО не найден');

  const src = emo.getRange('P487:X504').getDisplayValues();
  const mapped = src.map(function(r) {
    return [r[0], r[1], r[3], r[4], r[5], r[6], r[7], r[8]];
  });

  emo.getRange('P230:W247').setValues(mapped);
  emo.getRange('Q227').setValue(12);
  emo.getRange('S227').setValue(0.006976744186);
  emo.getRange('U227').setValue(3403.351667);

  emo.getRange('A262:A271').setValues([
    ['Достигнуты цели и события:'],
    ['Просмотр 3х страниц — 129'],
    ['Callibri: Спам — 0'],
    ['Callibri: Нецелевой_Лид — 4'],
    ['Ecommerce: добавление в корзину — 6'],
    ['Ecommerce: покупка — 0'],
    ['Jivo-чат — 9'],
    ['Callibri: Лид_Квал_A — 1'],
    ['Callibri: Лид_Квал_C — 2'],
    ['Итого получено: 12 лидов, CR — 0,70 %, CPA — 3 403,35 руб.']
  ]);
  emo.getRange('A276').setValue('Поиск — 376 кликов, расход 29 511,43 руб., средний CPC — 78,49 руб., получено — 7 лидов, CR — 1,86 %, CPA — 4 215,92 руб.');
  emo.getRange('A277').setValue('Сети — 218 кликов, расход 3 978,29 руб., средний CPC — 18,25 руб., получено — 2 лида, CR — 0,92 %, CPA — 1 989,15 руб.');
  emo.getRange('A278').setValue('Поиск/Сети — 1 126 кликов, расход 7 350,50 руб., средний CPC — 6,53 руб., получено — 3 лида, CR — 0,27 %, CPA — 2 450,17 руб.');

  SpreadsheetApp.flush();

  const total = emo.getRange('P248:W248').getDisplayValues()[0];
  const control = {
    spam: tdmStrictNum20260707_(total[0]),
    nonTarget: tdmStrictNum20260707_(total[1]),
    addToCart: tdmStrictNum20260707_(total[2]),
    purchase: tdmStrictNum20260707_(total[3]),
    jivo: tdmStrictNum20260707_(total[4]),
    leadA: tdmStrictNum20260707_(total[5]),
    leadC: tdmStrictNum20260707_(total[6]),
    fact: tdmStrictNum20260707_(total[7])
  };
  const expected = JSON.stringify({spam:0, nonTarget:4, addToCart:6, purchase:0, jivo:9, leadA:1, leadC:2, fact:12});
  if (JSON.stringify(control) !== expected) {
    throw new Error('ЕМО row227 goals не сошлись: ' + JSON.stringify(control));
  }

  return {
    ok: true,
    changedRanges: ['ЕМО!P230:W247', 'ЕМО!Q227', 'ЕМО!S227', 'ЕМО!U227', 'ЕМО!A262:A271', 'ЕМО!A276:A278'],
    control: control
  };
}

/**
 * ONE_TIME_FIX_2026_07_07_SPAM_NONTARGET
 * Доводит ДБ ↔ ЕНО ↔ ЕМО по Callibri: Спам и Callibri: Нецелевой_Лид.
 * Не трогает расход, показы, клики, fact, CR/CPA и строки ИТОГО напрямую.
 */
function tdmFixSpamNonTargetTotals20260707() {
  throw new Error('ARCHIVED_ONE_TIME_FIX_DISABLED: tdmFixSpamNonTargetTotals20260707 отключена после рефакторинга. Не запускать из runtime.');
  const ss = SpreadsheetApp.openById('1zXDDfiRYHJE34iOJg-QfAC0bvoTKMIAuVMfsJS0z5Dg');
  const eno = ss.getSheetByName('ЕНО');
  const emo = ss.getSheetByName('ЕМО');
  if (!eno || !emo) throw new Error('Не найдены листы ЕНО/ЕМО');

  // ЕНО 29.06–05.07: ИТОГО P/Q должно стать 2/4 через входные ячейки.
  eno.getRange('P921').setValue(2);
  eno.getRange('Q921').setValue(4);

  // ЕМО 01.07–05.07: приводим legacy P/Q к текущей логике Callibri.
  emo.getRange('P486').setValue('Callibri: Спам');
  emo.getRange('Q486').setValue('Callibri: Нецелевой_Лид');
  emo.getRange('P487:P504').setValues(Array.from({ length: 18 }, function() { return ['-']; }));

  SpreadsheetApp.flush();

  const validation = tdmStrictValidateDbEnoEmoGoals20260707();
  if (!validation.ok) throw new Error('После фикса сверка не сошлась: ' + JSON.stringify(validation));

  return {
    ok: true,
    changedCells: ['ЕНО!P921', 'ЕНО!Q921', 'ЕМО!P486', 'ЕМО!Q486', 'ЕМО!P487:P504'],
    validation: validation
  };
}

/**
 * ONE_TIME_FIX_2026_07_07_REGION_REPORT_0105
 * Приводит лист "Отчет  регионы" к текущему ДБ за 01.07–05.07.2026.
 * Не трогает показы/клики/расход по регионам. Правит только лиды, доли, CR/CPA и поясняющие комментарии.
 */
function tdmRegionReportMoney20260707_(value) {
  const s = String(value || '').replace(/\s/g, '').replace(/^р\.?/i, '').replace(/₽/g, '').replace(',', '.').replace(/[^0-9.-]/g, '');
  return s && s !== '-' ? Number(s) || 0 : 0;
}

function tdmInstallDynamicReportComments20260707() {
  throw new Error('ARCHIVED_ONE_TIME_FIX_DISABLED: tdmInstallDynamicReportComments20260707 отключена после аудита gpt_webapp_runner. Не запускать из runtime.');
  const ss = SpreadsheetApp.openById('1zXDDfiRYHJE34iOJg-QfAC0bvoTKMIAuVMfsJS0z5Dg');
  const region = ss.getSheetByName('Отчет  регионы');
  const eno = ss.getSheetByName('ЕНО');
  const emo = ss.getSheetByName('ЕМО');
  if (!region || !eno || !emo) throw new Error('Не найдены нужные листы: Отчет  регионы / ЕНО / ЕМО');

  region.getRange('I6').setValue(6);
  region.getRange('I7').setValue(4);
  region.getRange('N6').setFormula('="Основной регион: "&TEXT(I6;"# ##0")&" лидов, CPA "&TEXT(L6;"# ##0,00 ₽")&". Оставить в приоритете, отдельно смотреть качество лидов."');
  region.getRange('N7').setFormula('=TEXT(I7;"# ##0")&" лидов, CPA "&TEXT(L7;"# ##0,00 ₽")&". Проверить кампании RF с расходом без лидов и посадочные."');
  region.getRange('N8').setFormula('=TEXT(I8;"# ##0")&" лидов, CPA "&TEXT(L8;"# ##0,00 ₽")&". Нужна отдельная проверка B2B-качества обращений и запросов."');
  region.getRange('N9').setFormula('=TEXT(I9;"# ##0")&" лидов, CPA "&TEXT(L9;"# ##0,00 ₽")&". Держать в тесте, не масштабировать резко без проверки качества."');
  region.getRange('N21').setFormula('="Итого сходится с ЕМО/ЕНО за 01.07–05.07: "&TEXT(B21;"# ##0")&" показов, "&TEXT(C21;"# ##0")&" кликов, "&TEXT(E21;"# ##0,00 ₽")&", "&TEXT(I21;"# ##0")&" целевых лидов/конверсий, CPA "&TEXT(L21;"# ##0,00 ₽")&"."');
  region.getRange('A23').setFormula('="Вывод: отчёт пересобран по региону из названия РК. "&A6&" даёт "&TEXT(I6;"# ##0")&" лидов при CPA "&TEXT(L6;"# ##0,00 ₽")&"; "&A7&" — "&TEXT(I7;"# ##0")&" лидов при CPA "&TEXT(L7;"# ##0,00 ₽")&"; "&A8&" — "&TEXT(I8;"# ##0")&" лидов при CPA "&TEXT(L8;"# ##0,00 ₽")&", качество нужно проверять отдельно; "&A9&" — "&TEXT(I9;"# ##0")&" лидов при CPA "&TEXT(L9;"# ##0,00 ₽")&". Резких отключений по малой статистике не делаем."');

  eno.getRange('A944').setFormula('="Всего потрачено за неделю "&TEXT(G939;"# ##0,00 ₽")&". Показы — "&TEXT(C939;"# ##0")&", клики — "&TEXT(D939;"# ##0")&", средний CPC — "&TEXT(F939;"# ##0,00 ₽")&"."');
  eno.getRange('A947').setFormula('="Просмотр 3х страниц — "&TEXT(M939;"# ##0")');
  eno.getRange('A948').setFormula('="Callibri: Спам — "&TEXT(P939;"# ##0")');
  eno.getRange('A949').setFormula('="Callibri: Нецелевой_Лид — "&TEXT(Q939;"# ##0")');
  eno.getRange('A950').setFormula('="Ecommerce: добавление в корзину — "&TEXT(R939;"# ##0")');
  eno.getRange('A951').setFormula('="Ecommerce: покупка — "&TEXT(S939;"# ##0")');
  eno.getRange('A952').setFormula('="Jivo-чат — "&TEXT(T939;"# ##0")');
  eno.getRange('A953').setFormula('="Callibri: Лид_Квал_A — "&TEXT(U939;"# ##0")');
  eno.getRange('A954').setFormula('="Callibri: Лид_Квал_C — "&TEXT(V939;"# ##0")');
  eno.getRange('A966').setFormula('="Неделя 29.06–05.07 дала "&TEXT(W939;"# ##0")&" целевых лидов/конверсий по ЕНО при CPA "&TEXT(Y939;"# ##0,00 ₽")&". C-лиды, спам и нецелевые обращения учитываются отдельно и не передаются в целевую оптимизацию Директа. Основной фокус — удержать рабочие связки, проверить кампании с расходом без лидов и отдельно контролировать качество первичных Callibri/Jivo."');

  emo.getRange('A254').setFormula('="Всего задействовано за период "&TEXT(G248;"# ##0,00 ₽")&". Показы — "&TEXT(C248;"# ##0")&", клики — "&TEXT(D248;"# ##0")&", средний CPC — "&TEXT(F248;"# ##0,00 ₽")&"."');
  emo.getRange('A261').setFormula('="Просмотр 3х страниц — "&TEXT(M248;"# ##0")');
  emo.getRange('A262').setFormula('="Callibri: Спам — "&TEXT(P248;"# ##0")');
  emo.getRange('A263').setFormula('="Callibri: Нецелевой_Лид — "&TEXT(Q248;"# ##0")');
  emo.getRange('A264').setFormula('="Ecommerce: добавление в корзину — "&TEXT(R248;"# ##0")');
  emo.getRange('A265').setFormula('="Ecommerce: покупка — "&TEXT(S248;"# ##0")');
  emo.getRange('A266').setFormula('="Jivo-чат — "&TEXT(T248;"# ##0")');
  emo.getRange('A267').setFormula('="Callibri: Лид_Квал_A — "&TEXT(U248;"# ##0")');
  emo.getRange('A268').setFormula('="Callibri: Лид_Квал_C — "&TEXT(V248;"# ##0")');
  emo.getRange('A269').setFormula('="Итого целевых лидов/конверсий — "&TEXT(W248;"# ##0")&", CR — "&TEXT(X248;"0,00 %")&", CPA — "&TEXT(Y248;"# ##0,00 ₽")&". C-лиды учтены отдельно и не входят в целевой факт для Директа."');
  emo.getRange('A283').setFormula('="За 01.07–05.07 получено "&TEXT(W248;"# ##0")&" целевых лидов/конверсий при среднем CPA "&TEXT(Y248;"# ##0,00 ₽")&". C-лиды, спам и нецелевые обращения учитываются отдельно и не передаются в целевую оптимизацию Директа. Резких отключений по малой статистике не делаем."');

  SpreadsheetApp.flush();
  return { ok: true, rule: 'Комментарии текущих блоков ЕНО/ЕМО/регионов переведены на формулы от итоговых строк.', controls: { regionLeads: region.getRange('I21').getDisplayValue(), enoLeads: eno.getRange('W939').getDisplayValue(), emoLeads: emo.getRange('W248').getDisplayValue() } };
}

function tdmFixRegionReportJuly0105To15Leads20260707() {
  throw new Error('ARCHIVED_ONE_TIME_FIX_DISABLED: tdmFixRegionReportJuly0105To15Leads20260707 отключена после аудита gpt_webapp_runner. Не запускать из runtime.');
  const ss = SpreadsheetApp.openById('1zXDDfiRYHJE34iOJg-QfAC0bvoTKMIAuVMfsJS0z5Dg');
  const report = ss.getSheetByName('Отчет  регионы');
  if (!report) throw new Error('Лист Отчет  регионы не найден');

  report.getRange('I6').setValue(6);
  report.getRange('I7').setValue(4);
  report.getRange('N6').setValue('Основной качественный регион: 6 лидов, высокий объём кликов. Оставить в приоритете, отдельно смотреть качество лидов.');
  report.getRange('N7').setValue('4 лида при контролируемом CPA. Проверить кампании RF с расходом без лидов и посадочные.');
  report.getRange('N21').setValue('Итого сходится с ЕМО/ЕНО за 01.07–05.07: 67 846 показов, 1 720 кликов, 40 840,22 ₽, 15 целевых лидов/конверсий.');
  report.getRange('A23').setValue('Вывод: отчёт пересобран по региону из названия РК. СПБ даёт 6 лидов при CPA 2 620,63 ₽; РФ — 4 лида при CPA 3 433,69 ₽; Казахстан — 3 лида при CPA 1 950,98 ₽, качество нужно проверять отдельно; СПБ+РФ — 2 лида при CPA 2 764,38 ₽. Резких отключений по малой статистике не делаем.');

  SpreadsheetApp.flush();

  const control = {
    spbLeads: Number(report.getRange('I6').getValue() || 0),
    rfLeads: Number(report.getRange('I7').getValue() || 0),
    totalLeads: Number(report.getRange('I21').getValue() || 0),
    totalCpa: report.getRange('L21').getDisplayValue(),
    spbCpa: report.getRange('L6').getDisplayValue(),
    rfCpa: report.getRange('L7').getDisplayValue()
  };
  if (control.spbLeads !== 6 || control.rfLeads !== 4 || control.totalLeads !== 15) {
    throw new Error('Контроль регионального отчёта не сошёлся: ' + JSON.stringify(control));
  }
  return { ok: true, changedCells: ['I6','I7','N6','N7','N21','A23'], control: control };
}

function tdmFixRegionReportJuly0105FromDb20260707() {
  throw new Error('ARCHIVED_ONE_TIME_FIX_DISABLED: tdmFixRegionReportJuly0105FromDb20260707 отключена после рефакторинга. Не запускать из runtime.');
  const ss = SpreadsheetApp.openById('1zXDDfiRYHJE34iOJg-QfAC0bvoTKMIAuVMfsJS0z5Dg');
  const emo = ss.getSheetByName('ЕМО');
  const report = ss.getSheetByName('Отчет  регионы');
  if (!emo || !report) throw new Error('Не найдены листы ЕМО / Отчет  регионы');

  function regionFromCampaign_(campaignName) {
    const c = tdmGptNorm_(campaignName).replace(/-/g, '_');
    if (c.indexOf('spb_rf') !== -1 || c.indexOf('rf_spb') !== -1) return 'СПБ+РФ';
    if (/(^|_)kz($|_)/.test(c)) return 'Казахстан';
    if (/(^|_)spb($|_)/.test(c)) return 'СПБ';
    if (/(^|_)rf($|_)/.test(c)) return 'РФ';
    return 'Прочие';
  }

  function grade_(region, r, cr) {
    if (r.leads >= 4) return 'Хорошо';
    if (r.leads >= 2 && cr >= 0.01) return 'Хорошо/контроль';
    if (r.leads >= 2) return 'Средне';
    if (r.clicks >= 100 && r.leads === 0) return 'Проверить';
    return 'Мало данных';
  }

  function comment_(region) {
    if (region === 'СПБ') return 'Основной качественный регион: 4 лида, высокий объём кликов. Оставить в приоритете, отдельно смотреть качество лидов.';
    if (region === 'РФ') return '3 лида при самом высоком CPA среди групп. Проверить кампании RF с расходом без лидов и посадочные.';
    if (region === 'Казахстан') return '3 лида при низком CPA, но большой объём дешёвого трафика. Нужна отдельная проверка B2B-качества обращений и запросов.';
    if (region === 'СПБ+РФ') return '2 лида и лучший CR по срезу. Держать в тесте, не масштабировать резко без проверки качества.';
    return 'Прочие кампании без отдельной региональной метки.';
  }

  const rows = emo.getRange('A486:X503').getDisplayValues();
  const order = ['СПБ', 'РФ', 'Казахстан', 'СПБ+РФ'];
  const byRegion = {};
  order.forEach(function(region) {
    byRegion[region] = { region: region, impressions: 0, clicks: 0, cost: 0, threePages: 0, leads: 0, bounceWeighted: 0, depthWeighted: 0 };
  });

  rows.forEach(function(row) {
    const campaign = String(row[0] || '').trim();
    if (!campaign || campaign.toUpperCase() === 'ИТОГО') return;
    const region = regionFromCampaign_(campaign);
    if (!byRegion[region]) byRegion[region] = { region: region, impressions: 0, clicks: 0, cost: 0, threePages: 0, leads: 0, bounceWeighted: 0, depthWeighted: 0 };

    const clicks = tdmStrictNum20260707_(row[3]);
    const bounce = tdmStrictNum20260707_(row[10]);
    const depth = tdmStrictNum20260707_(row[11]);

    byRegion[region].impressions += tdmStrictNum20260707_(row[2]);
    byRegion[region].clicks += clicks;
    byRegion[region].cost += tdmRegionReportMoney20260707_(row[6]);
    byRegion[region].threePages += tdmStrictNum20260707_(row[12]);
    byRegion[region].leads += tdmStrictNum20260707_(row[23]);
    byRegion[region].bounceWeighted += bounce * clicks;
    byRegion[region].depthWeighted += depth * clicks;
  });

  const totals = order.reduce(function(acc, region) {
    const r = byRegion[region];
    acc.impressions += r.impressions;
    acc.clicks += r.clicks;
    acc.cost += r.cost;
    acc.threePages += r.threePages;
    acc.leads += r.leads;
    return acc;
  }, { impressions: 0, clicks: 0, cost: 0, threePages: 0, leads: 0 });
  totals.cost = Math.round(totals.cost * 100) / 100;

  const output = [];
  order.forEach(function(region) {
    const r = byRegion[region];
    const rowNumber = 6 + output.length;
    const bounce = r.clicks ? r.bounceWeighted / r.clicks : 0;
    const depth = r.clicks ? r.depthWeighted / r.clicks : 0;
    const share = totals.leads ? r.leads / totals.leads : 0;
    const cr = r.clicks ? r.leads / r.clicks : 0;
    const cpa = r.leads ? r.cost / r.leads : 0;
    output.push([region, r.impressions, r.clicks, '=IFERROR(C' + rowNumber + '/B' + rowNumber + ';0)', Math.round(r.cost * 100) / 100, r.threePages, bounce, depth, r.leads, share, cr, cpa, grade_(region, r, cr), comment_(region)]);
  });

  while (output.length < 15) {
    const rowNumber = 6 + output.length;
    output.push(['', 0, 0, '=IFERROR(C' + rowNumber + '/B' + rowNumber + ';0)', 0, 0, '', '', 0, 0, 0, 0, '', '']);
  }

  report.getRange('A1').setValue('Отчёт по регионам из названия РК — июль 2026 (01.07–05.07)');
  report.getRange('A2').setValue('Логика: лист оставлен как июльский срез 01.07–05.07. Регион берётся из названия РК: _spb → СПБ, _rf → РФ, _kz → Казахстан, spb-rf/rf_spb → СПБ+РФ. Итоги сверяются с ДБ/ЕМО за 01.07–05.07, не с полной неделей 29.06–05.07.');
  report.getRange('F5').setValue('Просмотр 3х страниц');
  report.getRange('A6:N20').setValues(output);

  report.getRange('B6:C21').setNumberFormat('#,##0');
  report.getRange('D6:D21').setNumberFormat('0.00%');
  report.getRange('E6:E21').setNumberFormat('#,##0.00 ₽');
  report.getRange('F6:F21').setNumberFormat('#,##0');
  report.getRange('G6:G20').setNumberFormat('0.00%');
  report.getRange('H6:H20').setNumberFormat('0.00');
  report.getRange('I6:I21').setNumberFormat('#,##0');
  report.getRange('J6:J21').setNumberFormat('0.00%');
  report.getRange('K6:K21').setNumberFormat('0.00%');
  report.getRange('L6:L21').setNumberFormat('#,##0.00 ₽');

  report.getRange('A21').setValue('Итого');
  report.getRange('B21').setFormula('=SUM(B6:B20)');
  report.getRange('C21').setFormula('=SUM(C6:C20)');
  report.getRange('D21').setFormula('=IFERROR(C21/B21;0)');
  report.getRange('E21').setFormula('=SUM(E6:E20)');
  report.getRange('F21').setFormula('=SUM(F6:F20)');
  report.getRange('I21').setFormula('=SUM(I6:I20)');
  report.getRange('J21').setFormula('=IFERROR(I21/I21;0)');
  report.getRange('K21').setFormula('=IFERROR(I21/C21;0)');
  report.getRange('L21').setFormula('=IFERROR(E21/I21;0)');
  report.getRange('N21').setValue('Итого сходится с ДБ/ЕМО за 01.07–05.07: 67 846 показов, 1 720 кликов, 40 840,22 ₽, 12 лидов.');
  report.getRange('A23').setValue('Вывод: отчёт пересобран по региону из названия РК. СПБ даёт 4 лида при CPA 3 930,94 ₽; РФ — 3 лида при CPA 4 578,26 ₽; Казахстан — 3 лида при низком CPA, но качество нужно проверять отдельно; СПБ+РФ — 2 лида и лучший CR по срезу. Резких отключений по малой статистике не делаем.');

  SpreadsheetApp.flush();

  const control = {
    reportImpressions: tdmStrictNum20260707_(report.getRange('B21').getDisplayValue()),
    reportClicks: tdmStrictNum20260707_(report.getRange('C21').getDisplayValue()),
    reportCost: Math.round(tdmRegionReportMoney20260707_(report.getRange('E21').getDisplayValue()) * 100) / 100,
    reportLeads: tdmStrictNum20260707_(report.getRange('I21').getDisplayValue())
  };
  if (control.reportImpressions !== 67846 || control.reportClicks !== 1720 || Math.abs(control.reportCost - 40840.22) > 1 || control.reportLeads !== 12) {
    throw new Error('Отчет регионы после пересборки по названию РК не сошёлся: ' + JSON.stringify(control));
  }

  return {
    ok: true,
    period: '2026-07-01—2026-07-05',
    source: 'ЕМО campaign rows + region parsed from CampaignName',
    changedRanges: ['Отчет  регионы!A1:A2', 'Отчет  регионы!F5', 'Отчет  регионы!A6:N21', 'Отчет  регионы!A23'],
    control: control
  };
}

/**
 * ONE_TIME_REBUILD_2026_07_07_REGION_FROM_CAMPAIGN_NAME
 * Пересобирает "Отчет  регионы" за 01.07–05.07.2026 по региону из названия РК.
 * Источник факта — текущий блок ЕМО 01.07–05.07, чтобы лиды сходились с ДБ/ЕМО.
 */
function tdmRebuildRegionReportFromCampaignName20260707() {
  throw new Error('ARCHIVED_ONE_TIME_FIX_DISABLED: tdmRebuildRegionReportFromCampaignName20260707 отключена после рефакторинга. Не запускать из runtime.');
  const ss = SpreadsheetApp.openById('1zXDDfiRYHJE34iOJg-QfAC0bvoTKMIAuVMfsJS0z5Dg');
  const emo = ss.getSheetByName('ЕМО');
  const report = ss.getSheetByName('Отчет  регионы');
  if (!emo || !report) throw new Error('Не найдены листы ЕМО / Отчет  регионы');

  const rows = emo.getRange('A486:X503').getDisplayValues();
  const order = ['СПБ', 'РФ', 'Казахстан', 'СПБ+РФ'];
  const byRegion = {};
  order.forEach(function(region) {
    byRegion[region] = { region: region, impressions: 0, clicks: 0, cost: 0, threePages: 0, leads: 0, bounceWeighted: 0, depthWeighted: 0, campaigns: [] };
  });

  rows.forEach(function(row) {
    const campaign = String(row[0] || '').trim();
    if (!campaign || campaign.toUpperCase() === 'ИТОГО') return;
    const region = tdmRegionFromCampaignName20260707_(campaign);
    if (!byRegion[region]) byRegion[region] = { region: region, impressions: 0, clicks: 0, cost: 0, threePages: 0, leads: 0, bounceWeighted: 0, depthWeighted: 0, campaigns: [] };

    const clicks = tdmStrictNum20260707_(row[3]);
    const bounce = tdmStrictNum20260707_(row[10]);
    const depth = tdmStrictNum20260707_(row[11]);

    byRegion[region].impressions += tdmStrictNum20260707_(row[2]);
    byRegion[region].clicks += clicks;
    byRegion[region].cost += tdmRegionReportMoney20260707_(row[6]);
    byRegion[region].threePages += tdmStrictNum20260707_(row[12]);
    byRegion[region].leads += tdmStrictNum20260707_(row[23]);
    byRegion[region].bounceWeighted += bounce * clicks;
    byRegion[region].depthWeighted += depth * clicks;
    byRegion[region].campaigns.push(campaign);
  });

  const totals = order.reduce(function(acc, region) {
    const r = byRegion[region];
    acc.impressions += r.impressions;
    acc.clicks += r.clicks;
    acc.cost += r.cost;
    acc.threePages += r.threePages;
    acc.leads += r.leads;
    return acc;
  }, { impressions: 0, clicks: 0, cost: 0, threePages: 0, leads: 0 });
  totals.cost = Math.round(totals.cost * 100) / 100;

  const output = [];
  order.forEach(function(region) {
    const r = byRegion[region];
    const bounce = r.clicks ? r.bounceWeighted / r.clicks : 0;
    const depth = r.clicks ? r.depthWeighted / r.clicks : 0;
    const share = totals.leads ? r.leads / totals.leads : 0;
    const cr = r.clicks ? r.leads / r.clicks : 0;
    const cpa = r.leads ? r.cost / r.leads : 0;
    output.push([
      region,
      r.impressions,
      r.clicks,
      '=IFERROR(C' + (6 + output.length) + '/B' + (6 + output.length) + ';0)',
      Math.round(r.cost * 100) / 100,
      r.threePages,
      bounce,
      depth,
      r.leads,
      share,
      cr,
      cpa,
      tdmRegionCampaignGrade20260707_(region, r, cr, cpa, bounce, depth),
      tdmRegionCampaignComment20260707_(region, r, cr, cpa, bounce, depth)
    ]);
  });

  while (output.length < 15) {
    const rowNumber = 6 + output.length;
    output.push(['', 0, 0, '=IFERROR(C' + rowNumber + '/B' + rowNumber + ';0)', 0, 0, '', '', 0, 0, 0, 0, '', '']);
  }

  report.getRange('A1').setValue('Отчёт по регионам из названия РК — июль 2026 (01.07–05.07)');
  report.getRange('A2').setValue('Логика: лист оставлен как июльский срез 01.07–05.07. Регион берётся из названия РК: _spb → СПБ, _rf → РФ, _kz → Казахстан, spb-rf/rf_spb → СПБ+РФ. Итоги сверяются с ДБ/ЕМО за 01.07–05.07, не с полной неделей 29.06–05.07.');
  report.getRange('F5').setValue('Просмотр 3х страниц');
  report.getRange('A6:N20').setValues(output);

  report.getRange('B6:C21').setNumberFormat('#,##0');
  report.getRange('D6:D21').setNumberFormat('0.00%');
  report.getRange('E6:E21').setNumberFormat('#,##0.00 ₽');
  report.getRange('F6:F21').setNumberFormat('#,##0');
  report.getRange('G6:G20').setNumberFormat('0.00%');
  report.getRange('H6:H20').setNumberFormat('0.00');
  report.getRange('I6:I21').setNumberFormat('#,##0');
  report.getRange('J6:J21').setNumberFormat('0.00%');
  report.getRange('K6:K21').setNumberFormat('0.00%');
  report.getRange('L6:L21').setNumberFormat('#,##0.00 ₽');

  report.getRange('A21').setValue('Итого');
  report.getRange('B21').setFormula('=SUM(B6:B20)');
  report.getRange('C21').setFormula('=SUM(C6:C20)');
  report.getRange('D21').setFormula('=IFERROR(C21/B21;0)');
  report.getRange('E21').setFormula('=SUM(E6:E20)');
  report.getRange('F21').setFormula('=SUM(F6:F20)');
  report.getRange('I21').setFormula('=SUM(I6:I20)');
  report.getRange('J21').setFormula('=IFERROR(I21/I21;0)');
  report.getRange('K21').setFormula('=IFERROR(I21/C21;0)');
  report.getRange('L21').setFormula('=IFERROR(E21/I21;0)');
  report.getRange('N21').setValue('Итого сходится с ДБ/ЕМО за 01.07–05.07: 67 846 показов, 1 720 кликов, 40 840,22 ₽, 12 лидов.');
  report.getRange('A23').setValue('Вывод: отчёт пересобран по региону из названия РК. СПБ даёт 4 лида при CPA 3 930,94 ₽; РФ — 3 лида при CPA 4 578,26 ₽; Казахстан — 3 лида при низком CPA, но качество нужно проверять отдельно; СПБ+РФ — 2 лида и лучший CR по срезу. Резких отключений по малой статистике не делаем.');

  SpreadsheetApp.flush();

  const control = {
    impressions: tdmStrictNum20260707_(report.getRange('B21').getDisplayValue()),
    clicks: tdmStrictNum20260707_(report.getRange('C21').getDisplayValue()),
    cost: Math.round(tdmRegionReportMoney20260707_(report.getRange('E21').getDisplayValue()) * 100) / 100,
    leads: tdmStrictNum20260707_(report.getRange('I21').getDisplayValue())
  };
  if (control.impressions !== 67846 || control.clicks !== 1720 || Math.abs(control.cost - 40840.22) > 1 || control.leads !== 12) {
    throw new Error('Контроль после пересборки не сошёлся: ' + JSON.stringify(control));
  }

  return {
    ok: true,
    period: '2026-07-01—2026-07-05',
    source: 'ЕМО campaign rows + region parsed from CampaignName',
    changedRanges: ['Отчет  регионы!A1:A2', 'Отчет  регионы!F5', 'Отчет  регионы!A6:N21', 'Отчет  регионы!A23'],
    control: control
  };
}

function tdmRegionFromCampaignName20260707_(campaignName) {
  const c = tdmGptNorm_(campaignName).replace(/-/g, '_');
  if (c.indexOf('spb_rf') !== -1 || c.indexOf('rf_spb') !== -1) return 'СПБ+РФ';
  if (/(^|_)kz($|_)/.test(c)) return 'Казахстан';
  if (/(^|_)spb($|_)/.test(c)) return 'СПБ';
  if (/(^|_)rf($|_)/.test(c)) return 'РФ';
  return 'Прочие';
}

function tdmRegionCampaignGrade20260707_(region, r, cr, cpa, bounce, depth) {
  if (r.leads >= 4) return 'Хорошо';
  if (r.leads >= 2 && cr >= 0.01) return 'Хорошо/контроль';
  if (r.leads >= 2) return 'Средне';
  if (r.clicks >= 100 && r.leads === 0) return 'Проверить';
  return 'Мало данных';
}

function tdmRegionCampaignComment20260707_(region, r, cr, cpa, bounce, depth) {
  if (region === 'СПБ') return 'Основной качественный регион: 4 лида, высокий объём кликов. Оставить в приоритете, отдельно смотреть качество лидов.';
  if (region === 'РФ') return '3 лида при самом высоком CPA среди групп. Проверить кампании RF с расходом без лидов и посадочные.';
  if (region === 'Казахстан') return '3 лида при низком CPA, но большой объём дешёвого трафика. Нужна отдельная проверка B2B-качества обращений и запросов.';
  if (region === 'СПБ+РФ') return '2 лида и лучший CR по срезу. Держать в тесте, не масштабировать резко без проверки качества.';
  return 'Прочие кампании без отдельной региональной метки.';
}

function tdmFixEnoRegionCommentRefSafe20260708() {
  throw new Error('ARCHIVED_ONE_TIME_FIX_DISABLED: tdmFixEnoRegionCommentRefSafe20260708 уже выполнена и отключена после исправления ЕНО!A963:C967.');
  const ss = SpreadsheetApp.openById('1zXDDfiRYHJE34iOJg-QfAC0bvoTKMIAuVMfsJS0z5Dg');
  const eno = ss.getSheetByName('ЕНО');

  if (!eno) throw new Error('Лист ЕНО не найден');

  const values = [
    ['СПБ — 7 690 показов, 317 кликов, 15 723,77 ₽, 4 лида, CPA 3 930,94 ₽.'],
    ['РФ — 15 054 показов, 324 клика, 13 734,77 ₽, 3 лида, CPA 4 578,26 ₽.'],
    ['Казахстан — 29 363 показа, 1 003 клика, 5 852,93 ₽, 3 лида, CPA 1 950,98 ₽. Качество проверить отдельно.'],
    ['СПБ+РФ — 15 739 показов, 76 кликов, 5 528,75 ₽, 2 лида, CPA 2 764,38 ₽.'],
    ['Итого — 67 846 показов, 1 720 кликов, 40 840,22 ₽, 12 лидов, CPA 3 403,35 ₽.']
  ];

  eno.getRange('A963:A967').setValues(values).setWrap(true);
  eno.getRange('C963:C967').setFormulas([['=A963'], ['=A964'], ['=A965'], ['=A966'], ['=A967']]).setWrap(true);
  SpreadsheetApp.flush();

  const check = eno.getRange('A963:C967').getDisplayValues();
  const joined = check.map(row => row.join(' ')).join('\n');

  if (joined.indexOf('#REF!') !== -1) {
    throw new Error('После исправления в ЕНО!A963:C967 остался #REF!: ' + joined);
  }

  return {
    ok: true,
    changedRanges: ['ЕНО!A963:A967', 'ЕНО!C963:C967'],
    rows: check.length
  };
}

function tdmFixEnoRegionComment963_20260707() {
  throw new Error('ARCHIVED_ONE_TIME_FIX_DISABLED: tdmFixEnoRegionComment963_20260707 отключена после рефакторинга. Не запускать из runtime.');
  const ss = SpreadsheetApp.openById('1zXDDfiRYHJE34iOJg-QfAC0bvoTKMIAuVMfsJS0z5Dg');
  const eno = ss.getSheetByName('ЕНО');
  if (!eno) throw new Error('Лист ЕНО не найден');
  const text = 'СПБ — 7 690 показов, 317 кликов, 15 723,77 ₽, 4 лида, CPA 3 930,94 ₽.\n\nРФ — 15 054 показов, 324 клика, 13 734,77 ₽, 3 лида, CPA 4 578,26 ₽.\n\nКазахстан — 29 363 показа, 1 003 клика, 5 852,93 ₽, 3 лида, CPA 1 950,98 ₽.\n\nСПБ+РФ — 15 739 показов, 76 кликов, 5 528,75 ₽, 2 лида, CPA 2 764,38 ₽.\n\nИтого — 67 846 показов, 1 720 кликов, 40 840,22 ₽, 12 лидов, CPA 3 403,35 ₽.';
  eno.getRange('A963').setValue(text).setWrap(true);
  SpreadsheetApp.flush();
  const check = eno.getRange('A963').getDisplayValue();
  if (check !== text) throw new Error('ЕНО!A963 не обновилась: ' + check);
  return { ok: true, changedCell: 'ЕНО!A963', value: check };
}

function tdmFixEmoCommentsRegionCleanup20260707() {
  throw new Error('ARCHIVED_ONE_TIME_FIX_DISABLED: tdmFixEmoCommentsRegionCleanup20260707 отключена после рефакторинга. Не запускать из runtime.');
  const ss = SpreadsheetApp.openById('1zXDDfiRYHJE34iOJg-QfAC0bvoTKMIAuVMfsJS0z5Dg');
  const emo = ss.getSheetByName('ЕМО');
  if (!emo) throw new Error('Лист ЕМО не найден');

  // Правильный июльский блок уже находится в строках 227:289.
  // Всё ниже 289 строки — дубли июль/апрель, которые появились после предыдущих правок.
  emo.getRange('A291:AB650').clearContent();
  SpreadsheetApp.flush();

  const keep = emo.getRange('A227:A289').getDisplayValues().map(function(row) { return row[0]; }).join('\n');
  const tail = emo.getRange('A291:A650').getDisplayValues().map(function(row) { return row[0]; }).join('\n');

  if (keep.indexOf('Отчет месячный по РК 1.07. - 5.07.2026') === -1 || keep.indexOf('Итого — 67 846 показов, 1 720 кликов, 40 840,22 ₽, 12 лидов, CPA 3 403,35 ₽.') === -1) {
    throw new Error('Правильный июльский блок ЕМО не найден в A227:A289');
  }
  if (tail.indexOf('Отчет месячный по РК 1.07') !== -1 || tail.indexOf('Комментарии к отчёту за апрель') !== -1 || tail.indexOf('13 лидов') !== -1) {
    throw new Error('Дубли ниже 289 строки не очищены');
  }

  return { ok: true, keptRange: 'ЕМО!A227:AB289', clearedRange: 'ЕМО!A291:AB650' };
}

/**
 * ONE_TIME_FIX_2026_07_07_DB_CALLIBRI_PRIMARY_STATUSES
 * ДБ: 29.06–05.07.2026. Правит только Callibri-статусы, отправленные клиентом, и формулу факта без C-лидов.
 * Только первичные обращения: в выгрузке клиента у всех учтённых строк № = 1.
 */
function tdmFixDbCallibriPrimaryStatuses20260707() {
  throw new Error('ARCHIVED_ONE_TIME_FIX_DISABLED: tdmFixDbCallibriPrimaryStatuses20260707 отключена после рефакторинга. Не запускать из runtime.');
  const ss = SpreadsheetApp.openById('1zXDDfiRYHJE34iOJg-QfAC0bvoTKMIAuVMfsJS0z5Dg');
  const db = ss.getSheetByName('ДБ');
  if (!db) throw new Error('Лист ДБ не найден');

  const rows = [
    { row: 264, date: '29.06.2026', spam: 1, nonTarget: 0, leadA: 4, leadC: 2 },
    { row: 265, date: '30.06.2026', spam: 0, nonTarget: 0, leadA: 0, leadC: 3 },
    { row: 290, date: '01.07.2026', spam: 0, nonTarget: 0, leadA: 3, leadC: 1 },
    { row: 291, date: '02.07.2026', spam: 0, nonTarget: 0, leadA: 3, leadC: 4 },
    { row: 292, date: '03.07.2026', spam: 0, nonTarget: 2, leadA: 0, leadC: 3 },
    { row: 293, date: '04.07.2026', spam: 0, nonTarget: 0, leadA: 0, leadC: 0 },
    { row: 294, date: '05.07.2026', spam: 0, nonTarget: 0, leadA: 0, leadC: 0 }
  ];

  rows.forEach(function(x) {
    if (String(db.getRange('B' + x.row).getDisplayValue()).trim() !== x.date) {
      throw new Error('Дата не совпала в ДБ!' + x.row + ': ожидали ' + x.date + ', получили ' + db.getRange('B' + x.row).getDisplayValue());
    }
    db.getRange('M' + x.row).setValue(x.spam || '-');
    db.getRange('N' + x.row).setValue(x.nonTarget || '-');
    db.getRange('R' + x.row).setValue(x.leadA || '-');
    db.getRange('S' + x.row).setValue(x.leadC || '-');
    db.getRange('T' + x.row).setFormula('=SUM(P' + x.row + ':R' + x.row + ')');
  });

  SpreadsheetApp.flush();

  const control = rows.reduce(function(acc, x) {
    acc.spam += Number(x.spam || 0);
    acc.nonTarget += Number(x.nonTarget || 0);
    acc.leadA += Number(x.leadA || 0);
    acc.leadC += Number(x.leadC || 0);
    return acc;
  }, { spam: 0, nonTarget: 0, leadA: 0, leadC: 0 });

  const fact = rows.reduce(function(sum, x) {
    const v = String(db.getRange('T' + x.row).getDisplayValue() || '').replace(/\s/g, '').replace(',', '.');
    return sum + (Number(v) || 0);
  }, 0);

  return {
    ok: true,
    period: '2026-06-29—2026-07-05',
    changedRanges: ['ДБ!M264:N265', 'ДБ!R264:T265', 'ДБ!M290:N294', 'ДБ!R290:T294'],
    rule: 'Только первичные Callibri-статусы. C/Спам/Нецелевой учитываются отдельно, но не входят в Факт сумма конверсий.',
    control: Object.assign(control, { factWithoutLeadC: fact })
  };
}

/**
 * SAFE_FIX_2026_07_07_DB_T295_FORMULA
 * Одноразово исправляет только формулу ДБ!T295: факт без Callibri: Лид_Квал_C.
 */
function tdmSafeFixDbT295Formula20260707() {
  throw new Error('ARCHIVED_ONE_TIME_FIX_DISABLED: tdmSafeFixDbT295Formula20260707 уже выполнена и отключена после рефакторинга.');
  const spreadsheetId = '1zXDDfiRYHJE34iOJg-QfAC0bvoTKMIAuVMfsJS0z5Dg';
  const ss = SpreadsheetApp.openById(spreadsheetId);
  const db = ss.getSheetByName('ДБ');
  if (!db) throw new Error('Лист ДБ не найден');

  const expectedDate = '06.07.2026';
  const actualDate = String(db.getRange('B295').getDisplayValue() || '').trim();
  if (actualDate !== expectedDate) {
    throw new Error('ДБ!B295 дата не совпала: ожидали ' + expectedDate + ', получили ' + actualDate);
  }

  db.getRange('T295').setFormula('=SUM(P295:R295)');
  SpreadsheetApp.flush();

  const formula = db.getRange('T295').getFormula();
  const value = db.getRange('T295').getDisplayValue();
  if (formula !== '=SUM(P295:R295)') {
    throw new Error('Формула ДБ!T295 не обновилась: ' + formula);
  }

  return { ok: true, changedCell: 'ДБ!T295', formula: formula, displayValue: value };
}

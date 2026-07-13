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
    .addSeparator()
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
function tdmInspectDirectHelperSignature() {
  return {
    eno: String(tdmeDirectRequest_),
    emo: String(tdmEnoMayMonthlyV3DirectRequest_),
    emoLoad: String(tdmEnoMayMonthlyV3LoadCampaignRows_),
    emoGoals: String(tdmEnoMayMonthlyV3AllGoalIds_),
    emoGoalConfig: TDM_ENO_MAY_MONTHLY_V3_CONFIG.GOALS,
    emoAttribution: TDM_ENO_MAY_MONTHLY_V3_CONFIG.ATTRIBUTION_MODEL,
    regionLeadsFix: String(tdmFixRegionReportJuly0105To15Leads20260707),
    regionDbFix: String(tdmFixRegionReportJuly0105FromDb20260707),
    compactRegionReport: String(tdmBuildCompactRegionReport_)
  };
}

function tdmUpdateRegionsCitiesReport() {
  updateRegionsCitiesReport();
  tdmBuildCompactRegionReport_();
}

function tdmFetchRegionGoalBreakdownDirect_(date1, date2) {
  const ids = {
    view3: 453453800,
    addToCart: 504318736,
    purchase: 504318735,
    jivo: 575188424
  };
  const body = {
    params: {
      SelectionCriteria: { DateFrom: date1, DateTo: date2 },
      Goals: Object.keys(ids).map(function(key) { return ids[key]; }),
      AttributionModels: ['AUTO'],
      FieldNames: ['LocationOfPresenceName', 'Conversions'],
      ReportName: 'TDM region goals ' + date1 + '_' + date2 + '_' + Utilities.getUuid(),
      ReportType: 'CUSTOM_REPORT',
      DateRangeType: 'CUSTOM_DATE',
      Format: 'TSV',
      IncludeVAT: 'YES',
      IncludeDiscount: 'NO'
    }
  };
  const text = tdmEnoMayMonthlyV3DirectRequest_(body);
  const lines = String(text || '').split(/\r?\n/).filter(Boolean);
  if (lines.length < 2) return {};
  const headers = lines[0].split('\t');
  const result = {};
  lines.slice(1).forEach(function(line) {
    const cols = line.split('\t');
    const region = String(cols[0] || '').trim();
    if (!region) return;
    const item = {};
    Object.keys(ids).forEach(function(key) {
      const id = String(ids[key]);
      const index = headers.findIndex(function(header) { return String(header).indexOf(id) >= 0; });
      item[key] = index >= 0 ? Number(String(cols[index] || '0').replace(',', '.')) || 0 : 0;
    });
    result[region] = item;
  });
  return result;
}

function tdmCallibriRegionBucketsFromEmo_() {
  const ss = SpreadsheetApp.openById(TDM_CONFIG.SPREADSHEET_ID);
  const emo = ss.getSheetByName('ЕМО');
  if (!emo) throw new Error('Не найден лист ЕМО');
  const rows = emo.getRange('A230:V247').getValues();
  const result = {
    'Санкт-Петербург': { spam: 0, unqualified: 0, qualA: 0, qualC: 0 },
    'Казахстан': { spam: 0, unqualified: 0, qualA: 0, qualC: 0 },
    'Прочие регионы': { spam: 0, unqualified: 0, qualA: 0, qualC: 0 }
  };
  rows.forEach(function(row) {
    const campaign = String(row[0] || '').toLowerCase().replace(/-/g, '_');
    if (!campaign || campaign === 'итого') return;
    let bucket = 'Прочие регионы';
    if (/(^|_)kz($|_)/.test(campaign)) bucket = 'Казахстан';
    else if (/(^|_)spb($|_)/.test(campaign) && campaign.indexOf('spb_rf') === -1 && campaign.indexOf('rf_spb') === -1) bucket = 'Санкт-Петербург';
    result[bucket].spam += Number(row[15]) || 0;
    result[bucket].unqualified += Number(row[16]) || 0;
    result[bucket].qualA += Number(row[20]) || 0;
    result[bucket].qualC += Number(row[21]) || 0;
  });
  return result;
}

function archived_tdmFetchRegionGoalBreakdownAgentLegacy_(date1, date2) {
  const token = PropertiesService.getScriptProperties().getProperty('METRIKA_TOKEN');
  if (!token) throw new Error('Не найден METRIKA_TOKEN');

  const headers = { Authorization: 'OAuth ' + token };
  const goalsResponse = UrlFetchApp.fetch(
    'https://api-metrika.yandex.net/management/v1/counter/92370926/goals',
    { headers: headers, muteHttpExceptions: true }
  );
  if (goalsResponse.getResponseCode() !== 200) {
    throw new Error('Не удалось получить цели Метрики: ' + goalsResponse.getContentText().slice(0, 300));
  }

  const goals = JSON.parse(goalsResponse.getContentText()).goals || [];
  const wanted = {
    view3: ['просмотр 3х страниц','просмотр 3 страниц'],
    spam: ['callibri: спам','спам'],
    unqualified: ['callibri: нецелевой_лид','нецелевой лид'],
    addToCart: ['ecommerce: добавление в корзину','добавление в корзину'],
    purchase: ['ecommerce: покупка','покупка'],
    jivo: ['jivo-чат','jivo чат'],
    qualA: ['callibri: лид_квал_a','лид квал a'],
    qualC: ['callibri: лид_квал_c','лид квал c']
  };

  const normalizeName = function(value) {
    return String(value || '').toLowerCase().replace(/ё/g, 'е').replace(/[^a-zа-я0-9]+/g, ' ').trim();
  };
  const goalIds = {};
  Object.keys(wanted).forEach(function(key) {
    const goal = goals.find(function(item) {
      const name = normalizeName(item.name);
      return wanted[key].some(function(alias) { return name === normalizeName(alias); });
    });
    if (goal) goalIds[key] = goal.id;
  });

  const keys = Object.keys(goalIds);
  if (!keys.length) return {};

  const metrics = keys.map(function(key) { return 'ym:s:goal' + goalIds[key] + 'reaches'; });
  const url = 'https://api-metrika.yandex.net/stat/v1/data?' + [
    'ids=92370926',
    'date1=' + encodeURIComponent(date1),
    'date2=' + encodeURIComponent(date2),
    'dimensions=ym:s:regionArea',
    'metrics=' + encodeURIComponent(metrics.join(',')),
    'accuracy=full',
    'limit=10000'
  ].join('&');

  const reportResponse = UrlFetchApp.fetch(url, { headers: headers, muteHttpExceptions: true });
  if (reportResponse.getResponseCode() !== 200) {
    throw new Error('Не удалось получить цели по регионам: ' + reportResponse.getContentText().slice(0, 300));
  }

  const rows = JSON.parse(reportResponse.getContentText()).data || [];
  const result = {};
  rows.forEach(function(row) {
    const region = String(row.dimensions && row.dimensions[0] && row.dimensions[0].name || '').trim();
    if (!region) return;
    const item = {};
    keys.forEach(function(key, index) { item[key] = Number(row.metrics[index]) || 0; });
    result[region] = item;
  });
  return result;
}

function tdmBuildCompactRegionReport_() {
  const ss = SpreadsheetApp.openById(TDM_CONFIG.SPREADSHEET_ID);
  const source = ss.getSheetByName('Регионы_города');
  const settings = ss.getSheetByName('Настройки_Регионы');
  const target = ss.getSheetByName('Отчет  регионы');
  const emo = ss.getSheetByName('ЕМО');
  if (!source || !settings || !target || !emo) throw new Error('Не найдены обязательные листы');

  const enabledBase = settings.getRange(3, 1, Math.max(settings.getLastRow() - 2, 1), 3).getValues()
    .filter(r => String(r[0]).trim() && String(r[2]).trim().toLowerCase() === 'да')
    .map(r => String(r[0]).trim())
    .filter(name => name !== 'Санкт-Петербург' && name !== 'Гатчина');
  const enabled = Array.from(new Set(enabledBase.concat([
    'Санкт-Петербург',
    'Ленинградская область',
    'Казахстан',
    'Прочие регионы'
  ])));

  const aliases = {
    'Москва':'Москва и Московская область','Санкт-Петербург':'Санкт-Петербург и Ленинградская область',
    'Архангельск':'Архангельская область','Воронеж':'Воронежская область','Великий Новгород':'Новгородская область',
    'Екатеринбург':'Свердловская область','Вологда':'Вологодская область','Казань':'Республика Татарстан',
    'Краснодар':'Краснодарский край','Калининград':'Калининградская область','Красноярск':'Красноярский край',
    'Мурманск':'Мурманская область','Петрозаводск':'Республика Карелия','Нижний Новгород':'Нижегородская область',
    'Псков':'Псковская область','Новосибирск':'Новосибирская область','Омск':'Омская область',
    'Сыктывкар':'Республика Коми','Пермь':'Пермский край','Самара':'Самарская область',
    'Тюмень':'Тюменская область','Уфа':'Республика Башкортостан','Челябинск':'Челябинская область',
    'Ростов-на Дону':'Ростовская область'
  };

  const values = source.getDataRange().getValues();
  let headerRow = -1;
  for (let i = 0; i < values.length; i++) {
    if (String(values[i][0]).trim() === 'Месяц' && String(values[i][1]).trim() === 'Регион' && String(values[i][5]).trim() === 'Показы') {
      headerRow = i;
      break;
    }
  }
  if (headerRow < 0) throw new Error('Не найдена детальная таблица регионов');

  const normalizeRegion = value => String(value || '').toLowerCase().replace(/ё/g, 'е').replace(/[^a-zа-я0-9]+/g, ' ').trim();
  const selectedByKey = {};
  enabledBase.forEach(name => {
    selectedByKey[normalizeRegion(aliases[name] || name)] = name;
  });

  const leningradPlaces = new Set([
    'кириши','токсово','выборг','кузьмоловский','тосно','сосновый бор','федоровское',
    'новодевяткинское','коммунар','вырица','шушары','всеволожск','парголово','сестрорецк',
    'пушкин','ломоносов','кудрово','мурино','гатчина','колпино','кронштадт','петергоф',
    'сланцы','кингисепп','луга','волхов','тихвин','приозерск','сертолово','отрадное',
    'никольское','кировск','светогорск','подпорожье','пикалево','бокситогорск'
  ]);
  const kazakhstanPlaces = new Set([
    'казахстан','караганда','алматы','усть каменогорск','семей','астана','павлодар','темиртау',
    'костанай','конаев','экибастуз','шымкент','петропавловск','аягоз','рудный','жезказган',
    'актау','актобе','атырау','кокшетау','тараз','талдыкорган','уральск','кызылорда'
  ]);

  const classifyRegion = rawName => {
    const key = normalizeRegion(rawName);
    if (!key || key === 'не определен' || key === 'не определено') return 'Прочие регионы';
    if (key === 'санкт петербург' || key === 'санкт петербург и ленинградская область') return 'Санкт-Петербург';
    if (key === 'ленинградская область' || leningradPlaces.has(key)) return 'Ленинградская область';
    if (kazakhstanPlaces.has(key)) return 'Казахстан';
    const canonicalKey = normalizeRegion(directRegionKey_(rawName));
    if (selectedByKey[canonicalKey]) return selectedByKey[canonicalKey];
    if (selectedByKey[key]) return selectedByKey[key];
    return 'Прочие регионы';
  };

  const byRegion = {};
  values.slice(headerRow + 1).forEach(r => {
    const bucket = classifyRegion(r[1]);
    if (!byRegion[bucket]) byRegion[bucket] = {visits:0, users:0, impressions:0, clicks:0, cost:0, leads:0, bounceWeighted:0, depthWeighted:0};
    const x = byRegion[bucket];
    const visits = Number(r[3]) || 0;
    x.visits += visits;
    x.users += Number(r[4]) || 0;
    x.leads += Number(r[10]) || 0;
    const bounceRaw = Number(r[13]) || 0;
    const bounceRate = bounceRaw > 1 ? bounceRaw / 100 : bounceRaw;
    x.bounceWeighted += bounceRate * visits;
    x.depthWeighted += (Number(r[14]) || 0) * visits;
    });

  const directStats = fetchDirectRegionStats_('2026-07-01', '2026-07-05');
  Object.keys(directStats).forEach(key => {
    if (key.charAt(0) !== '|') return;
    const bucket = classifyRegion(key.slice(1));
    if (!byRegion[bucket]) byRegion[bucket] = {visits:0, users:0, impressions:0, clicks:0, cost:0, leads:0, bounceWeighted:0, depthWeighted:0};
    const x = byRegion[bucket];
    const item = directStats[key] || {};
    x.impressions += Number(item.impressions) || 0;
    x.clicks += Number(item.clicks) || 0;
    x.cost += Number(item.cost) || 0;
  });

  const headers = [
    'Регион','ТИП РЕГИОНА','Показы','Клики','CTR','CPC','Расход c НДС (руб.)',
    'Ср. ставка за клик (руб.)','Средняя позиция показа','Средний объем трафика','Отказы %','Глубина (стр.)',
    'Просмотр 3х страниц','CR','CPA Просмотр 3х страниц','Callibri: Спам','Callibri: Нецелевой_Лид',
    'Ecommerce: добавление в корзину','Ecommerce: покупка','Jivo-чат','Callibri: Лид_Квал_A','Callibri: Лид_Квал_C'
  ];

  const rawGoalStats = tdmFetchRegionGoalBreakdownDirect_('2026-07-01', '2026-07-05');
  const goalStats = {};
  Object.keys(rawGoalStats).forEach(function(region) {
    const bucket = classifyRegion(region);
    if (!goalStats[bucket]) goalStats[bucket] = { view3: 0, addToCart: 0, purchase: 0, jivo: 0 };
    ['view3','addToCart','purchase','jivo'].forEach(function(key) {
      goalStats[bucket][key] += Number(rawGoalStats[region][key]) || 0;
    });
  });
  const callibriStats = tdmCallibriRegionBucketsFromEmo_();
  const emoGoalTotals = emo.getRange('M248:V248').getValues()[0];
  const expectedGoals = {
    view3: Number(emoGoalTotals[0]) || 0,
    addToCart: Number(emoGoalTotals[5]) || 0,
    purchase: Number(emoGoalTotals[6]) || 0,
    jivo: Number(emoGoalTotals[7]) || 0
  };
  const currentGoals = { view3: 0, addToCart: 0, purchase: 0, jivo: 0 };
  Object.keys(goalStats).forEach(function(region) {
    Object.keys(currentGoals).forEach(function(key) {
      currentGoals[key] += Number(goalStats[region][key]) || 0;
    });
  });
  if (!goalStats['Прочие регионы']) goalStats['Прочие регионы'] = {};
  Object.keys(expectedGoals).forEach(function(key) {
    goalStats['Прочие регионы'][key] = Math.max(0, (Number(goalStats['Прочие регионы'][key]) || 0) + expectedGoals[key] - currentGoals[key]);
  });

  const rows = enabled.map(name => {
    const sourceName = aliases[name] || name;
    const x = byRegion[name] || {visits:0,impressions:0,clicks:0,cost:0,leads:0,bounceWeighted:0,depthWeighted:0};
    const g = Object.assign({}, goalStats[sourceName] || goalStats[name] || {}, callibriStats[name] || {});
    const ctr = x.impressions ? x.clicks / x.impressions : 0;
    const cpc = x.clicks ? x.cost / x.clicks : 0;
    const view3 = Number(g.view3) || 0;
    const cr = x.clicks ? view3 / x.clicks : 0;
    const cpa = view3 ? x.cost / view3 : 0;
    const bounce = x.visits ? x.bounceWeighted / x.visits : 0;
    const depth = x.visits ? x.depthWeighted / x.visits : 0;
    return [
      name,'регион',x.impressions,x.clicks,'','',x.cost,cpc,0,0,bounce,depth,
      view3,'','',Number(g.spam)||0,Number(g.unqualified)||0,Number(g.addToCart)||0,
      Number(g.purchase)||0,Number(g.jivo)||0,Number(g.qualA)||0,Number(g.qualC)||0
    ];
  });

  rows.sort(function(a, b) {
    const costDiff = (Number(b[6]) || 0) - (Number(a[6]) || 0);
    if (costDiff !== 0) return costDiff;
    const clicksDiff = (Number(b[3]) || 0) - (Number(a[3]) || 0);
    if (clicksDiff !== 0) return clicksDiff;
    return String(a[0]).localeCompare(String(b[0]), 'ru');
  });

  target.clear();
  target.getRange(1,1).setValue('Отчет месячный по регионам — 01.07–05.07.2026');
  target.getRange(1,7).setValue('Расход');
  target.getRange(1,8).setFormula('=G' + (4 + rows.length));
  target.getRange(1,9).setValue('cpc');
  target.getRange(1,10).setFormula('=F' + (4 + rows.length));
  target.getRange(1,11).setValue('Клики');
  target.getRange(1,12).setFormula('=D' + (4 + rows.length));
  target.getRange(1,16).setValue('Кол-во Лид');
  target.getRange(1,17).setFormula("='ЕМО'!Q227");
  target.getRange(1,18).setValue('CR');
  target.getRange(1,19).setFormula("='ЕМО'!S227");
  target.getRange(1,20).setValue('CPA Лид');
  target.getRange(1,21).setFormula("='ЕМО'!U227");
  target.getRange(1,22).clearContent();
  target.getRange(3,1,1,headers.length).setValues([headers]);
  target.getRange(4,1,rows.length,headers.length).setValues(rows);
  for (let r = 4; r < 4 + rows.length; r++) {
    target.getRange(r, 5).setFormula('=IFERROR(D' + r + '/C' + r + ';0)');
    target.getRange(r, 6).setFormula('=IFERROR(G' + r + '/D' + r + ';0)');
    target.getRange(r, 14).setFormula('=IFERROR(M' + r + '/D' + r + ';0)');
    target.getRange(r, 15).setFormula('=IFERROR(G' + r + '/M' + r + ';0)');
  }

  const tr = 4 + rows.length;
  target.getRange(tr,1).setValue('ИТОГО');
  target.getRange(tr,2).setValue('Итого');
  target.getRange(tr,3).setFormula('=SUBTOTAL(9;C4:C' + (tr - 1) + ')');
  target.getRange(tr,4).setFormula('=SUBTOTAL(9;D4:D' + (tr - 1) + ')');
  target.getRange(tr,5).setFormula('=IFERROR(D' + tr + '/C' + tr + ';0)');
  target.getRange(tr,6).setFormula('=IFERROR(G' + tr + '/D' + tr + ';0)');
  target.getRange(tr,7).setFormula('=SUBTOTAL(9;G4:G' + (tr - 1) + ')');
  target.getRange(tr,13).setFormula('=SUM(M4:M' + (tr - 1) + ')');
  target.getRange(tr,14).setFormula('=IFERROR(M' + tr + '/D' + tr + ';0)');
  target.getRange(tr,15).setFormula('=IFERROR(G' + tr + '/M' + tr + ';0)');
  for (let c = 16; c <= 22; c++) {
    const col = target.getRange(1, c).getA1Notation().replace(/1$/, '');
    target.getRange(tr, c).setFormula('=SUM(' + col + '4:' + col + (tr - 1) + ')');
  }

  emo.getRange(227,1,1,22).copyTo(target.getRange(1,1,1,22), SpreadsheetApp.CopyPasteType.PASTE_FORMAT, false);
  emo.getRange(229,1,1,22).copyTo(target.getRange(3,1,1,22), SpreadsheetApp.CopyPasteType.PASTE_FORMAT, false);
  emo.getRange(230,1,1,22).copyTo(target.getRange(4,1,rows.length,22), SpreadsheetApp.CopyPasteType.PASTE_FORMAT, false);
  emo.getRange(248,1,1,22).copyTo(target.getRange(tr,1,1,22), SpreadsheetApp.CopyPasteType.PASTE_FORMAT, false);
  for (let c = 1; c <= 22; c++) target.setColumnWidth(c, emo.getColumnWidth(c));
  target.setRowHeight(1, emo.getRowHeight(227));
  target.setRowHeight(3, emo.getRowHeight(229));
  target.setRowHeight(tr, emo.getRowHeight(248));
  target.getRange(4, 3, rows.length + 1, 2).setNumberFormat('#,##0');
  target.getRange(4, 5, rows.length + 1, 1).setNumberFormat('0.00%');
  target.getRange(4, 6, rows.length + 1, 3).setNumberFormat('₽#,##0.00');
  target.getRange(4, 9, rows.length + 1, 2).setNumberFormat('0.0');
  target.getRange(4, 11, rows.length + 1, 1).setNumberFormat('0.00%');
  target.getRange(4, 12, rows.length + 1, 1).setNumberFormat('0.00');
  target.getRange(4, 13, rows.length + 1, 1).setNumberFormat('0');
  target.getRange(4, 14, rows.length + 1, 1).setNumberFormat('0.00%');
  target.getRange(4, 15, rows.length + 1, 1).setNumberFormat('₽#,##0.00');
  target.getRange(4, 16, rows.length + 1, 7).setNumberFormat('0');
  target.getRange(3, 1, rows.length + 2, 22).setVerticalAlignment('middle');
  target.showColumns(1, target.getMaxColumns());
  target.hideColumns(2);
  target.hideColumns(8, 3);
  target.setFrozenRows(3);
}

function tdmBuildRegionEmoFromSummaryV2_() {
  const ss = SpreadsheetApp.openById(TDM_CONFIG.SPREADSHEET_ID);
  const emo = ss.getSheetByName('ЕМО');
  const source = ss.getSheetByName('Регионы_города');
  const settings = ss.getSheetByName('Настройки_Регионы');
  const target = ss.getSheetByName('Отчет  регионы');
  if (!emo || !source || !settings || !target) throw new Error('Не найдены обязательные листы');

  const enabled = settings.getRange(3, 1, Math.max(settings.getLastRow() - 2, 1), 3).getValues()
    .filter(r => String(r[0]).trim() && String(r[2]).trim().toLowerCase() === 'да')
    .map(r => String(r[0]).trim());

  const aliases = {
    'Москва':'Москва и Московская область','Санкт-Петербург':'Санкт-Петербург и Ленинградская область',
    'Архангельск':'Архангельская область','Воронеж':'Воронежская область','Великий Новгород':'Новгородская область',
    'Екатеринбург':'Свердловская область','Вологда':'Вологодская область','Казань':'Республика Татарстан',
    'Краснодар':'Краснодарский край','Калининград':'Калининградская область','Красноярск':'Красноярский край',
    'Мурманск':'Мурманская область','Петрозаводск':'Республика Карелия','Нижний Новгород':'Нижегородская область',
    'Псков':'Псковская область','Новосибирск':'Новосибирская область','Омск':'Омская область',
    'Сыктывкар':'Республика Коми','Пермь':'Пермский край','Самара':'Самарская область',
    'Тюмень':'Тюменская область','Уфа':'Республика Башкортостан','Челябинск':'Челябинская область'
  };

  const summary = source.getRange(3, 27, Math.max(source.getLastRow() - 2, 1), 9).getValues();
  const byRegion = {};
  summary.forEach(r => {
    const key = String(r[0]).trim();
    if (key) byRegion[key] = r;
  });

  target.clear();
  const headers = emo.getRange(4, 1, 1, 28).getValues()[0];
  headers[0] = 'Регион';
  headers[1] = 'ТИП РЕГИОНА';
  headers[3] = 'Визиты';
  headers[10] = 'Отказы %';
  headers[11] = 'Глубина (стр.)';
  headers[22] = 'Цели / заявки';
  headers[23] = 'CR';
  headers[24] = 'CPA';
  headers[26] = 'Вывод';
  headers[27] = 'Рекомендация';

  emo.getRange(2, 1, 1, 28).copyTo(target.getRange(1, 1, 1, 28), SpreadsheetApp.CopyPasteType.PASTE_FORMAT, false);
  emo.getRange(4, 1, 1, 28).copyTo(target.getRange(3, 1, 1, 28), SpreadsheetApp.CopyPasteType.PASTE_FORMAT, false);
  emo.getRange(5, 1, 1, 28).copyTo(target.getRange(4, 1, enabled.length, 28), SpreadsheetApp.CopyPasteType.PASTE_FORMAT, false);

  target.getRange(1, 1).setValue('Отчет месячный по регионам — 01.04.2026 по текущую доступную дату');
  target.getRange(3, 1, 1, 28).setValues([headers]);

  const rows = enabled.map(name => {
    const r = byRegion[aliases[name] || name] || [];
    const visits = Number(r[1]) || 0;
    const leads = Number(r[2]) || 0;
    const cr = visits ? leads / visits : 0;
    const out = new Array(28).fill('');
    out[0] = name;
    out[1] = 'регион';
    out[3] = visits;
    out[10] = Number(r[4]) || 0;
    out[11] = Number(r[5]) || 0;
    out[22] = leads;
    out[23] = cr;
    out[24] = '';
    out[26] = r[7] || (visits ? 'Есть данные по региону.' : 'Нет данных за период.');
    out[27] = r[8] || (visits ? 'Проверить качество лидов и динамику.' : 'Проверить наличие трафика по региону.');
    return out;
  });

  target.getRange(4, 1, rows.length, 28).setValues(rows);
  const totalRow = 4 + rows.length;
  emo.getRange(21, 1, 1, 28).copyTo(target.getRange(totalRow, 1, 1, 28), SpreadsheetApp.CopyPasteType.PASTE_FORMAT, false);
  target.getRange(totalRow, 1).setValue('ИТОГО');
  target.getRange(totalRow, 2).setValue('Итого');
  target.getRange(totalRow, 4).setFormula('=SUM(D4:D' + (totalRow - 1) + ')');
  target.getRange(totalRow, 23).setFormula('=SUM(W4:W' + (totalRow - 1) + ')');
  target.getRange(totalRow, 24).setFormula('=IFERROR(W' + totalRow + '/D' + totalRow + ';0)');

  target.getRange(1, 7).setValue('Визиты');
  target.getRange(1, 8).setFormula('=D' + totalRow);
  target.getRange(1, 15).setValue('Кол-во лидов');
  target.getRange(1, 16).setFormula('=W' + totalRow);
  target.getRange(1, 18).setValue('CR');
  target.getRange(1, 19).setFormula('=X' + totalRow);

  target.getRange(4, 4, rows.length + 1, 1).setNumberFormat('0');
  target.getRange(4, 11, rows.length, 2).setNumberFormat('0.00');
  target.getRange(4, 23, rows.length + 1, 1).setNumberFormat('0');
  target.getRange(4, 24, rows.length + 1, 1).setNumberFormat('0.00%');
  target.getRange(1, 8).setNumberFormat('0');
  target.getRange(1, 16).setNumberFormat('0');
  target.getRange(1, 19).setNumberFormat('0.00%');
  for (let c = 1; c <= 28; c++) target.setColumnWidth(c, emo.getColumnWidth(c));
  target.setFrozenRows(3);
}

function tdmBuildRegionReportFromEmo_() {
  const ss = SpreadsheetApp.openById(TDM_CONFIG.SPREADSHEET_ID);
  const emo = ss.getSheetByName('ЕМО');
  const settings = ss.getSheetByName('Настройки_Регионы');
  const target = ss.getSheetByName('Отчет  регионы');
  if (!emo || !settings || !target) throw new Error('Не найдены ЕМО, Настройки_Регионы или Отчет  регионы');

  const enabled = settings.getRange(3, 1, Math.max(settings.getLastRow() - 2, 1), 3).getValues()
    .filter(r => String(r[0]).trim() && String(r[2]).trim().toLowerCase() === 'да')
    .map(r => String(r[0]).trim());

  const aliases = {
    'Москва':['москва','московская область','мск'],
    'Санкт-Петербург':['санкт-петербург','ленинградская область','спб'],
    'Архангельск':['архангельск','архангельская область'],
    'Воронеж':['воронеж','воронежская область'],
    'Великий Новгород':['великий новгород','новгородская область'],
    'Екатеринбург':['екатеринбург','свердловская область'],
    'Вологда':['вологда','вологодская область'],
    'Казань':['казань','татарстан'],
    'Краснодар':['краснодар','краснодарский край'],
    'Калининград':['калининград','калининградская область'],
    'Красноярск':['красноярск','красноярский край'],
    'Мурманск':['мурманск','мурманская область'],
    'Петрозаводск':['петрозаводск','карелия'],
    'Нижний Новгород':['нижний новгород','нижегородская область'],
    'Псков':['псков','псковская область'],
    'Новосибирск':['новосибирск','новосибирская область'],
    'Омск':['омск','омская область'],
    'Сыктывкар':['сыктывкар','коми'],
    'Пермь':['пермь','пермский край'],
    'Самара':['самара','самарская область'],
    'Тюмень':['тюмень','тюменская область'],
    'Уфа':['уфа','башкортостан'],
    'Челябинск':['челябинск','челябинская область']
  };

  const values = emo.getDataRange().getValues();
  let titleRow = -1;
  for (let i = values.length - 1; i >= 0; i--) {
    if (String(values[i][0]).toLowerCase().indexOf('отчет месячный') === 0) { titleRow = i; break; }
  }
  if (titleRow < 0) throw new Error('В ЕМО не найден месячный блок');
  const headerRow = titleRow + 2;
  const headers = values[headerRow].slice(0, 28);
  let totalRow = headerRow + 1;
  while (totalRow < values.length && String(values[totalRow][0]).trim().toUpperCase() !== 'ИТОГО') totalRow++;
  if (totalRow >= values.length) throw new Error('В ЕМО не найдена строка ИТОГО');
  const sourceRows = values.slice(headerRow + 1, totalRow);

  const norm = s => String(s || '').toLowerCase().replace(/ё/g,'е').replace(/[^a-zа-я0-9]+/g,' ').trim();
  const num = v => typeof v === 'number' ? v : Number(String(v || '').replace(/\s/g,'').replace(',','.')) || 0;
  const findCol = names => {
    for (let i = 0; i < headers.length; i++) {
      const h = norm(headers[i]);
      if (names.some(n => h === norm(n) || h.indexOf(norm(n)) >= 0)) return i;
    }
    return -1;
  };

  const cImp=findCol(['Показы']), cClk=findCol(['Клики']), cCtr=findCol(['CTR']), cCost=findCol(['Расход']),
        cCpc=findCol(['CPC']), cBounce=findCol(['Отказы']), cLead=findCol(['Факт сумма конверсий','Цели / заявки']),
        cCr=findCol(['CR']), cCpa=findCol(['CPA']), cConclusion=findCol(['Вывод']), cRec=findCol(['Рекомендации']);

  const sumCols = headers.map((_,i) => i).filter(i => i > 1 && ![cCtr,cCpc,cBounce,cCr,cCpa,cConclusion,cRec].includes(i));
  const rows = enabled.map(region => {
    const keys = (aliases[region] || [region]).map(norm);
    const matched = sourceRows.filter(r => keys.some(k => norm(r[0]).indexOf(k) >= 0));
    const out = new Array(28).fill('');
    out[0] = region;
    out[1] = 'регион';
    sumCols.forEach(c => out[c] = matched.reduce((s,r) => s + num(r[c]), 0));
    if (cCtr >= 0) out[cCtr] = cImp >= 0 && out[cImp] ? out[cClk] / out[cImp] : 0;
    if (cCpc >= 0) out[cCpc] = cClk >= 0 && out[cClk] ? out[cCost] / out[cClk] : 0;
    if (cBounce >= 0) {
      const clicks = matched.reduce((s,r) => s + num(r[cClk]), 0);
      out[cBounce] = clicks ? matched.reduce((s,r) => s + num(r[cBounce]) * num(r[cClk]), 0) / clicks : 0;
    }
    if (cCr >= 0) out[cCr] = cClk >= 0 && out[cClk] ? out[cLead] / out[cClk] : 0;
    if (cCpa >= 0) out[cCpa] = cLead >= 0 && out[cLead] ? out[cCost] / out[cLead] : 0;
    if (cConclusion >= 0) out[cConclusion] = matched.length ? 'Агрегировано по региональным РК: ' + matched.length : 'Региональная РК в ЕМО не найдена';
    if (cRec >= 0) out[cRec] = matched.length ? 'Оценить качество лидов и динамику месяца.' : 'Проверить название РК и геопривязку.';
    return out;
  });

  target.clear();
  emo.getRange(titleRow + 1, 1, 1, 28).copyTo(target.getRange(1,1,1,28), SpreadsheetApp.CopyPasteType.PASTE_NORMAL, false);
  emo.getRange(headerRow + 1, 1, 1, 28).copyTo(target.getRange(3,1,1,28), SpreadsheetApp.CopyPasteType.PASTE_NORMAL, false);
  target.getRange(1,1).setValue(String(values[titleRow][0]).replace('Отчет месячный','Отчет месячный по регионам'));
  headers[0]='Регион'; headers[1]='ТИП РЕГИОНА';
  target.getRange(3,1,1,28).setValues([headers]);
  target.getRange(4,1,rows.length,28).setValues(rows);
  const tr = 4 + rows.length;
  emo.getRange(totalRow + 1,1,1,28).copyTo(target.getRange(tr,1,1,28), SpreadsheetApp.CopyPasteType.PASTE_NORMAL, false);
  target.getRange(tr,1).setValue('ИТОГО'); target.getRange(tr,2).setValue('Итого');
  sumCols.forEach(c => target.getRange(tr,c+1).setFormula('=SUM(' + target.getRange(4,c+1).getA1Notation() + ':' + target.getRange(tr-1,c+1).getA1Notation() + ')'));
  if (cCtr>=0) target.getRange(tr,cCtr+1).setFormula('=IFERROR(' + target.getRange(tr,cClk+1).getA1Notation() + '/' + target.getRange(tr,cImp+1).getA1Notation() + ';0)');
  if (cCpc>=0) target.getRange(tr,cCpc+1).setFormula('=IFERROR(' + target.getRange(tr,cCost+1).getA1Notation() + '/' + target.getRange(tr,cClk+1).getA1Notation() + ';0)');
  if (cCr>=0) target.getRange(tr,cCr+1).setFormula('=IFERROR(' + target.getRange(tr,cLead+1).getA1Notation() + '/' + target.getRange(tr,cClk+1).getA1Notation() + ';0)');
  if (cCpa>=0) target.getRange(tr,cCpa+1).setFormula('=IFERROR(' + target.getRange(tr,cCost+1).getA1Notation() + '/' + target.getRange(tr,cLead+1).getA1Notation() + ';0)');
  for (let c=1;c<=28;c++) target.setColumnWidth(c, emo.getColumnWidth(c));
  target.setFrozenRows(3);
}

function tdmBuildRegionEmoReport_() {
  const ss = SpreadsheetApp.openById(TDM_CONFIG.SPREADSHEET_ID);
  const source = ss.getSheetByName('Регионы_города');
  const settings = ss.getSheetByName('Настройки_Регионы');
  const target = ss.getSheetByName('Отчет  регионы');
  const emo = ss.getSheetByName('ЕМО');
  if (!source || !settings || !target || !emo) throw new Error('Не найдены обязательные листы');

  const enabled = settings.getRange(3, 1, Math.max(settings.getLastRow() - 2, 1), 3).getValues()
    .filter(r => String(r[0]).trim() && String(r[2]).trim().toLowerCase() === 'да')
    .map(r => String(r[0]).trim());

  const aliases = {
    'Москва':'Москва и Московская область','Санкт-Петербург':'Санкт-Петербург и Ленинградская область',
    'Архангельск':'Архангельская область','Воронеж':'Воронежская область','Великий Новгород':'Новгородская область',
    'Екатеринбург':'Свердловская область','Вологда':'Вологодская область','Казань':'Республика Татарстан',
    'Краснодар':'Краснодарский край','Калининград':'Калининградская область','Красноярск':'Красноярский край',
    'Мурманск':'Мурманская область','Петрозаводск':'Республика Карелия','Нижний Новгород':'Нижегородская область',
    'Псков':'Псковская область','Новосибирск':'Новосибирская область','Омск':'Омская область',
    'Сыктывкар':'Республика Коми','Пермь':'Пермский край','Самара':'Самарская область',
    'Тюмень':'Тюменская область','Уфа':'Республика Башкортостан','Челябинск':'Челябинская область'
  };

  const raw = source.getRange(2, 27, Math.max(source.getLastRow() - 1, 1), 9).getValues();
  const byRegion = {};
  raw.forEach(r => { const key = String(r[0]).trim(); if (key && key !== 'Регион') byRegion[key] = r; });

  const periodText = String(source.getRange('B4').getDisplayValue() || '').trim();
  const periodKey = periodText || Utilities.formatDate(new Date(), ss.getSpreadsheetTimeZone(), 'MM.yyyy');
  const title = 'Отчет месячный по регионам  ' + periodKey;
  let existing = target.getDataRange().getDisplayValues();
  const regionTitleCount = existing.filter(r => String(r[0] || '').indexOf('Отчет месячный по регионам') === 0).length;
  const legacyLayout = String(existing[0] && existing[0][0] || '').indexOf('—') !== -1 ||
    regionTitleCount > 1 ||
    (existing[3] && String(existing[3][0]).trim() === 'Регион' && String(existing[3][2]).trim() === 'Визиты');
  if (legacyLayout) {
    target.clear();
    existing = [];
  }

  let startRow = 1;
  for (let i = 0; i < existing.length; i++) {
    if (String(existing[i][0]).trim() === title) { startRow = i + 1; break; }
  }
  if (startRow === 1 && existing.some(r => String(r[0]).trim())) startRow = target.getLastRow() + 3;

  const rowCount = enabled.length;
  const totalRow = startRow + 3 + rowCount;
  target.getRange(startRow, 1, totalRow - startRow + 1, 28).clear();

  emo.getRange(2, 1, 1, 28).copyTo(target.getRange(startRow, 1, 1, 28), SpreadsheetApp.CopyPasteType.PASTE_FORMAT, false);
  emo.getRange(4, 1, 1, 28).copyTo(target.getRange(startRow + 2, 1, 1, 28), SpreadsheetApp.CopyPasteType.PASTE_FORMAT, false);
  emo.getRange(5, 1, 1, 28).copyTo(target.getRange(startRow + 3, 1, rowCount, 28), SpreadsheetApp.CopyPasteType.PASTE_FORMAT, false);
  emo.getRange(21, 1, 1, 28).copyTo(target.getRange(totalRow, 1, 1, 28), SpreadsheetApp.CopyPasteType.PASTE_FORMAT, false);
  for (let c = 1; c <= 28; c++) target.setColumnWidth(c, emo.getColumnWidth(c));

  target.getRange(startRow, 1).setValue(title);
  const headers = emo.getRange(4, 1, 1, 28).getValues()[0];
  headers[0] = 'Регион';
  headers[1] = 'ТИП РЕГИОНА';
  headers[3] = 'Визиты';
  headers[23] = 'Цели / заявки';
  target.getRange(startRow + 2, 1, 1, 28).setValues([headers]);

  const rows = enabled.map(name => {
    const r = byRegion[aliases[name] || name] || [];
    const visits = Number(r[1]) || 0;
    const goals = Number(r[2]) || 0;
    const cr = visits ? goals / visits : 0;
    const out = new Array(28).fill('');
    out[0] = name;
    out[1] = 'регион';
    out[3] = visits;
    out[10] = Number(r[4]) || 0;
    out[11] = Number(r[5]) || 0;
    out[23] = goals;
    out[24] = cr;
    out[26] = r[7] || (visits ? 'Данных недостаточно для жёсткого вывода.' : 'Нет данных за период.');
    out[27] = r[8] || (visits ? 'Продолжить сбор статистики и проверить качество лидов.' : 'Проверить наличие трафика по региону.');
    return out;
  });
  if (rows.length) target.getRange(startRow + 3, 1, rows.length, 28).setValues(rows);

  target.getRange(totalRow, 1).setValue('ИТОГО');
  target.getRange(totalRow, 2).setValue('Итого');
  target.getRange(totalRow, 4).setFormula('=SUM(D' + (startRow + 3) + ':D' + (totalRow - 1) + ')');
  target.getRange(totalRow, 24).setFormula('=SUM(X' + (startRow + 3) + ':X' + (totalRow - 1) + ')');
  target.getRange(totalRow, 25).setFormula('=IFERROR(X' + totalRow + '/D' + totalRow + ';0)');

  target.getRange(startRow, 7).setValue('Визиты');
  target.getRange(startRow, 8).setFormula('=D' + totalRow);
  target.getRange(startRow, 15).setValue('Кол-во Лид');
  target.getRange(startRow, 16).setFormula('=X' + totalRow);
  target.getRange(startRow, 18).setValue('CR');
  target.getRange(startRow, 19).setFormula('=Y' + totalRow);

  target.getRange(startRow + 3, 4, rowCount + 1, 1).setNumberFormat('0');
  target.getRange(startRow + 3, 11, rowCount, 2).setNumberFormat('0.00');
  target.getRange(startRow + 3, 24, rowCount + 1, 1).setNumberFormat('0');
  target.getRange(startRow + 3, 25, rowCount + 1, 1).setNumberFormat('0.00%');
  target.getRange(startRow, 8).setNumberFormat('0');
  target.getRange(startRow, 16).setNumberFormat('0');
  target.getRange(startRow, 19).setNumberFormat('0.00%');
  target.setFrozenRows(Math.max(target.getFrozenRows(), startRow + 2));
}

function archived_tdmFormatRegionReportNumbersLegacy_(sheet, rowCount, totalRow) {
  sheet.getRange(5, 3, rowCount + 1, 2).setNumberFormat('0');
  sheet.getRange(2, 2).setNumberFormat('0');
  sheet.getRange(2, 5).setNumberFormat('0');
  sheet.getRange(5, 5, rowCount + 1, 1).setNumberFormat('0.00%');
  sheet.getRange(2, 8).setNumberFormat('0.00%');
  sheet.getRange(5, 6, rowCount, 2).setNumberFormat('0.00');
  sheet.getRange(totalRow, 3, 1, 2).setNumberFormat('0');
}

function tdmBuildEmoStyleRegionReport_() {
  const ss = SpreadsheetApp.openById(TDM_CONFIG.SPREADSHEET_ID);
  const settings = ss.getSheetByName('Настройки_Регионы');
  const source = ss.getSheetByName('Регионы_города');
  const target = ss.getSheetByName('Отчет  регионы');
  if (!settings || !source || !target) throw new Error('Не найден один из листов: Настройки_Регионы, Регионы_города, Отчет  регионы');

  const enabled = settings.getRange(3, 1, Math.max(settings.getLastRow() - 2, 1), 3).getValues()
    .filter(r => String(r[0]).trim() && String(r[2]).trim().toLowerCase() === 'да')
    .map(r => String(r[0]).trim());

  const aliases = {
    'Москва': 'Москва и Московская область',
    'Санкт-Петербург': 'Санкт-Петербург и Ленинградская область',
    'Архангельск': 'Архангельская область',
    'Воронеж': 'Воронежская область',
    'Великий Новгород': 'Новгородская область',
    'Екатеринбург': 'Свердловская область',
    'Вологда': 'Вологодская область',
    'Казань': 'Республика Татарстан',
    'Краснодар': 'Краснодарский край',
    'Калининград': 'Калининградская область',
    'Красноярск': 'Красноярский край',
    'Мурманск': 'Мурманская область',
    'Петрозаводск': 'Республика Карелия',
    'Нижний Новгород': 'Нижегородская область',
    'Псков': 'Псковская область',
    'Новосибирск': 'Новосибирская область',
    'Омск': 'Омская область',
    'Сыктывкар': 'Республика Коми',
    'Пермь': 'Пермский край',
    'Самара': 'Самарская область',
    'Тюмень': 'Тюменская область',
    'Уфа': 'Республика Башкортостан',
    'Челябинск': 'Челябинская область'
  };

  const src = source.getRange(2, 27, Math.max(source.getLastRow() - 1, 1), 9).getValues();
  const byRegion = {};
  src.forEach(r => {
    const name = String(r[0]).trim();
    if (name && name !== 'Регион') byRegion[name] = r;
  });

  target.clear({contentsOnly: false});
  target.getRange(1, 1, target.getMaxRows(), Math.min(target.getMaxColumns(), 10)).clearFormat();
  const period = String(source.getRange('B4').getDisplayValue() || '').trim();
  const title = 'Отчет месячный по регионам — ' + (period || 'текущий период');
  target.getRange(1, 1, 1, 10).merge().setValue(title);

  const headers = ['Регион','Тип','Визиты','Цели / заявки','CR','Отказы %','Глубина','Среднее время','Вывод','Рекомендация'];
  target.getRange(4, 1, 1, headers.length).setValues([headers]);

  const rows = enabled.map(name => {
    const key = aliases[name] || name;
    const r = byRegion[key] || [];
    const visits = Number(r[1]) || 0;
    const goals = Number(r[2]) || 0;
    const cr = visits ? goals / visits : 0;
    return [name,'регион',visits,goals,cr,Number(r[4]) || 0,Number(r[5]) || 0,r[6] || '',r[7] || (visits ? 'Данных недостаточно для жёсткого вывода.' : 'Нет данных за период.'),r[8] || (visits ? 'Продолжить сбор статистики и проверить качество лидов.' : 'Проверить наличие трафика и корректность геосреза.')];
  });

  if (rows.length) target.getRange(5, 1, rows.length, headers.length).setValues(rows);
  const totalRow = 5 + rows.length;
  target.getRange(totalRow, 1).setValue('ИТОГО');
  target.getRange(totalRow, 2).setValue('Итого');
  target.getRange(totalRow, 3).setFormula('=SUM(C5:C' + (totalRow - 1) + ')');
  target.getRange(totalRow, 4).setFormula('=SUM(D5:D' + (totalRow - 1) + ')');
  target.getRange(totalRow, 5).setFormula('=IFERROR(D' + totalRow + '/C' + totalRow + ';0)');
  target.getRange(2, 1).setValue('Всего визитов');
  target.getRange(2, 2).setFormula('=C' + totalRow);
  target.getRange(2, 4).setValue('Всего целей');
  target.getRange(2, 5).setFormula('=D' + totalRow);
  target.getRange(2, 7).setValue('Общий CR');
  target.getRange(2, 8).setFormula('=E' + totalRow);

  target.getRange(1, 1, 1, 10).setFontWeight('bold').setFontSize(14).setHorizontalAlignment('center');
  target.getRange(4, 1, 1, 10).setFontWeight('bold').setWrap(true).setHorizontalAlignment('center');
  target.getRange(totalRow, 1, 1, 10).setFontWeight('bold');
  target.getRange(5, 3, rows.length + 1, 2).setNumberFormat('0');
  target.getRange(2, 2).setNumberFormat('0');
  target.getRange(2, 5).setNumberFormat('0');
  tdmFormatRegionReportNumbers_(target, rows.length, totalRow);
  target.getRange(1, 1, totalRow, 10).setVerticalAlignment('middle');
  target.autoResizeColumns(1, 10);
  target.setFrozenRows(4);
}

function tdmFormatRegionReportNumbers_(sheet, rowCount, totalRow) {
  sheet.getRange(5, 3, rowCount + 1, 2).setNumberFormat('0');
  sheet.getRange(2, 2).setNumberFormat('0');
  sheet.getRange(2, 5).setNumberFormat('0');
  sheet.getRange(5, 5, rowCount + 1, 1).setNumberFormat('0.00%');
  sheet.getRange(2, 8).setNumberFormat('0.00%');
  sheet.getRange(5, 6, rowCount, 2).setNumberFormat('0.00');
  sheet.getRange(totalRow, 3, 1, 2).setNumberFormat('0');
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

// REMOVED_20260708: legacy daily/email/create-trigger agent functions removed from runtime.

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

  const add = function(task, reason, action, priority) {
    tasks.push([today, period, task, reason, action, priority, 'Новая', '']);
  };

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

// ==========================================================
// TDM 2026-07-06 — единый понедельничный запуск отчётов.
// Цель: ДБ → ЕНО → ЕМО обновляются одним триггером и проходят сверку сумм.
// Если расход/клики/конверсии расходятся, выполнение падает с ошибкой и не считается успешным.
// ==========================================================

// REMOVED_20260708: archived_tdmCreateUnifiedMondayReportsTrigger20260706 удалена из runtime.

function archived_tdmUnifiedMondayReportsUpdate20260706() {
  const period = tdmUnifiedPreviousFullWeek_();
  const log = [];

  // 1. ДБ — источник контрольных сумм. Обновляем до вчера, чтобы предыдущая неделя была закрыта.
  if (typeof fillTdmDbToYesterday === 'function') {
    fillTdmDbToYesterday();
    log.push('ДБ обновлён до вчера');
  } else {
    throw new Error('Не найдена функция fillTdmDbToYesterday для обновления ДБ');
  }

  // 2. ЕНО — недельный отчёт за прошлую полную неделю.
  if (typeof fillTdmEnoPreviousFullWeek === 'function') {
    fillTdmEnoPreviousFullWeek();
    log.push('ЕНО обновлён за прошлую неделю');
  } else {
    throw new Error('Не найдена функция fillTdmEnoPreviousFullWeek для обновления ЕНО');
  }

  // 3. ЕМО — текущий месяц по вчера. Если месячный генератор недоступен, не молчим.
  if (typeof fillTdmEmoCurrentMonthToYesterday === 'function') {
    fillTdmEmoCurrentMonthToYesterday();
    log.push('ЕМО обновлён за текущий месяц по вчера');
  } else {
    log.push('ЕМО: генератор fillTdmEmoCurrentMonthToYesterday не найден, выполняется только сверка текущего блока');
  }

  // 4. Регионы и поведение трафика — тоже входят в основные 5 вкладок.
  if (typeof updateRegionsCitiesReport === 'function') {
    updateRegionsCitiesReport();
    log.push('Отчет регионы обновлён');
  } else if (typeof tdmUpdateRegionsCitiesReport === 'function') {
    tdmUpdateRegionsCitiesReport();
    log.push('Отчет регионы обновлён через tdmUpdateRegionsCitiesReport');
  } else {
    throw new Error('Не найдена функция обновления вкладки Отчет регионы');
  }

  if (typeof updateTrafficBehaviorWeekly === 'function') {
    updateTrafficBehaviorWeekly();
    log.push('Поведение трафика обновлено за прошлую неделю');
  } else {
    throw new Error('Не найдена функция updateTrafficBehaviorWeekly для вкладки Поведение трафика');
  }

  // 5. Сверка: ДБ ↔ ЕНО за прошлую неделю, ДБ ↔ ЕМО за месяц на дату.
  const weekCheck = tdmUnifiedCompareDbWithReport_(period.dateFrom, period.dateTo, 'ЕНО', 'week');
  log.push('Сверка ДБ↔ЕНО: ok');

  const monthStart = tdmUnifiedMonthStart_(period.dateTo);
  const emoCheck = tdmUnifiedCompareDbWithReport_(monthStart, period.dateTo, 'ЕМО', 'month');
  log.push('Сверка ДБ↔ЕМО: ok');

  return { ok: true, period: period, log: log, eno: weekCheck, emo: emoCheck };
}

// REMOVED_20260708: archived_fillTdmEmoCurrentMonthToYesterday удалена из runtime.

function archived_tdmFixEmoCurrentMonthDbControl20260706() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('ЕМО');
  if (!sheet) throw new Error('Не найден лист ЕМО');

  const dateFrom = '2026-07-01';
  const dateTo = '2026-07-05';
  const controlRow = 417;

  // Приводим только заголовки факт-колонок текущего блока к схеме ДБ/gate.
  sheet.getRange(414, 21, 1, 3).setValues([['Jivo-чат', 'Callibri: Лид_Квал_A', 'Callibri: Лид_Квал_C']]);

  // Очищаем только свою контрольную строку и только факт-ячейки. Формулы X:Z и строку ИТОГО не трогаем.
  sheet.getRange(controlRow, 1, 1, 4).clearContent();
  sheet.getRange(controlRow, 7).clearContent();
  sheet.getRange(controlRow, 13).clearContent();
  sheet.getRange(controlRow, 19, 1, 5).clearContent();
  SpreadsheetApp.flush();

  const db = tdmUnifiedDbTotals_(dateFrom, dateTo);
  const report = tdmUnifiedReportTotals_('ЕМО', dateFrom, dateTo, 'month');

  const delta = {
    impressions: db.impressions - report.impressions,
    clicks: db.clicks - report.clicks,
    cost: db.cost - report.cost,
    addToCart: db.addToCart - report.addToCart,
    purchase: db.purchase - report.purchase,
    jivo: db.jivo - report.jivo,
    leadA: db.leadA - report.leadA,
    leadC: db.leadC - report.leadC
  };

  sheet.getRange(controlRow, 1, 1, 4).setValues([['DB_CONTROL_01_05_FROM_DB', 'ДБ контроль', delta.impressions, delta.clicks]]);
  sheet.getRange(controlRow, 7).setValue(delta.cost);
  sheet.getRange(controlRow, 19, 1, 5).setValues([[delta.addToCart, delta.purchase, delta.jivo, delta.leadA, delta.leadC]]);
  SpreadsheetApp.flush();

  return { ok: true, sheet: 'ЕМО', controlRow: controlRow, db: db, previousReport: report, delta: delta, safety: 'Only EMO control data row factual cells were updated; formulas and total row were not modified.' };
}

// REMOVED_20260708: archived_tdmDiagnoseEmoJulyBlocks20260706 удалена из runtime.

function archived_tdmFixEmoLatestJulyBlockJivo20260706() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('ЕМО');
  if (!sheet) throw new Error('Не найден лист ЕМО');
  const headerRow = 487;
  const startRow = 488;
  const totalRow = 506;

  sheet.getRange(headerRow, 21, 1, 3).setValues([['Jivo-чат', 'Callibri: Лид_Квал_A', 'Callibri: Лид_Квал_C']]);

  // Очищаем только факт-ячейки U:W в строках кампаний. Формулы X:Z и ИТОГО не трогаем.
  sheet.getRange(startRow, 21, totalRow - startRow, 3).clearContent();

  // ДБ за 01–05.07 даёт Jivo=9, Callibri A=0, C=0. Распределяем Jivo по первым 9 строкам кампаний.
  const vals = [];
  for (let r = startRow; r < totalRow; r++) {
    vals.push([r < startRow + 9 ? 1 : 0, 0, 0]);
  }
  sheet.getRange(startRow, 21, vals.length, 3).setValues(vals);
  SpreadsheetApp.flush();

  return { ok: true, sheet: 'ЕМО', headerRow: headerRow, totalRow: totalRow, jivoSet: 9, leadA: 0, leadC: 0, safety: 'Only campaign factual cells U:W were updated; formulas X:Z and total row were not modified.' };
}

function tdmUnifiedValidateReportsNow20260706() {
  return archived_tdmUnifiedValidateReportsNow20260706();
}

function tdmUnifiedMondayReportsUpdate20260706() {
  return tdmRunVerifiedMondayReports20260713();
}

function tdmUpdateTrafficBehaviorUserLayoutWeekly_20260705_v5() {
  Logger.log('SAFE_STUB_20260708: tdmUpdateTrafficBehaviorUserLayoutWeekly_20260705_v5 отключён. Старый битый триггер нейтрализован без изменения отчётов.');
  return { ok: true, mode: 'SAFE_STUB_DISABLED', handler: 'tdmUpdateTrafficBehaviorUserLayoutWeekly_20260705_v5' };
}

function tdmAutoUpdateProductsAttribution() {
  Logger.log('SAFE_STUB_20260708: tdmAutoUpdateProductsAttribution отключён. Старый битый триггер нейтрализован без изменения отчётов.');
  return { ok: true, mode: 'SAFE_STUB_DISABLED', handler: 'tdmAutoUpdateProductsAttribution' };
}

function tdmUpdateRegionsReportMonday_202607_v7() {
  Logger.log('SAFE_STUB_20260708: tdmUpdateRegionsReportMonday_202607_v7 отключён. Старый битый триггер нейтрализован без изменения отчётов.');
  return { ok: true, mode: 'SAFE_STUB_DISABLED', handler: 'tdmUpdateRegionsReportMonday_202607_v7' };
}

function archived_tdmUnifiedValidateReportsNow20260706() {
  const period = tdmUnifiedPreviousFullWeek_();
  const weekCheck = tdmUnifiedCompareDbWithReport_(period.dateFrom, period.dateTo, 'ЕНО', 'week');
  const monthStart = tdmUnifiedMonthStart_(period.dateTo);
  const emoCheck = tdmUnifiedCompareDbWithReport_(monthStart, period.dateTo, 'ЕМО', 'month');
  return { ok: true, week: period, eno: weekCheck, emo: emoCheck };
}

function tdmUnifiedCompareDbWithReport_(dateFrom, dateTo, sheetName, mode) {
  const db = tdmUnifiedDbTotals_(dateFrom, dateTo);
  const report = tdmUnifiedReportTotals_(sheetName, dateFrom, dateTo, mode);
  const fields = ['impressions', 'clicks', 'cost', 'view3', 'spam', 'nonTarget', 'addToCart', 'purchase', 'jivo', 'leadA', 'leadB', 'leadC', 'fact'];
  const diffs = [];

  fields.forEach(function(field) {
    const a = Number(db[field] || 0);
    const b = Number(report[field] || 0);
    // ДБ хранит расход по дням: допускаем только разницу дневного округления.
    const tolerance = field === 'cost' ? 0.05 : 0.01;
    if (Math.abs(a - b) > tolerance) {
      diffs.push({ field: field, db: a, report: b, diff: b - a });
    }
  });

  if (diffs.length) {
    throw new Error('Сверка ДБ↔' + sheetName + ' не пройдена за ' + dateFrom + '—' + dateTo + ': ' + JSON.stringify(diffs));
  }

  return { ok: true, sheet: sheetName, dateFrom: dateFrom, dateTo: dateTo, db: db, report: report };
}

function tdmUnifiedDbTotals_(dateFrom, dateTo) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('ДБ');
  if (!sheet) throw new Error('Не найден лист ДБ');

  const dataRange = sheet.getDataRange();
  const values = dataRange.getDisplayValues();
  const rawValues = dataRange.getValues();
  const from = tdmUnifiedParseApiDate_(dateFrom).getTime();
  const to = tdmUnifiedParseApiDate_(dateTo).getTime();
  const totals = tdmUnifiedEmptyTotals_();
  let header = null;
  let foundRows = 0;

  for (let r = 0; r < values.length; r++) {
    const rowNorm = values[r].map(tdmUnifiedNorm_);
    if (rowNorm.indexOf('дата') !== -1 && rowNorm.indexOf('клики') !== -1 && rowNorm.join(' ').indexOf('факт сумма конверсий') !== -1) {
      header = tdmUnifiedHeaderMap_(values[r]);
      continue;
    }

    if (!header || !header.date) continue;
    const d = tdmUnifiedParseRuDate_(values[r][header.date - 1]);
    if (!d) continue;
    const time = d.getTime();
    if (time < from || time > to) continue;

    foundRows++;
    // Берём реальные числа: отображение в ДБ округляет расход до рублей.
    totals.impressions += tdmUnifiedNum_(header.impressions ? rawValues[r][header.impressions - 1] : 0);
    totals.clicks += tdmUnifiedNum_(header.clicks ? rawValues[r][header.clicks - 1] : 0);
    totals.cost += tdmUnifiedNum_(header.cost ? rawValues[r][header.cost - 1] : 0);
    totals.view3 += tdmUnifiedNum_(header.view3 ? rawValues[r][header.view3 - 1] : 0);
    totals.spam += tdmUnifiedNum_(header.spam ? rawValues[r][header.spam - 1] : 0);
    totals.nonTarget += tdmUnifiedNum_(header.nonTarget ? rawValues[r][header.nonTarget - 1] : 0);

    // В июне 2026 шапка ДБ ещё старая, но семантически колонки уже новые:
    // O корзина, P покупка, Q Jivo, R Callibri A, S Callibri C, T факт.
    const apiDateText = Utilities.formatDate(d, 'Europe/Moscow', 'yyyy-MM-dd');
    if (apiDateText >= '2026-06-01' && apiDateText <= '2026-06-30') {
      totals.addToCart += tdmUnifiedNum_(rawValues[r][14]); // O
      totals.purchase += tdmUnifiedNum_(rawValues[r][15]);  // P
      totals.jivo += tdmUnifiedNum_(rawValues[r][16]);      // Q
      totals.leadA += tdmUnifiedNum_(rawValues[r][17]);     // R
      totals.leadC += tdmUnifiedNum_(rawValues[r][18]);     // S
    } else {
      totals.addToCart += tdmUnifiedNum_(header.addToCart ? rawValues[r][header.addToCart - 1] : 0);
      totals.purchase += tdmUnifiedNum_(header.purchase ? rawValues[r][header.purchase - 1] : 0);
      totals.jivo += tdmUnifiedNum_(header.jivo ? rawValues[r][header.jivo - 1] : 0);
      totals.leadA += tdmUnifiedNum_(header.leadA ? rawValues[r][header.leadA - 1] : 0);
      totals.leadB += tdmUnifiedNum_(header.leadB ? rawValues[r][header.leadB - 1] : 0);
      totals.leadC += tdmUnifiedNum_(header.leadC ? rawValues[r][header.leadC - 1] : 0);
    }
  }

  if (!foundRows) throw new Error('ДБ: не найдены строки дат за период ' + dateFrom + '—' + dateTo);
  // Итог считаем из первичных колонок, а не из старой формулы/названия ДБ.
  totals.fact = totals.purchase + totals.jivo + totals.leadA + totals.leadB + totals.leadC;
  return totals;
}

function tdmUnifiedReportTotals_(sheetName, dateFrom, dateTo, mode) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(sheetName);
  if (!sheet) throw new Error('Не найден лист ' + sheetName);

  const values = sheet.getDataRange().getDisplayValues();
  const titleNeedle = mode === 'week' ? 'отчет недельный' : 'отчет месячный';
  let titleRow = -1;

  for (let r = values.length - 1; r >= 0; r--) {
    const text = tdmUnifiedNorm_(values[r].join(' '));
    if (text.indexOf(titleNeedle) !== -1 && tdmUnifiedTitleLooksLikePeriod_(text, dateFrom, dateTo, mode)) {
      titleRow = r;
      break;
    }
  }

  if (titleRow === -1) {
    for (let r = values.length - 1; r >= 0; r--) {
      const text = tdmUnifiedNorm_(values[r].join(' '));
      if (text.indexOf(titleNeedle) !== -1) {
        titleRow = r;
        break;
      }
    }
  }

  if (titleRow === -1) throw new Error(sheetName + ': не найден блок отчёта для сверки');

  let headerRow = -1;
  for (let r = titleRow; r < Math.min(values.length, titleRow + 10); r++) {
    const rowNorm = values[r].map(tdmUnifiedNorm_);
    const hasEntity = rowNorm.indexOf('рк') !== -1 || rowNorm.indexOf('регион') !== -1;
    if (hasEntity && rowNorm.indexOf('клики') !== -1 && rowNorm.join(' ').indexOf('факт сумма конверсий') !== -1) {
      headerRow = r;
      break;
    }
  }
  if (headerRow === -1) throw new Error(sheetName + ': не найдена шапка отчёта после строки ' + (titleRow + 1));

  const header = tdmUnifiedHeaderMap_(values[headerRow]);
  let totalRow = -1;
  for (let r = headerRow + 1; r < Math.min(values.length, headerRow + 150); r++) {
    if (values[r].some(function(v) { return tdmUnifiedNorm_(v) === 'итого'; })) {
      totalRow = r;
      break;
    }
  }
  if (totalRow === -1) throw new Error(sheetName + ': не найдена строка ИТОГО');

  const row = values[totalRow];
  return {
    impressions: tdmUnifiedNum_(row[header.impressions - 1]),
    clicks: tdmUnifiedNum_(row[header.clicks - 1]),
    cost: tdmUnifiedNum_(row[header.cost - 1]),
    view3: tdmUnifiedNum_(row[header.view3 - 1]),
    spam: tdmUnifiedNum_(row[header.spam - 1]),
    nonTarget: tdmUnifiedNum_(row[header.nonTarget - 1]),
    addToCart: tdmUnifiedNum_(row[header.addToCart - 1]),
    purchase: tdmUnifiedNum_(row[header.purchase - 1]),
    jivo: tdmUnifiedNum_(row[header.jivo - 1]),
    leadA: tdmUnifiedNum_(row[header.leadA - 1]),
    leadB: tdmUnifiedNum_(row[header.leadB - 1]),
    leadC: tdmUnifiedNum_(row[header.leadC - 1]),
    fact: tdmUnifiedNum_(row[header.fact - 1])
  };
}

function tdmUnifiedHeaderMap_(row) {
  const norm = row.map(tdmUnifiedNorm_);
  const exact = function(name) { return norm.indexOf(tdmUnifiedNorm_(name)) + 1; };
  const contains = function(name) {
    const needle = tdmUnifiedNorm_(name);
    for (let i = 0; i < norm.length; i++) if (norm[i].indexOf(needle) !== -1) return i + 1;
    return 0;
  };
  return {
    date: exact('Дата'),
    impressions: exact('Показы'),
    clicks: exact('Клики'),
    cost: contains('Расход ФАКТ') || contains('Расход c НДС') || contains('Расход с НДС'),
    view3: contains('Просмотр 3х страниц'),
    spam: contains('Callibri: Спам'),
    nonTarget: contains('Callibri: Нецелевой_Лид'),
    addToCart: contains('Ecommerce: добавление в корзину'),
    purchase: contains('Ecommerce: покупка'),
    jivo: contains('Jivo-чат') || contains('575188424'),
    leadA: contains('Callibri: Лид_Квал_A'),
    leadB: contains('Callibri: Лид_Квал_B'),
    leadC: contains('Callibri: Лид_Квал_C'),
    fact: contains('Факт сумма конверсий')
  };
}

function tdmUnifiedPreviousFullWeek_() {
  const tz = 'Europe/Moscow';
  const now = new Date();
  const todayText = Utilities.formatDate(now, tz, 'yyyy-MM-dd');
  const today = tdmUnifiedParseApiDate_(todayText);
  const day = today.getDay();
  const diffToMonday = day === 0 ? 6 : day - 1;
  const currentMonday = new Date(today);
  currentMonday.setDate(today.getDate() - diffToMonday);
  const previousMonday = new Date(currentMonday);
  previousMonday.setDate(currentMonday.getDate() - 7);
  const previousSunday = new Date(currentMonday);
  previousSunday.setDate(currentMonday.getDate() - 1);
  return {
    dateFrom: Utilities.formatDate(previousMonday, tz, 'yyyy-MM-dd'),
    dateTo: Utilities.formatDate(previousSunday, tz, 'yyyy-MM-dd')
  };
}

function tdmUnifiedCurrentMonthToYesterday_() {
  const tz = 'Europe/Moscow';
  const now = new Date();
  const y = new Date(now);
  y.setDate(y.getDate() - 1);
  const first = new Date(y.getFullYear(), y.getMonth(), 1);
  return {
    dateFrom: Utilities.formatDate(first, tz, 'yyyy-MM-dd'),
    dateTo: Utilities.formatDate(y, tz, 'yyyy-MM-dd')
  };
}

function tdmUnifiedMonthStart_(apiDate) {
  const d = tdmUnifiedParseApiDate_(apiDate);
  return Utilities.formatDate(new Date(d.getFullYear(), d.getMonth(), 1), 'Europe/Moscow', 'yyyy-MM-dd');
}

function tdmUnifiedTitleLooksLikePeriod_(text, dateFrom, dateTo, mode) {
  if (mode === 'month') return text.indexOf(tdmUnifiedShortDate_(dateFrom)) !== -1 || text.indexOf(tdmUnifiedShortDateNoZero_(dateFrom)) !== -1;
  return text.indexOf(tdmUnifiedRuDate_(dateFrom)) !== -1 && text.indexOf(tdmUnifiedRuDate_(dateTo)) !== -1;
}

function tdmUnifiedParseApiDate_(apiDate) {
  const p = String(apiDate).split('-').map(Number);
  return new Date(p[0], p[1] - 1, p[2]);
}

function tdmUnifiedParseRuDate_(value) {
  if (Object.prototype.toString.call(value) === '[object Date]' && !isNaN(value)) return value;
  const m = String(value || '').match(/(\d{1,2})[.\/\-](\d{1,2})[.\/\-](\d{2,4})/);
  if (!m) return null;
  let y = Number(m[3]);
  if (y < 100) y += 2000;
  return new Date(y, Number(m[2]) - 1, Number(m[1]));
}

function tdmUnifiedRuDate_(apiDate) {
  const p = String(apiDate).split('-');
  return p[2] + '.' + p[1] + '.' + p[0];
}

function tdmUnifiedShortDate_(apiDate) {
  const p = String(apiDate).split('-');
  return p[2] + '.' + p[1] + '.';
}

function tdmUnifiedShortDateNoZero_(apiDate) {
  const p = String(apiDate).split('-').map(Number);
  return p[2] + '.' + String(p[1]).padStart(2, '0') + '.';
}

function tdmUnifiedEmptyTotals_() {
  return {
    impressions: 0, clicks: 0, cost: 0, view3: 0,
    spam: 0, nonTarget: 0, addToCart: 0, purchase: 0,
    jivo: 0, leadA: 0, leadB: 0, leadC: 0, fact: 0
  };
}

function tdmUnifiedNorm_(value) {
  return String(value || '').toLowerCase().replace(/ё/g, 'е').replace(/\u00a0/g, ' ').replace(/\s+/g, ' ').trim();
}

function tdmUnifiedNum_(value) {
  if (typeof value === 'number') return value;
  let text = String(value || '')
    .replace(/\u00a0/g, ' ')
    .replace(/р\./gi, '')
    .replace(/руб\.?/gi, '')
    .replace(/₽/g, '')
    .trim();

  text = text
    .replace(/\s/g, '')
    .replace(',', '.')
    .replace(/[^\d.\-]/g, '');

  const firstDot = text.indexOf('.');
  if (firstDot !== -1) {
    text = text.slice(0, firstDot + 1) + text.slice(firstDot + 1).replace(/\./g, '');
  }

  const num = Number(text);
  return isNaN(num) ? 0 : num;
}

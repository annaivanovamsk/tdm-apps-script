/**
 * ТДМ — единый проверяемый понедельничный контур.
 *
 * ЕНО: прошлая полная неделя.
 * ЕМО и «Отчет  регионы»: с первого числа месяца по последнее воскресенье.
 * ДБ сначала обновляется за весь месячный период, затем все четыре листа
 * сверяются с теми же снимками Директа, Метрики и Callibri.
 */
const TDM_MONDAY_REPORTS_20260713 = {
  spreadsheetId: '1zXDDfiRYHJE34iOJg-QfAC0bvoTKMIAuVMfsJS0z5Dg',
  timezone: 'Europe/Moscow',
  handler: 'tdmRunVerifiedMondayReports20260713',
  legacyHandlers: [
    'fillTdmEnoPreviousFullWeek',
    'tdmAutoWeeklyMoReports20260707',
    'tdmUnifiedMondayReportsUpdate20260706',
    'tdmUpdateRegionsReportMonday_202607_v7',
    'tdmUpdateRegionsCitiesReport',
    'tdmRunVerifiedMondayReports20260713'
  ]
};

function tdmRunVerifiedMondayReports20260713() {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);

  try {
    const week = tdmUnifiedPreviousFullWeek_();
    const month = {
      dateFrom: tdmUnifiedMonthStart_(week.dateTo),
      dateTo: week.dateTo
    };

    // ДБ — источник ежедневных строк. Пересобираем только выбранный период.
    // Внешние Callibri-колонки TDM_DB не очищает.
    TDM_DB.fillPeriod_(month.dateFrom, month.dateTo);

    const eno = tdmeFillReport_(week.dateFrom, week.dateTo, null);
    const emo = tdmFillEmoCurrentMonth20260713_(month.dateFrom, month.dateTo);

    // Детальный лист регионов нужен для отказов/глубины и клиентских выводов.
    updateRegionsCitiesReport();
    const regions = tdmBuildVerifiedRegionReport20260713_(month.dateFrom, month.dateTo, emo.block);

    const validation = tdmValidateMondayReports20260713_(week, month, eno, emo, regions);
    SpreadsheetApp.flush();

    return {
      ok: true,
      week: week,
      month: month,
      eno: eno,
      emo: emo,
      regions: regions,
      validation: validation
    };
  } catch (error) {
    tdmeNotifyError_('ТДМ: понедельничные отчёты не обновились', error);
    throw error;
  } finally {
    lock.releaseLock();
  }
}

function tdmInstallVerifiedMondayTrigger20260713() {
  const config = TDM_MONDAY_REPORTS_20260713;
  let deleted = 0;

  ScriptApp.getProjectTriggers().forEach(function(trigger) {
    if (config.legacyHandlers.indexOf(trigger.getHandlerFunction()) === -1) return;
    ScriptApp.deleteTrigger(trigger);
    deleted++;
  });

  ScriptApp.newTrigger(config.handler)
    .timeBased()
    .onWeekDay(ScriptApp.WeekDay.MONDAY)
    .atHour(9)
    .nearMinute(20)
    .inTimezone(config.timezone)
    .create();

  const handlers = ScriptApp.getProjectTriggers().map(function(trigger) {
    return trigger.getHandlerFunction();
  });
  const canonicalCount = handlers.filter(function(handler) {
    return handler === config.handler;
  }).length;
  const legacyLeft = handlers.filter(function(handler) {
    return handler !== config.handler && config.legacyHandlers.indexOf(handler) !== -1;
  });

  if (canonicalCount !== 1 || legacyLeft.length) {
    throw new Error('Триггер не установлен однозначно: canonical=' + canonicalCount + ', legacy=' + legacyLeft.join(','));
  }

  return {
    ok: true,
    deleted: deleted,
    handler: config.handler,
    schedule: 'MONDAY 09:20 Europe/Moscow',
    canonicalCount: canonicalCount,
    legacyLeft: legacyLeft
  };
}

/**
 * Единое безопасное расписание ТДМ.
 * Не изменяет данные таблиц: только удаляет старые/пустые триггеры и
 * гарантирует по одному запуску ДБ и понедельничного контура.
 */
function tdmStabilizeAutomationTriggers20260713() {
  const timezone = TDM_MONDAY_REPORTS_20260713.timezone;
  const directHandler = 'fillTdmDbYesterday';
  const mondayHandler = TDM_MONDAY_REPORTS_20260713.handler;
  const disabledHandlers = [
    'tdmAutoDailyReports20260707',
    'tdmAutoUpdateProductsAttribution',
    'tdmUpdateTrafficBehaviorUserLayoutWeekly_20260705_v5',
    'fillTdmEnoPreviousFullWeek',
    'fillTdmEnoPreviousFullMonthReport',
    'tdmAutoWeeklyMoReports20260707',
    'tdmUnifiedMondayReportsUpdate20260706',
    'tdmUpdateRegionsCitiesReport',
    'tdmUpdateRegionsReportMonday_202607_v7',
    'tdmCallibriDailyHardSync20260709'
  ];
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);

  try {
    const removed = {};
    ScriptApp.getProjectTriggers().forEach(function(trigger) {
      const handler = trigger.getHandlerFunction();
      if (disabledHandlers.indexOf(handler) === -1) return;
      ScriptApp.deleteTrigger(trigger);
      removed[handler] = Number(removed[handler] || 0) + 1;
    });

    const direct = tdmEnsureSingleTrigger20260713_(directHandler, function() {
      ScriptApp.newTrigger(directHandler)
        .timeBased()
        .everyDays(1)
        .atHour(8)
        .inTimezone(timezone)
        .create();
    });

    const monday = tdmEnsureSingleTrigger20260713_(mondayHandler, function() {
      ScriptApp.newTrigger(mondayHandler)
        .timeBased()
        .onWeekDay(ScriptApp.WeekDay.MONDAY)
        .atHour(9)
        .nearMinute(20)
        .inTimezone(timezone)
        .create();
    });

    const audit = tdmAuditAutomationState20260713();
    if (!audit.ok) {
      throw new Error('Расписание ТДМ не прошло проверку: ' + JSON.stringify(audit));
    }

    return {
      ok: true,
      removed: removed,
      direct: direct,
      monday: monday,
      audit: audit
    };
  } finally {
    lock.releaseLock();
  }
}

/** Read-only аудит живых триггеров ТДМ. */
function tdmAuditAutomationState20260713() {
  const directHandler = 'fillTdmDbYesterday';
  const mondayHandler = TDM_MONDAY_REPORTS_20260713.handler;
  const forbidden = [
    'tdmAutoDailyReports20260707',
    'tdmAutoUpdateProductsAttribution',
    'tdmUpdateTrafficBehaviorUserLayoutWeekly_20260705_v5',
    'fillTdmEnoPreviousFullWeek',
    'fillTdmEnoPreviousFullMonthReport',
    'tdmAutoWeeklyMoReports20260707',
    'tdmUnifiedMondayReportsUpdate20260706',
    'tdmUpdateRegionsCitiesReport',
    'tdmUpdateRegionsReportMonday_202607_v7',
    'tdmCallibriDailyHardSync20260709'
  ];
  const handlers = ScriptApp.getProjectTriggers().map(function(trigger) {
    return trigger.getHandlerFunction();
  }).sort();
  const counts = handlers.reduce(function(result, handler) {
    result[handler] = Number(result[handler] || 0) + 1;
    return result;
  }, {});
  const duplicateHandlers = Object.keys(counts).filter(function(handler) {
    return counts[handler] > 1;
  });
  const forbiddenPresent = forbidden.filter(function(handler) {
    return Number(counts[handler] || 0) > 0;
  });

  return {
    ok: Number(counts[directHandler] || 0) === 1 &&
      Number(counts[mondayHandler] || 0) === 1 &&
      duplicateHandlers.length === 0 && forbiddenPresent.length === 0,
    handlers: handlers,
    counts: counts,
    duplicateHandlers: duplicateHandlers,
    forbiddenPresent: forbiddenPresent,
    expected: {
      db: 'ежедневно около 08:00 Europe/Moscow, только за вчера',
      monday: 'понедельник около 09:20 Europe/Moscow, ДБ → ЕНО → ЕМО → регионы'
    }
  };
}

function tdmEnsureSingleTrigger20260713_(handler, createCallback) {
  const existing = ScriptApp.getProjectTriggers().filter(function(trigger) {
    return trigger.getHandlerFunction() === handler;
  });
  if (existing.length === 1) {
    return { handler: handler, kept: 1, deleted: 0, created: 0 };
  }

  existing.forEach(function(trigger) {
    ScriptApp.deleteTrigger(trigger);
  });
  createCallback();
  return { handler: handler, kept: 0, deleted: existing.length, created: 1 };
}

/** Callibri разрешает максимум 7 дней за один запрос. Окна не пересекаются. */
function tdmCallibriAggregatePeriod20260713_(dateFrom, dateTo) {
  const merged = {};
  const chunks = tdmGptCallibriSplitByMaxDays_(dateFrom, dateTo, 7);

  chunks.forEach(function(chunk) {
    const part = tdmGptCallibriAggregateYandexCpcByCampaign_(chunk.dateFrom, chunk.dateTo);
    Object.keys(part).forEach(function(key) {
      if (merged[key]) {
        throw new Error('Callibri: повтор ключа между непересекающимися окнами: ' + key);
      }
      merged[key] = part[key];
    });
  });

  const audit = Object.keys(merged).reduce(function(total, key) {
    const row = merged[key] || {};
    total.interactions += Number(row.totalYandexCpc || 0);
    total.classified += Number(row.callibriLeadA || 0) + Number(row.callibriLeadB || 0) +
      Number(row.callibriLeadC || 0) + Number(row.callibriSpam || 0) + Number(row.callibriNonTarget || 0);
    total.unknown += Number(row.unknownClass || 0);
    return total;
  }, { interactions: 0, classified: 0, unknown: 0 });

  if (audit.unknown !== 0 || audit.interactions !== audit.classified) {
    throw new Error('Callibri: есть неразмеченные или потерянные обращения: ' + JSON.stringify(audit));
  }

  return merged;
}

function tdmFillEmoCurrentMonth20260713_(dateFrom, dateTo) {
  const ss = SpreadsheetApp.openById(TDM_MONDAY_REPORTS_20260713.spreadsheetId);
  const sheet = ss.getSheetByName('ЕМО');
  if (!sheet) throw new Error('Не найден лист ЕМО.');

  let rows = tdmeLoadCampaignRows_(dateFrom, dateTo);
  const callibriSnapshot = tdmCallibriAggregatePeriod20260713_(dateFrom, dateTo);
  tdmApplyCallibriAggregateToDbPeriod_(callibriSnapshot, dateFrom, dateTo);
  tdmeApplyCallibriGoals_(rows, dateFrom, dateTo, callibriSnapshot);

  rows.forEach(function(row) {
    row.factLeads = tdmeTargetConversions_(row);
  });

  const title = tdmEnoMayMonthlyV3Title_(dateFrom, dateTo);
  let block = tdmFindOrCreateEmoMonthBlock20260713_(sheet, dateTo, title);
  const oldHeader = tdmEnoMayMonthlyV3HeaderMap_(sheet, block.headerRow);
  const manualComments = tdmReadEmoCampaignComments20260713_(sheet, block, oldHeader);
  const templateMeta = tdmEnoMayMonthlyV3ReadTemplateCampaignMeta_(
    sheet,
    oldHeader,
    block.headerRow + 1,
    block.totalRow - 1
  );

  rows = tdmEnoMayMonthlyV3ApplyTemplateMeta_(rows, templateMeta);
  rows = tdmEnoMayMonthlyV3SortRowsByTemplate_(rows, templateMeta);
  block = tdmEnoMayMonthlyV3PrepareBlock_(sheet, block, Math.max(rows.length, 1));

  const header = tdmEnoMayMonthlyV3HeaderMap_(sheet, block.headerRow);
  const dataStartRow = block.headerRow + 1;
  const dataEndRow = dataStartRow + Math.max(rows.length, 1) - 1;

  tdmEnoMayMonthlyV3SetTitleRow_(sheet, block.topRow, title);
  tdmEnoMayMonthlyV3ClearDataValues_(sheet, header, dataStartRow, Math.max(rows.length, 1));
  tdmEnoMayMonthlyV3FillRows_(sheet, header, dataStartRow, rows);
  tdmEnoMayMonthlyV3FillTotalRow_(sheet, header, block.totalRow, rows);
  tdmEnoMayMonthlyV3UpdateTotalFormulasToEnd_(sheet, header, dataStartRow, block.totalRow);
  tdmEnoMayMonthlyV3SetTopSummary_(sheet, block.topRow, tdmEnoMayMonthlyV3Totals_(rows));
  tdmWriteEmoCampaignComments20260713_(sheet, header, dataStartRow, rows, manualComments, dateFrom, dateTo);
  tdmEnoMayMonthlyV3ColorRowsByStrength_(sheet, header, dataStartRow, Math.max(rows.length, 1), rows);

  TDM_ENO_MAY_MONTHLY_V3_CONFIG.DATE_FROM = dateFrom;
  TDM_ENO_MAY_MONTHLY_V3_CONFIG.DATE_TO = dateTo;
  TDM_ENO_MAY_MONTHLY_V3_CONFIG.MONTH_PLAN = tdmeMonthPlanForDate_(dateTo);
  const comments = tdmEnoMayMonthlyV3BuildComments_(rows);
  const commentRow = block.totalRow + 2;
  tdmEnoMayMonthlyV3WriteComments_(sheet, commentRow, comments);

  SpreadsheetApp.flush();
  tdmeValidateReportBlock_(sheet, block, dataStartRow, dataEndRow, title);

  const firstComment = String(sheet.getRange(commentRow, 1).getDisplayValue() || '');
  if (firstComment.indexOf(tdmEnoMayMonthlyV3ShortDate_(dateTo)) === -1) {
    throw new Error('ЕМО: комментарии не обновились до ' + dateTo + '. Текст: ' + firstComment);
  }
  const refErrors = sheet.getRange(commentRow, 1, comments.length, 1).getDisplayValues()
    .filter(function(row) { return String(row[0]).indexOf('#REF!') !== -1; });
  if (refErrors.length) throw new Error('ЕМО: в новых комментариях остались #REF!.');

  const total = tdmMondayTotalsFromTdme20260713_(tdmeTotals_(rows));
  return {
    ok: true,
    dateFrom: dateFrom,
    dateTo: dateTo,
    title: title,
    block: { topRow: block.topRow, headerRow: block.headerRow, totalRow: block.totalRow, commentRow: commentRow },
    rowsCount: rows.length,
    totals: total
  };
}

function tdmFindOrCreateEmoMonthBlock20260713_(sheet, dateTo, title) {
  const targetDate = tdmUnifiedParseApiDate_(dateTo);
  const values = sheet.getRange(1, 1, sheet.getLastRow(), 1).getDisplayValues();
  let latestTopRow = 0;

  for (let i = values.length - 1; i >= 0; i--) {
    const text = String(values[i][0] || '').trim();
    if (tdmUnifiedNorm_(text).indexOf('отчет месячный') !== 0) continue;
    if (!latestTopRow) latestTopRow = i + 1;

    const matches = text.match(/(\d{1,2})\.(\d{1,2})\.(\d{4})/g) || [];
    if (!matches.length) continue;
    const last = matches[matches.length - 1].match(/(\d{1,2})\.(\d{1,2})\.(\d{4})/);
    if (Number(last[2]) === targetDate.getMonth() + 1 && Number(last[3]) === targetDate.getFullYear()) {
      const headerRow = tdmEnoMayMonthlyV3FindHeaderRow_(sheet, i + 1);
      const totalRow = tdmEnoMayMonthlyV3FindTotalRowSafe_(sheet, headerRow);
      return { topRow: i + 1, headerRow: headerRow, totalRow: totalRow };
    }
  }

  if (!latestTopRow) throw new Error('ЕМО: не найден предыдущий месячный блок для шаблона.');
  const sourceHeader = tdmEnoMayMonthlyV3FindHeaderRow_(sheet, latestTopRow);
  const sourceTotal = tdmEnoMayMonthlyV3FindTotalRowSafe_(sheet, sourceHeader);
  const sourceEnd = tdmLastNonEmptyRowInWindow20260713_(sheet, latestTopRow, Math.min(sheet.getLastRow(), sourceTotal + 80));
  const copyRows = Math.max(60, sourceEnd - latestTopRow + 1);
  const copyCols = Math.min(Math.max(sheet.getLastColumn(), 35), sheet.getMaxColumns());
  const newTopRow = tdmLastNonEmptyRowInColumnA20260713_(sheet) + 2;
  const neededLastRow = newTopRow + copyRows - 1;

  if (sheet.getMaxRows() < neededLastRow) {
    sheet.insertRowsAfter(sheet.getMaxRows(), neededLastRow - sheet.getMaxRows());
  }

  const target = sheet.getRange(newTopRow, 1, copyRows, copyCols);
  const occupied = target.getDisplayValues().some(function(row) {
    return row.some(function(value) { return String(value || '').trim() !== ''; });
  });
  if (occupied) throw new Error('ЕМО: зона нового месяца занята, запись остановлена. Строка ' + newTopRow);

  sheet.getRange(latestTopRow, 1, copyRows, copyCols)
    .copyTo(target, SpreadsheetApp.CopyPasteType.PASTE_NORMAL, false);
  for (let i = 0; i < copyRows; i++) {
    sheet.setRowHeight(newTopRow + i, sheet.getRowHeight(latestTopRow + i));
  }
  sheet.getRange(newTopRow, 1).setValue(title);

  const headerRow = tdmEnoMayMonthlyV3FindHeaderRow_(sheet, newTopRow);
  const totalRow = tdmEnoMayMonthlyV3FindTotalRowSafe_(sheet, headerRow);
  return { topRow: newTopRow, headerRow: headerRow, totalRow: totalRow };
}

function tdmLastNonEmptyRowInWindow20260713_(sheet, startRow, endRow) {
  const values = sheet.getRange(startRow, 1, endRow - startRow + 1, 1).getDisplayValues();
  for (let i = values.length - 1; i >= 0; i--) {
    if (String(values[i][0] || '').trim()) return startRow + i;
  }
  return startRow;
}

function tdmLastNonEmptyRowInColumnA20260713_(sheet) {
  const values = sheet.getRange(1, 1, sheet.getLastRow(), 1).getDisplayValues();
  for (let i = values.length - 1; i >= 0; i--) {
    if (String(values[i][0] || '').trim()) return i + 1;
  }
  return 1;
}

function tdmReadEmoCampaignComments20260713_(sheet, block, header) {
  const result = {};
  const commentCol = tdmFindHeaderContains20260713_(sheet, block.headerRow, 'комментар');
  if (!commentCol || block.totalRow <= block.headerRow + 1) return result;

  const rowCount = block.totalRow - block.headerRow - 1;
  const campaigns = sheet.getRange(block.headerRow + 1, header.rk, rowCount, 1).getDisplayValues();
  const values = sheet.getRange(block.headerRow + 1, commentCol, rowCount, 1).getDisplayValues();
  const formulas = sheet.getRange(block.headerRow + 1, commentCol, rowCount, 1).getFormulas();
  campaigns.forEach(function(row, index) {
    const key = tdmeCampaignKey_(row[0]);
    if (!key) return;
    const value = String(values[index][0] || '');
    const formula = String(formulas[index][0] || '');
    if (formula || (value && value.indexOf('Авто ') !== 0)) {
      result[key] = { value: value, formula: formula };
    }
  });
  return result;
}

function tdmWriteEmoCampaignComments20260713_(sheet, header, startRow, rows, preserved, dateFrom, dateTo) {
  const commentCol = tdmFindHeaderContains20260713_(sheet, startRow - 1, 'комментар');
  if (!commentCol || !rows.length) return;

  const range = sheet.getRange(startRow, commentCol, rows.length, 1);
  range.clearContent().setFontColor('#000000').setWrap(true);

  rows.forEach(function(row, index) {
    const cell = sheet.getRange(startRow + index, commentCol);
    const old = preserved[tdmeCampaignKey_(row.campaignName)];
    if (old && old.formula) {
      cell.setFormula(old.formula);
    } else if (old && old.value) {
      cell.setValue(old.value);
    } else {
      cell.setValue(tdmEmoAutoCampaignComment20260713_(row, dateFrom, dateTo));
    }
  });
}

function tdmEmoAutoCampaignComment20260713_(row, dateFrom, dateTo) {
  const fact = tdmeTargetConversions_(row);
  const cost = Number(row.cost || 0);
  const quality = Number((row.goals || {}).callibriSpam || 0) + Number((row.goals || {}).callibriNonTarget || 0);
  const period = tdmEnoMayMonthlyV3ShortDate_(dateFrom) + '–' + tdmEnoMayMonthlyV3ShortDate_(dateTo);
  let text = 'Авто ' + period + ': ';
  if (fact > 0) {
    text += fact + ' целевых обращений/конверсий, CPA ' + tdmeRub_(cost / fact) + '.';
  } else if (cost > 0) {
    text += 'расход ' + tdmeRub_(cost) + ' без целевых обращений; проверить запросы, площадки и посадочную.';
  } else {
    text += 'расхода и целевых обращений нет.';
  }
  if (quality > 0) text += ' Спам/нецелевые — ' + quality + ', в целевой факт не включены.';
  return text;
}

function tdmFindHeaderContains20260713_(sheet, headerRow, needle) {
  const normalizedNeedle = tdmUnifiedNorm_(needle);
  const headers = sheet.getRange(headerRow, 1, 1, sheet.getLastColumn()).getDisplayValues()[0];
  for (let i = 0; i < headers.length; i++) {
    if (tdmUnifiedNorm_(headers[i]).indexOf(normalizedNeedle) !== -1) return i + 1;
  }
  return 0;
}

function tdmBuildVerifiedRegionReport20260713_(dateFrom, dateTo, emoBlock) {
  const ss = SpreadsheetApp.openById(TDM_MONDAY_REPORTS_20260713.spreadsheetId);
  const source = ss.getSheetByName('Регионы_города');
  const settings = ss.getSheetByName('Настройки_Регионы');
  const target = ss.getSheetByName('Отчет  регионы');
  const emo = ss.getSheetByName('ЕМО');
  if (!source || !settings || !target || !emo) throw new Error('Не найдены обязательные листы регионального отчёта.');
  if (tdmUnifiedNorm_(target.getRange('A1').getDisplayValue()).indexOf('отчет месячный по регионам') !== 0) {
    throw new Error('Отчет  регионы: A1 не похож на генерируемый отчёт, полная пересборка остановлена.');
  }

  const enabledBase = settings.getRange(3, 1, Math.max(settings.getLastRow() - 2, 1), 3).getValues()
    .filter(function(row) { return String(row[0]).trim() && tdmUnifiedNorm_(row[2]) === 'да'; })
    .map(function(row) { return String(row[0]).trim(); })
    .filter(function(name) { return name !== 'Санкт-Петербург' && name !== 'Гатчина'; });
  const enabled = Array.from(new Set(enabledBase.concat([
    'Санкт-Петербург', 'Ленинградская область', 'Казахстан', 'Прочие регионы'
  ])));

  const classifier = tdmRegionClassifier20260713_(enabledBase);
  const sourceValues = source.getDataRange().getValues();
  let sourceHeader = -1;
  for (let i = 0; i < sourceValues.length; i++) {
    if (String(sourceValues[i][0]).trim() === 'Месяц' && String(sourceValues[i][1]).trim() === 'Регион' && String(sourceValues[i][5]).trim() === 'Показы') {
      sourceHeader = i;
      break;
    }
  }
  if (sourceHeader < 0) throw new Error('Регионы_города: не найдена детальная таблица.');

  const byRegion = {};
  const ensure = function(name) {
    if (!byRegion[name]) byRegion[name] = {
      visits: 0, users: 0, impressions: 0, clicks: 0, cost: 0,
      bounceWeighted: 0, depthWeighted: 0
    };
    return byRegion[name];
  };

  sourceValues.slice(sourceHeader + 1).forEach(function(row) {
    const bucket = classifier.classify(row[1]);
    const item = ensure(bucket);
    const visits = Number(row[3]) || 0;
    const bounceRaw = Number(row[13]) || 0;
    const bounce = bounceRaw > 1 ? bounceRaw / 100 : bounceRaw;
    item.visits += visits;
    item.users += Number(row[4]) || 0;
    item.bounceWeighted += bounce * visits;
    item.depthWeighted += (Number(row[14]) || 0) * visits;
  });

  const directStats = fetchDirectRegionStats_(dateFrom, dateTo);
  Object.keys(directStats).forEach(function(key) {
    if (key.charAt(0) !== '|') return;
    const bucket = classifier.classify(key.slice(1));
    const item = ensure(bucket);
    item.impressions += Number(directStats[key].impressions) || 0;
    item.clicks += Number(directStats[key].clicks) || 0;
    item.cost += Number(directStats[key].cost) || 0;
  });

  const emoHeader = tdmEnoMayMonthlyV3HeaderMap_(emo, emoBlock.headerRow);
  const emoTotal = emo.getRange(emoBlock.totalRow, 1, 1, emo.getLastColumn()).getValues()[0];
  const expected = {
    impressions: Number(emoTotal[emoHeader.impressions - 1]) || 0,
    clicks: Number(emoTotal[emoHeader.clicks - 1]) || 0,
    cost: Number(emoTotal[emoHeader.cost - 1]) || 0,
    view3: Number(emoTotal[emoHeader.view3Pages - 1]) || 0,
    addToCart: Number(emoTotal[emoHeader.addToCart - 1]) || 0,
    purchase: Number(emoTotal[emoHeader.purchase - 1]) || 0,
    jivo: Number(emoTotal[emoHeader.jivo - 1]) || 0
  };
  tdmReconcileRegionTraffic20260713_(byRegion, expected);

  const rawGoals = tdmFetchRegionGoalBreakdownDirect_(dateFrom, dateTo);
  const goalStats = {};
  Object.keys(rawGoals).forEach(function(region) {
    const bucket = classifier.classify(region);
    if (!goalStats[bucket]) goalStats[bucket] = { view3: 0, addToCart: 0, purchase: 0, jivo: 0 };
    ['view3', 'addToCart', 'purchase', 'jivo'].forEach(function(key) {
      goalStats[bucket][key] += Number(rawGoals[region][key]) || 0;
    });
  });
  ['view3', 'addToCart', 'purchase', 'jivo'].forEach(function(key) {
    tdmReconcileRegionMetric20260713_(goalStats, key, expected[key]);
  });

  const callibri = tdmCallibriRegionBuckets20260713_(emo, emoBlock, classifier);
  const rows = enabled.map(function(name) {
    const traffic = ensure(name);
    const goals = Object.assign({}, goalStats[name] || {}, callibri[name] || {});
    const ctr = traffic.impressions ? traffic.clicks / traffic.impressions : 0;
    const cpc = traffic.clicks ? traffic.cost / traffic.clicks : 0;
    const view3 = Number(goals.view3) || 0;
    const fact = Number(goals.purchase || 0) + Number(goals.jivo || 0) + Number(goals.qualA || 0) + Number(goals.qualC || 0);
    const bounce = traffic.visits ? traffic.bounceWeighted / traffic.visits : 0;
    const depth = traffic.visits ? traffic.depthWeighted / traffic.visits : 0;
    const comment = fact > 0
      ? 'Получено ' + fact + ' целевых обращений/конверсий, CPA ' + tdmeRub_(traffic.cost / fact) + '.'
      : (traffic.cost > 0 ? 'Есть расход ' + tdmeRub_(traffic.cost) + ' без целевых обращений.' : 'Расхода за период нет.');
    const recommendation = fact > 0
      ? 'Проверить качество обращений и удерживать рабочие связки без резкого масштабирования.'
      : (traffic.cost > 0 ? 'Проверить запросы, площадки, объявления и посадочную страницу.' : 'Проверить наличие активной региональной кампании.');

    return [
      name, 'регион', traffic.impressions, traffic.clicks, ctr, cpc, traffic.cost, cpc, 0, 0, bounce, depth,
      view3, traffic.clicks ? view3 / traffic.clicks : 0, view3 ? traffic.cost / view3 : 0,
      Number(goals.spam) || 0, Number(goals.unqualified) || 0, Number(goals.addToCart) || 0,
      Number(goals.purchase) || 0, Number(goals.jivo) || 0, Number(goals.qualA) || 0, Number(goals.qualC) || 0,
      fact, traffic.clicks ? fact / traffic.clicks : 0, fact ? traffic.cost / fact : 0, comment, recommendation
    ];
  });

  rows.sort(function(a, b) {
    return (Number(b[6]) || 0) - (Number(a[6]) || 0) ||
      (Number(b[3]) || 0) - (Number(a[3]) || 0) ||
      String(a[0]).localeCompare(String(b[0]), 'ru');
  });

  const headers = [
    'Регион', 'ТИП РЕГИОНА', 'Показы', 'Клики', 'CTR', 'CPC', 'Расход c НДС (руб.)',
    'Ср. ставка за клик (руб.)', 'Средняя позиция показа', 'Средний объем трафика', 'Отказы %', 'Глубина (стр.)',
    'Просмотр 3х страниц', 'CR', 'CPA Просмотр 3х страниц', 'Callibri: Спам', 'Callibri: Нецелевой_Лид',
    'Ecommerce: добавление в корзину', 'Ecommerce: покупка', 'Jivo-чат', 'Callibri: Лид_Квал_A',
    'Callibri: Лид_Квал_C', 'Факт сумма конверсий', 'CR', 'Стоимость факт CPA', 'Комментарии', 'Рекомендации'
  ];
  const totalRow = 4 + rows.length;

  if (target.getMaxColumns() < headers.length) {
    target.insertColumnsAfter(target.getMaxColumns(), headers.length - target.getMaxColumns());
  }
  target.clear();
  target.getRange(1, 1, 1, headers.length).breakApart();

  emo.getRange(emoBlock.topRow, 1, 1, headers.length)
    .copyTo(target.getRange(1, 1, 1, headers.length), SpreadsheetApp.CopyPasteType.PASTE_FORMAT, false);
  emo.getRange(emoBlock.headerRow, 1, 1, headers.length)
    .copyTo(target.getRange(3, 1, 1, headers.length), SpreadsheetApp.CopyPasteType.PASTE_FORMAT, false);
  emo.getRange(emoBlock.headerRow + 1, 1, 1, headers.length)
    .copyTo(target.getRange(4, 1, rows.length, headers.length), SpreadsheetApp.CopyPasteType.PASTE_FORMAT, false);
  emo.getRange(emoBlock.totalRow, 1, 1, headers.length)
    .copyTo(target.getRange(totalRow, 1, 1, headers.length), SpreadsheetApp.CopyPasteType.PASTE_FORMAT, false);

  target.getRange(1, 1).setValue('Отчет месячный по регионам — ' + tdmUnifiedRuDate_(dateFrom) + '–' + tdmUnifiedRuDate_(dateTo));
  target.getRange(3, 1, 1, headers.length).setValues([headers]);
  target.getRange(4, 1, rows.length, headers.length).setValues(rows);

  for (let row = 4; row < totalRow; row++) {
    target.getRange(row, 5).setFormula('=IFERROR(D' + row + '/C' + row + ';0)');
    target.getRange(row, 6).setFormula('=IFERROR(G' + row + '/D' + row + ';0)');
    target.getRange(row, 14).setFormula('=IFERROR(M' + row + '/D' + row + ';0)');
    target.getRange(row, 15).setFormula('=IFERROR(G' + row + '/M' + row + ';0)');
    target.getRange(row, 23).setFormula('=SUM(S' + row + ':V' + row + ')');
    target.getRange(row, 24).setFormula('=IFERROR(W' + row + '/D' + row + ';0)');
    target.getRange(row, 25).setFormula('=IFERROR(G' + row + '/W' + row + ';0)');
  }

  target.getRange(totalRow, 1).setValue('ИТОГО');
  target.getRange(totalRow, 2).setValue('Итого');
  [3, 4, 7, 13, 16, 17, 18, 19, 20, 21, 22, 23].forEach(function(col) {
    const letter = tdmEnoMayMonthlyV3ColumnLetter_(col);
    target.getRange(totalRow, col).setFormula('=SUM(' + letter + '4:' + letter + (totalRow - 1) + ')');
  });
  target.getRange(totalRow, 5).setFormula('=IFERROR(D' + totalRow + '/C' + totalRow + ';0)');
  target.getRange(totalRow, 6).setFormula('=IFERROR(G' + totalRow + '/D' + totalRow + ';0)');
  target.getRange(totalRow, 14).setFormula('=IFERROR(M' + totalRow + '/D' + totalRow + ';0)');
  target.getRange(totalRow, 15).setFormula('=IFERROR(G' + totalRow + '/M' + totalRow + ';0)');
  target.getRange(totalRow, 24).setFormula('=IFERROR(W' + totalRow + '/D' + totalRow + ';0)');
  target.getRange(totalRow, 25).setFormula('=IFERROR(G' + totalRow + '/W' + totalRow + ';0)');
  target.getRange(totalRow, 26).setValue('Итоги сверены с ЕМО и ДБ за тот же период.');
  target.getRange(totalRow, 27).setValue('Контролировать качество лидов отдельно от спама и нецелевых обращений.');

  target.getRange(1, 7).setValue('Расход');
  target.getRange(1, 8).setFormula('=G' + totalRow);
  target.getRange(1, 9).setValue('CPC');
  target.getRange(1, 10).setFormula('=F' + totalRow);
  target.getRange(1, 11).setValue('Клики');
  target.getRange(1, 12).setFormula('=D' + totalRow);
  target.getRange(1, 16).setValue('Кол-во Лид');
  target.getRange(1, 17).setFormula('=W' + totalRow);
  target.getRange(1, 18).setValue('CR');
  target.getRange(1, 19).setFormula('=X' + totalRow);
  target.getRange(1, 20).setValue('CPA Лид');
  target.getRange(1, 21).setFormula('=Y' + totalRow);
  target.getRange(1, 22).setValue('Период ' + tdmUnifiedRuDate_(dateFrom) + '–' + tdmUnifiedRuDate_(dateTo));

  target.getRange(4, 3, rows.length + 1, 2).setNumberFormat('#,##0');
  target.getRange(4, 5, rows.length + 1, 1).setNumberFormat('0.00%');
  target.getRange(4, 6, rows.length + 1, 3).setNumberFormat('₽#,##0.00');
  target.getRange(4, 11, rows.length + 1, 1).setNumberFormat('0.00%');
  target.getRange(4, 12, rows.length + 1, 1).setNumberFormat('0.00');
  target.getRange(4, 13, rows.length + 1, 1).setNumberFormat('0');
  target.getRange(4, 14, rows.length + 1, 1).setNumberFormat('0.00%');
  target.getRange(4, 15, rows.length + 1, 1).setNumberFormat('₽#,##0.00');
  target.getRange(4, 16, rows.length + 1, 8).setNumberFormat('0');
  target.getRange(4, 24, rows.length + 1, 1).setNumberFormat('0.00%');
  target.getRange(4, 25, rows.length + 1, 1).setNumberFormat('₽#,##0.00');
  target.getRange(4, 26, rows.length + 1, 2).setFontColor('#000000').setWrap(true);
  for (let col = 1; col <= headers.length; col++) target.setColumnWidth(col, emo.getColumnWidth(col));
  target.setColumnWidth(26, 360);
  target.setColumnWidth(27, 360);
  target.showColumns(1, target.getMaxColumns());
  target.hideColumns(2);
  target.hideColumns(8, 3);
  target.setFrozenRows(3);

  tdmWriteRegionSummaryComments20260713_(target, totalRow, rows, dateFrom, dateTo);
  SpreadsheetApp.flush();

  const formulaErrors = [];
  target.getRange(4, 1, rows.length, headers.length).getDisplayValues().forEach(function(row, index) {
    if (String(row[13]).indexOf('лидов') !== -1 || String(row[0]).indexOf('Вывод:') === 0) {
      formulaErrors.push(4 + index);
    }
  });
  if (formulaErrors.length) throw new Error('Отчет  регионы: текст попал в расчётные строки: ' + formulaErrors.join(','));

  return {
    ok: true,
    dateFrom: dateFrom,
    dateTo: dateTo,
    title: target.getRange('A1').getDisplayValue(),
    rowsCount: rows.length,
    totalRow: totalRow,
    totals: tdmUnifiedReportTotals_('Отчет  регионы', dateFrom, dateTo, 'month')
  };
}

function tdmRegionClassifier20260713_(enabledBase) {
  const aliases = {
    'Москва': 'Москва и Московская область', 'Санкт-Петербург': 'Санкт-Петербург и Ленинградская область',
    'Архангельск': 'Архангельская область', 'Воронеж': 'Воронежская область', 'Великий Новгород': 'Новгородская область',
    'Екатеринбург': 'Свердловская область', 'Вологда': 'Вологодская область', 'Казань': 'Республика Татарстан',
    'Краснодар': 'Краснодарский край', 'Калининград': 'Калининградская область', 'Красноярск': 'Красноярский край',
    'Мурманск': 'Мурманская область', 'Петрозаводск': 'Республика Карелия', 'Нижний Новгород': 'Нижегородская область',
    'Псков': 'Псковская область', 'Новосибирск': 'Новосибирская область', 'Омск': 'Омская область',
    'Сыктывкар': 'Республика Коми', 'Пермь': 'Пермский край', 'Самара': 'Самарская область',
    'Тюмень': 'Тюменская область', 'Уфа': 'Республика Башкортостан', 'Челябинск': 'Челябинская область',
    'Ростов-на Дону': 'Ростовская область'
  };
  const norm = function(value) {
    return String(value || '').toLowerCase().replace(/ё/g, 'е').replace(/[^a-zа-я0-9]+/g, ' ').trim();
  };
  const selected = {};
  enabledBase.forEach(function(name) { selected[norm(aliases[name] || name)] = name; });
  const leningrad = new Set(['кириши','токсово','выборг','кузьмоловский','тосно','сосновый бор','федоровское','новодевяткинское','коммунар','вырица','шушары','всеволожск','парголово','сестрорецк','пушкин','ломоносов','кудрово','мурино','гатчина','колпино','кронштадт','петергоф','сланцы','кингисепп','луга','волхов','тихвин','приозерск','сертолово','отрадное','никольское','кировск','светогорск','подпорожье','пикалево','бокситогорск']);
  const kazakhstan = new Set(['казахстан','караганда','алматы','усть каменогорск','семей','астана','павлодар','темиртау','костанай','конаев','экибастуз','шымкент','петропавловск','аягоз','рудный','жезказган','актау','актобе','атырау','кокшетау','тараз','талдыкорган','уральск','кызылорда']);

  return {
    classify: function(rawName) {
      const key = norm(rawName);
      if (!key || key === 'не определен' || key === 'не определено') return 'Прочие регионы';
      if (key === 'санкт петербург' || key === 'санкт петербург и ленинградская область') return 'Санкт-Петербург';
      if (key === 'ленинградская область' || leningrad.has(key)) return 'Ленинградская область';
      if (kazakhstan.has(key)) return 'Казахстан';
      const canonical = norm(directRegionKey_(rawName));
      return selected[canonical] || selected[key] || 'Прочие регионы';
    }
  };
}

function tdmReconcileRegionTraffic20260713_(byRegion, expected) {
  if (!byRegion['Прочие регионы']) byRegion['Прочие регионы'] = {
    visits: 0, users: 0, impressions: 0, clicks: 0, cost: 0, bounceWeighted: 0, depthWeighted: 0
  };
  ['impressions', 'clicks', 'cost'].forEach(function(key) {
    const actual = Object.keys(byRegion).reduce(function(sum, region) {
      return sum + Number(byRegion[region][key] || 0);
    }, 0);
    const delta = Number(expected[key] || 0) - actual;
    if (Math.abs(delta) < (key === 'cost' ? 0.005 : 0.0001)) return;
    byRegion['Прочие регионы'][key] += delta;
    if (byRegion['Прочие регионы'][key] < -0.001) {
      throw new Error('Отчет  регионы: нельзя безопасно сверить ' + key + ', delta=' + delta);
    }
  });
}

function tdmReconcileRegionMetric20260713_(stats, key, expected) {
  if (!stats['Прочие регионы']) stats['Прочие регионы'] = {};
  let actual = Object.keys(stats).reduce(function(sum, region) {
    return sum + Number((stats[region] || {})[key] || 0);
  }, 0);
  let delta = Number(expected || 0) - actual;
  if (Math.abs(delta) < 0.0001) return;
  if (delta > 0) {
    stats['Прочие регионы'][key] = Number(stats['Прочие регионы'][key] || 0) + delta;
    return;
  }
  let remaining = -delta;
  Object.keys(stats).sort(function(a, b) {
    return Number(stats[b][key] || 0) - Number(stats[a][key] || 0);
  }).forEach(function(region) {
    if (remaining <= 0) return;
    const value = Number(stats[region][key] || 0);
    const take = Math.min(value, remaining);
    stats[region][key] = value - take;
    remaining -= take;
  });
  if (remaining > 0.0001) throw new Error('Отчет  регионы: не удалось сверить цель ' + key + '.');
}

function tdmCallibriRegionBuckets20260713_(emo, block, classifier) {
  const header = tdmEnoMayMonthlyV3HeaderMap_(emo, block.headerRow);
  const values = emo.getRange(block.headerRow + 1, 1, block.totalRow - block.headerRow - 1, emo.getLastColumn()).getValues();
  const result = {};
  values.forEach(function(row) {
    const campaign = String(row[header.rk - 1] || '');
    if (!campaign) return;
    let bucket = 'Прочие регионы';
    const key = campaign.toLowerCase().replace(/-/g, '_');
    if (/(^|_)kz($|_)/.test(key)) bucket = 'Казахстан';
    else if (/(^|_)spb($|_)/.test(key) && key.indexOf('spb_rf') === -1 && key.indexOf('rf_spb') === -1) bucket = 'Санкт-Петербург';
    else if (!/(^|_)rf($|_)/.test(key)) bucket = classifier.classify(campaign);
    if (!result[bucket]) result[bucket] = { spam: 0, unqualified: 0, qualA: 0, qualC: 0 };
    result[bucket].spam += Number(row[header.callibriSpam - 1]) || 0;
    result[bucket].unqualified += Number(row[header.callibriNonTarget - 1]) || 0;
    result[bucket].qualA += Number(row[header.callibriLeadA - 1]) || 0;
    result[bucket].qualC += Number(row[header.callibriLeadC - 1]) || 0;
  });
  return result;
}

function tdmWriteRegionSummaryComments20260713_(sheet, totalRow, rows, dateFrom, dateTo) {
  const startRow = totalRow + 2;
  const withLeads = rows.filter(function(row) { return Number(row[22]) > 0; }).slice(0, 8);
  const withoutLeads = rows.filter(function(row) { return Number(row[6]) > 0 && Number(row[22]) === 0; }).slice(0, 8);
  const comments = [
    ['Комментарии по регионам за ' + tdmUnifiedRuDate_(dateFrom) + '–' + tdmUnifiedRuDate_(dateTo)],
    [''],
    ['Регионы с целевыми обращениями:']
  ];
  withLeads.forEach(function(row) { comments.push([row[0] + ': ' + row[25]]); });
  comments.push(['']);
  comments.push(['Регионы с расходом без целевых обращений:']);
  withoutLeads.forEach(function(row) { comments.push([row[0] + ': ' + row[25]]); });
  comments.push(['']);
  comments.push(['Фокус: проверять качество Callibri/Jivo отдельно; спам и нецелевые не включать в целевой факт.']);

  sheet.getRange(startRow, 1, comments.length, 10)
    .clearContent().setBackground('#ffffff').setFontColor('#000000').setFontWeight('normal').setWrap(true);
  sheet.getRange(startRow, 1, comments.length, 1).setValues(comments);
  [0, 2, 4 + withLeads.length].forEach(function(offset) {
    if (offset >= comments.length) return;
    sheet.getRange(startRow + offset, 1, 1, 10).setBackground('#d9ead3').setFontWeight('bold');
  });
}

function tdmValidateMondayReports20260713_(week, month, eno, emo, regions) {
  const dbWeek = tdmUnifiedDbTotals_(week.dateFrom, week.dateTo);
  const dbMonth = tdmUnifiedDbTotals_(month.dateFrom, month.dateTo);
  const enoReport = tdmUnifiedReportTotals_('ЕНО', week.dateFrom, week.dateTo, 'week');
  const emoReport = tdmUnifiedReportTotals_('ЕМО', month.dateFrom, month.dateTo, 'month');
  const regionReport = tdmUnifiedReportTotals_('Отчет  регионы', month.dateFrom, month.dateTo, 'month');

  tdmAssertMondayTotals20260713_('Директ/Callibri ↔ ДБ 06–12', eno.totals, dbWeek);
  tdmAssertMondayTotals20260713_('ДБ ↔ ЕНО 06–12', dbWeek, enoReport);
  tdmAssertMondayTotals20260713_('Директ/Callibri ↔ ДБ 01–12', emo.totals, dbMonth);
  tdmAssertMondayTotals20260713_('ДБ ↔ ЕМО 01–12', dbMonth, emoReport);
  tdmAssertMondayTotals20260713_('ЕМО ↔ Отчет регионы 01–12', emoReport, regionReport);
  tdmAssertMondayTotals20260713_('Возврат функции регионов ↔ лист', regions.totals, regionReport);

  return {
    ok: true,
    week: { source: eno.totals, db: dbWeek, report: enoReport },
    month: { source: emo.totals, db: dbMonth, emo: emoReport, regions: regionReport }
  };
}

function tdmAssertMondayTotals20260713_(label, expected, actual) {
  const fields = ['impressions', 'clicks', 'cost', 'view3', 'spam', 'nonTarget', 'addToCart', 'purchase', 'jivo', 'leadA', 'leadB', 'leadC', 'fact'];
  const diffs = [];
  fields.forEach(function(field) {
    const a = Number(expected[field] || 0);
    const b = Number(actual[field] || 0);
    // Сумма округлённых дневных расходов ДБ может отличаться от агрегата
    // Директа на несколько копеек. Для всех лидов и целей допуск остаётся нулевым.
    const tolerance = field === 'cost' ? 0.05 : 0.0001;
    if (Math.abs(a - b) > tolerance) diffs.push({ field: field, expected: a, actual: b, diff: b - a });
  });
  if (diffs.length) throw new Error(label + ': ' + JSON.stringify(diffs));
  return true;
}

function tdmMondayTotalsFromTdme20260713_(total) {
  const goals = (total || {}).goals || {};
  return {
    impressions: Number(total.impressions || 0),
    clicks: Number(total.clicks || 0),
    cost: Number(total.cost || 0),
    view3: Number(goals.view3Pages || 0),
    spam: Number(goals.callibriSpam || 0),
    nonTarget: Number(goals.callibriNonTarget || 0),
    addToCart: Number(goals.addToCart || 0),
    purchase: Number(goals.purchase || 0),
    jivo: Number(goals.jivo || 0),
    leadA: Number(goals.callibriLeadA || 0),
    leadB: Number(goals.callibriLeadB || 0),
    leadC: Number(goals.callibriLeadC || 0),
    fact: tdmeTargetConversions_(total)
  };
}

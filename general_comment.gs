/**
 * ТДМ — автоматический общий комментарий.
 *
 * Источники:
 * - ЕНО за прошлую полную неделю;
 * - предыдущий недельный блок ЕНО для сравнения;
 * - Callibri_Сверка для контроля неразмеченных обращений;
 * - Отчет  регионы, только если его период заканчивается той же датой.
 */
const TDM_GENERAL_COMMENT_20260727 = {
  spreadsheetId: '1zXDDfiRYHJE34iOJg-QfAC0bvoTKMIAuVMfsJS0z5Dg',
  targetSheet: 'Общий комментарий',
  enoSheet: 'ЕНО',
  regionSheet: 'Отчет  регионы',
  callibriSheet: 'Callibri_Сверка',
  auditSheet: '_GPT Weekly Audit',
  timezone: 'Europe/Moscow',
  targetCpa: 2300
};

/** Вызывается из единого понедельничного контура после ЕНО/ЕМО/регионов. */
function tdmBuildGeneralComment20260727_(week, month, eno, emo, regions) {
  if (!week || !week.dateFrom || !week.dateTo) {
    throw new Error('Общий комментарий: не передан недельный период.');
  }
  return tdmBuildGeneralCommentForPeriod20260727_(week.dateFrom, week.dateTo, eno && eno.block);
}

/** Ручной безопасный запуск по последнему недельному блоку ЕНО. */
function tdmBuildGeneralCommentNow20260727() {
  const ss = SpreadsheetApp.openById(TDM_GENERAL_COMMENT_20260727.spreadsheetId);
  const eno = ss.getSheetByName(TDM_GENERAL_COMMENT_20260727.enoSheet);
  if (!eno) throw new Error('Общий комментарий: не найден лист ЕНО.');
  const latest = tdmGcFindLatestWeeklyBlock20260727_(eno);
  return tdmBuildGeneralCommentForPeriod20260727_(latest.dateFrom, latest.dateTo, latest);
}

function tdmBuildGeneralCommentForPeriod20260727_(dateFrom, dateTo, knownBlock) {
  const cfg = TDM_GENERAL_COMMENT_20260727;
  const ss = SpreadsheetApp.openById(cfg.spreadsheetId);
  const enoSheet = ss.getSheetByName(cfg.enoSheet);
  const target = ss.getSheetByName(cfg.targetSheet);
  if (!enoSheet || !target) throw new Error('Общий комментарий: не найдены обязательные листы.');

  const currentBlock = knownBlock && knownBlock.headerRow && knownBlock.totalRow
    ? {
        topRow: Number(knownBlock.topRow || 0),
        headerRow: Number(knownBlock.headerRow),
        totalRow: Number(knownBlock.totalRow),
        dateFrom: dateFrom,
        dateTo: dateTo
      }
    : tdmGcFindWeeklyBlockByPeriod20260727_(enoSheet, dateFrom, dateTo);

  const current = tdmGcReadEnoBlock20260727_(enoSheet, currentBlock);
  const previousBlock = tdmGcFindPreviousWeeklyBlock20260727_(enoSheet, currentBlock.topRow || currentBlock.headerRow);
  const previous = previousBlock ? tdmGcReadEnoBlock20260727_(enoSheet, previousBlock) : null;
  const quality = tdmGcReadCallibriQuality20260727_(ss, dateFrom, dateTo);
  const region = tdmGcReadRegionSummary20260727_(ss, dateTo);

  const comments = tdmGcBuildTexts20260727_(dateFrom, dateTo, current, previous, quality, region);
  const stamp = 'Обновлено автоматически: ' + Utilities.formatDate(new Date(), cfg.timezone, 'dd.MM.yyyy HH:mm');

  target.getRange('A2').setValue(stamp);
  target.getRange('C2').setValue(stamp);
  target.getRange('A4').setValue(comments.internal);
  target.getRange('C4').setValue(comments.client);
  SpreadsheetApp.flush();

  if (String(target.getRange('A4').getDisplayValue() || '').indexOf(tdmGcRuDate20260727_(dateTo)) === -1) {
    throw new Error('Общий комментарий: после записи не найден конец периода ' + dateTo + '.');
  }

  const result = {
    ok: true,
    dateFrom: dateFrom,
    dateTo: dateTo,
    current: current.total,
    previous: previous ? previous.total : null,
    quality: quality,
    regionsIncluded: region.current,
    updatedAt: stamp
  };
  tdmGcWriteAudit20260727_(ss, result);
  return result;
}

function tdmGcBuildTexts20260727_(dateFrom, dateTo, current, previous, quality, region) {
  const cfg = TDM_GENERAL_COMMENT_20260727;
  const total = current.total;
  const period = tdmGcRuDate20260727_(dateFrom) + '–' + tdmGcRuDate20260727_(dateTo);
  const overallCpa = total.fact > 0 ? total.cost / total.fact : 0;
  const effectiveLimit = Math.max(overallCpa || 0, cfg.targetCpa);

  const effective = current.rows.filter(function(row) {
    return row.fact >= 2 && row.cpa > 0 && row.cpa <= effectiveLimit;
  }).sort(function(a, b) {
    return b.fact - a.fact || a.cpa - b.cpa;
  }).slice(0, 5);

  const attention = current.rows.filter(function(row) {
    return (row.fact === 0 && row.cost >= 2000) ||
      (row.fact > 0 && row.cpa > cfg.targetCpa * 1.5);
  }).sort(function(a, b) {
    return b.cost - a.cost;
  }).slice(0, 6);

  const internal = [];
  internal.push('ИТОГИ НЕДЕЛИ ' + period);
  internal.push('');
  internal.push('Факт сумма конверсий — ' + tdmGcInt20260727_(total.fact) + '. Расход — ' +
    tdmGcRub20260727_(total.cost) + '. CPA — ' + tdmGcRub20260727_(overallCpa) +
    '. Рабочий ориентир CPA — ' + tdmGcRub20260727_(cfg.targetCpa) + '.');

  if (previous && previous.total) {
    const prev = previous.total;
    const prevCpa = prev.fact > 0 ? prev.cost / prev.fact : 0;
    internal.push('');
    internal.push('По сравнению с предыдущей неделей: конверсии — ' + tdmGcInt20260727_(total.fact) +
      ' против ' + tdmGcInt20260727_(prev.fact) + ' (' + tdmGcSignedPct20260727_(total.fact, prev.fact) +
      '); расход — ' + tdmGcSignedPct20260727_(total.cost, prev.cost) + '; CPA — ' +
      tdmGcSignedPct20260727_(overallCpa, prevCpa) + '.');
  }

  internal.push('');
  internal.push('ЭФФЕКТИВНОСТЬ КАМПАНИЙ');
  internal.push('');
  if (effective.length) {
    internal.push('Основной результат:');
    effective.forEach(function(row) {
      internal.push('• ' + row.name + ' — ' + tdmGcInt20260727_(row.fact) +
        ' конверсий, CPA ' + tdmGcRub20260727_(row.cpa) + '.');
    });
  } else {
    internal.push('Кампаний с устойчивым результатом по текущему объёму пока недостаточно.');
  }

  internal.push('');
  internal.push('Требуют внимания:');
  if (attention.length) {
    attention.forEach(function(row) {
      internal.push('• ' + row.name + ' — ' + (row.fact > 0
        ? tdmGcInt20260727_(row.fact) + ' конверсий, CPA ' + tdmGcRub20260727_(row.cpa)
        : 'расход ' + tdmGcRub20260727_(row.cost) + ' без целевых конверсий') + '.');
    });
  } else {
    internal.push('• Явных зон перерасхода по заданным порогам не выявлено.');
  }

  internal.push('');
  internal.push('КАЧЕСТВО ОБРАЩЕНИЙ');
  internal.push('');
  internal.push('Callibri: лид A — ' + quality.leadA + ', лид B — ' + quality.leadB +
    ', лид C — ' + quality.leadC + ', нецелевые — ' + quality.nonTarget +
    ', спам — ' + quality.spam + '. Обращения без класса — ' + quality.unknown +
    '; они не включены в целевой факт до проверки.');

  internal.push('');
  internal.push('РЕГИОНЫ');
  internal.push('');
  if (region.current) {
    if (region.good.length) internal.push('Стабильно эффективные: ' + tdmGcJoinNamed20260727_(region.good) + '.');
    if (region.attention.length) internal.push('Требуют внимания: ' + tdmGcJoinNamed20260727_(region.attention) + '.');
  } else {
    internal.push('Региональный лист не обновлён до ' + tdmGcRuDate20260727_(dateTo) +
      ', поэтому региональные выводы в текущий комментарий не включены.');
  }

  internal.push('');
  internal.push('РАБОЧИЕ ДЕЙСТВИЯ');
  internal.push('');
  internal.push('• Сохранять и аккуратно усиливать кампании с CPA не выше рабочего ориентира.');
  internal.push('• Ограничивать расход и проверять запросы/объявления в кампаниях без целевых конверсий.');
  if (quality.unknown > 0) internal.push('• Проверить и классифицировать обращения Callibri без категории — ' + quality.unknown + '.');
  if (!region.current) internal.push('• Обновить региональный отчёт до ' + tdmGcRuDate20260727_(dateTo) + '.');

  const client = [];
  client.push('РЕЗУЛЬТАТЫ НЕДЕЛИ ' + period);
  client.push('');
  client.push('Получено целевых конверсий — ' + tdmGcInt20260727_(total.fact) +
    '. Расход — ' + tdmGcRub20260727_(total.cost) +
    '. Средняя стоимость целевой конверсии — ' + tdmGcRub20260727_(overallCpa) + '.');

  if (previous && previous.total) {
    const prev = previous.total;
    const prevCpa = prev.fact > 0 ? prev.cost / prev.fact : 0;
    client.push('');
    client.push('По сравнению с предыдущей неделей количество конверсий изменилось с ' +
      tdmGcInt20260727_(prev.fact) + ' до ' + tdmGcInt20260727_(total.fact) +
      ', расход изменился на ' + tdmGcSignedPct20260727_(total.cost, prev.cost) +
      ', стоимость конверсии — на ' + tdmGcSignedPct20260727_(overallCpa, prevCpa) + '.');
  }

  client.push('');
  client.push('ЧТО РАБОТАЕТ ХОРОШО');
  client.push('');
  if (effective.length) {
    const fact = effective.reduce(function(sum, row) { return sum + row.fact; }, 0);
    client.push('Основной объём результата обеспечили наиболее эффективные кампании: суммарно ' +
      tdmGcInt20260727_(fact) + ' конверсий при стоимости ниже или на уровне общего результата недели.');
  } else {
    client.push('Статистики пока недостаточно, чтобы выделить устойчивую группу кампаний для масштабирования.');
  }

  client.push('');
  client.push('ЧТО ТРЕБУЕТ ВНИМАНИЯ');
  client.push('');
  client.push(attention.length
    ? 'По части кампаний зафиксирован расход без целевых конверсий или повышенная стоимость обращения. Продолжаем оптимизацию и ограничиваем неэффективный расход без резких изменений.'
    : 'Явного перерасхода по текущим рабочим порогам не выявлено.');

  client.push('');
  client.push('КАЧЕСТВО ОБРАЩЕНИЙ');
  client.push('');
  client.push('По Callibri зафиксировано квалифицированных обращений — ' +
    (quality.leadA + quality.leadB + quality.leadC) + ', нецелевых — ' + quality.nonTarget +
    ', спам — ' + quality.spam + ', без категории — ' + quality.unknown +
    '. Обращения без категории не включены в целевой результат до проверки.');

  client.push('');
  client.push('ФОКУС НА СЛЕДУЮЩУЮ НЕДЕЛЮ');
  client.push('');
  client.push('• Поддерживать эффективные кампании.');
  client.push('• Продолжить оптимизацию кампаний с повышенной стоимостью обращения.');
  if (quality.unknown > 0) client.push('• Проверить обращения Callibri без категории.');
  client.push(region.current
    ? '• Сохранить рабочую географию и отдельно контролировать регионы с повышенным CPA.'
    : '• Обновить региональную статистику и скорректировать географию после получения полного среза.');

  return { internal: internal.join('\n'), client: client.join('\n') };
}

function tdmGcReadEnoBlock20260727_(sheet, block) {
  const headerValues = sheet.getRange(block.headerRow, 1, 1, sheet.getLastColumn()).getDisplayValues()[0];
  const map = tdmGcHeaderMap20260727_(headerValues);
  if (!map.name || !map.cost || !map.fact) {
    throw new Error('Общий комментарий: в ЕНО не найдены РК, расход или факт конверсий.');
  }

  const rowCount = Math.max(block.totalRow - block.headerRow - 1, 0);
  const values = rowCount ? sheet.getRange(block.headerRow + 1, 1, rowCount, sheet.getLastColumn()).getDisplayValues() : [];
  const rows = values.map(function(row) {
    const name = String(row[map.name - 1] || '').trim();
    const cost = tdmGcNum20260727_(row[map.cost - 1]);
    const fact = tdmGcNum20260727_(row[map.fact - 1]);
    return { name: name, cost: cost, fact: fact, cpa: fact > 0 ? cost / fact : 0 };
  }).filter(function(row) { return row.name && tdmGcNorm20260727_(row.name) !== 'итого'; });

  const totalRow = sheet.getRange(block.totalRow, 1, 1, sheet.getLastColumn()).getDisplayValues()[0];
  const total = {
    cost: tdmGcNum20260727_(totalRow[map.cost - 1]),
    fact: tdmGcNum20260727_(totalRow[map.fact - 1]),
    spam: map.spam ? tdmGcNum20260727_(totalRow[map.spam - 1]) : 0,
    nonTarget: map.nonTarget ? tdmGcNum20260727_(totalRow[map.nonTarget - 1]) : 0,
    leadA: map.leadA ? tdmGcNum20260727_(totalRow[map.leadA - 1]) : 0,
    leadB: map.leadB ? tdmGcNum20260727_(totalRow[map.leadB - 1]) : 0,
    leadC: map.leadC ? tdmGcNum20260727_(totalRow[map.leadC - 1]) : 0
  };
  return { block: block, rows: rows, total: total };
}

function tdmGcHeaderMap20260727_(headers) {
  const result = {};
  headers.forEach(function(value, index) {
    const n = tdmGcNorm20260727_(value);
    const col = index + 1;
    if (n === 'рк') result.name = col;
    else if (n.indexOf('расход c ндс') !== -1 || n.indexOf('расход с ндс') !== -1) result.cost = col;
    else if (n.indexOf('факт сумма конверсий') !== -1) result.fact = col;
    else if (n.indexOf('callibri спам') !== -1) result.spam = col;
    else if (n.indexOf('callibri нецелевой лид') !== -1) result.nonTarget = col;
    else if (n.indexOf('callibri лид квал a') !== -1) result.leadA = col;
    else if (n.indexOf('callibri лид квал b') !== -1) result.leadB = col;
    else if (n.indexOf('callibri лид квал c') !== -1) result.leadC = col;
  });
  return result;
}

function tdmGcFindLatestWeeklyBlock20260727_(sheet) {
  const values = sheet.getRange(1, 1, sheet.getLastRow(), 1).getDisplayValues();
  for (let i = values.length - 1; i >= 0; i--) {
    const period = tdmGcParseWeeklyTitle20260727_(values[i][0]);
    if (!period) continue;
    return tdmGcResolveBlock20260727_(sheet, i + 1, period.dateFrom, period.dateTo);
  }
  throw new Error('Общий комментарий: в ЕНО не найден недельный блок.');
}

function tdmGcFindWeeklyBlockByPeriod20260727_(sheet, dateFrom, dateTo) {
  const values = sheet.getRange(1, 1, sheet.getLastRow(), 1).getDisplayValues();
  for (let i = values.length - 1; i >= 0; i--) {
    const period = tdmGcParseWeeklyTitle20260727_(values[i][0]);
    if (period && period.dateFrom === dateFrom && period.dateTo === dateTo) {
      return tdmGcResolveBlock20260727_(sheet, i + 1, dateFrom, dateTo);
    }
  }
  throw new Error('Общий комментарий: в ЕНО не найден период ' + dateFrom + '–' + dateTo + '.');
}

function tdmGcFindPreviousWeeklyBlock20260727_(sheet, beforeRow) {
  if (!beforeRow || beforeRow <= 1) return null;
  const values = sheet.getRange(1, 1, beforeRow - 1, 1).getDisplayValues();
  for (let i = values.length - 1; i >= 0; i--) {
    const period = tdmGcParseWeeklyTitle20260727_(values[i][0]);
    if (!period) continue;
    return tdmGcResolveBlock20260727_(sheet, i + 1, period.dateFrom, period.dateTo);
  }
  return null;
}

function tdmGcResolveBlock20260727_(sheet, topRow, dateFrom, dateTo) {
  let headerRow = 0;
  let totalRow = 0;
  const end = Math.min(sheet.getLastRow(), topRow + 80);
  const values = sheet.getRange(topRow, 1, end - topRow + 1, Math.min(sheet.getLastColumn(), 30)).getDisplayValues();
  for (let i = 0; i < values.length; i++) {
    const norm = values[i].map(tdmGcNorm20260727_);
    if (!headerRow && norm.indexOf('рк') !== -1 && norm.indexOf('клики') !== -1 && norm.join(' ').indexOf('факт сумма конверсий') !== -1) {
      headerRow = topRow + i;
      continue;
    }
    if (headerRow && norm.indexOf('итого') !== -1) {
      totalRow = topRow + i;
      break;
    }
  }
  if (!headerRow || !totalRow) throw new Error('Общий комментарий: повреждён блок ЕНО со строки ' + topRow + '.');
  return { topRow: topRow, headerRow: headerRow, totalRow: totalRow, dateFrom: dateFrom, dateTo: dateTo };
}

function tdmGcParseWeeklyTitle20260727_(value) {
  const text = String(value || '').trim();
  if (tdmGcNorm20260727_(text).indexOf('отчет недельный по рк') !== 0) return null;
  const dates = text.match(/\d{1,2}\.\d{1,2}\.\d{4}/g) || [];
  if (dates.length < 2) return null;
  return { dateFrom: tdmGcApiDate20260727_(dates[0]), dateTo: tdmGcApiDate20260727_(dates[1]) };
}

function tdmGcReadCallibriQuality20260727_(ss, dateFrom, dateTo) {
  const sheet = ss.getSheetByName(TDM_GENERAL_COMMENT_20260727.callibriSheet);
  const result = { total: 0, leadA: 0, leadB: 0, leadC: 0, spam: 0, nonTarget: 0, unknown: 0 };
  if (!sheet || sheet.getLastRow() < 2) return result;
  const values = sheet.getRange(2, 1, sheet.getLastRow() - 1, 7).getDisplayValues();
  values.forEach(function(row) {
    const date = tdmGcAnyDateToApi20260727_(row[0]);
    if (!date || date < dateFrom || date > dateTo) return;
    result.total += tdmGcNum20260727_(row[1]);
    result.leadA += tdmGcNum20260727_(row[2]);
    result.leadC += tdmGcNum20260727_(row[3]);
    result.spam += tdmGcNum20260727_(row[4]);
    result.nonTarget += tdmGcNum20260727_(row[5]);
  });
  result.unknown = Math.max(0, result.total - result.leadA - result.leadB - result.leadC - result.spam - result.nonTarget);
  return result;
}

function tdmGcReadRegionSummary20260727_(ss, dateTo) {
  const sheet = ss.getSheetByName(TDM_GENERAL_COMMENT_20260727.regionSheet);
  const result = { current: false, good: [], attention: [] };
  if (!sheet) return result;
  const title = String(sheet.getRange('A1').getDisplayValue() || '');
  if (title.indexOf(tdmGcRuDate20260727_(dateTo)) === -1) return result;

  const values = sheet.getDataRange().getDisplayValues();
  let headerRow = -1;
  for (let i = 0; i < Math.min(values.length, 15); i++) {
    if (tdmGcNorm20260727_(values[i][0]) === 'регион' && values[i].map(tdmGcNorm20260727_).join(' ').indexOf('факт сумма конверсий') !== -1) {
      headerRow = i;
      break;
    }
  }
  if (headerRow < 0) return result;
  const headers = values[headerRow];
  const nameCol = headers.findIndex(function(v) { return tdmGcNorm20260727_(v) === 'регион'; });
  const costCol = headers.findIndex(function(v) { const n = tdmGcNorm20260727_(v); return n.indexOf('расход c ндс') !== -1 || n.indexOf('расход с ндс') !== -1; });
  const factCol = headers.findIndex(function(v) { return tdmGcNorm20260727_(v).indexOf('факт сумма конверсий') !== -1; });
  const cpaCol = headers.findIndex(function(v) { return tdmGcNorm20260727_(v).indexOf('стоимость факт cpa') !== -1; });
  if (nameCol < 0 || costCol < 0 || factCol < 0 || cpaCol < 0) return result;

  const rows = [];
  let totalCpa = 0;
  for (let i = headerRow + 1; i < values.length; i++) {
    const name = String(values[i][nameCol] || '').trim();
    if (!name) continue;
    if (tdmGcNorm20260727_(name) === 'итого') {
      totalCpa = tdmGcNum20260727_(values[i][cpaCol]);
      break;
    }
    const cost = tdmGcNum20260727_(values[i][costCol]);
    const fact = tdmGcNum20260727_(values[i][factCol]);
    const cpa = fact > 0 ? tdmGcNum20260727_(values[i][cpaCol]) : 0;
    rows.push({ name: name, cost: cost, fact: fact, cpa: cpa });
  }
  if (!totalCpa) totalCpa = 2300;
  result.current = true;
  result.good = rows.filter(function(row) { return row.fact >= 2 && row.cpa > 0 && row.cpa <= totalCpa; })
    .sort(function(a, b) { return b.fact - a.fact || a.cpa - b.cpa; }).slice(0, 5);
  result.attention = rows.filter(function(row) {
    return (row.fact === 0 && row.cost >= 2000) || (row.fact > 0 && row.cpa > totalCpa * 2);
  }).sort(function(a, b) { return b.cost - a.cost; }).slice(0, 6);
  return result;
}

function tdmGcWriteAudit20260727_(ss, result) {
  const sheet = ss.getSheetByName(TDM_GENERAL_COMMENT_20260727.auditSheet);
  if (!sheet) return;
  sheet.appendRow([
    Utilities.formatDate(new Date(), TDM_GENERAL_COMMENT_20260727.timezone, 'yyyy-MM-dd HH:mm:ss'),
    result.dateFrom,
    result.dateTo,
    'TDM_2026',
    'ТДМ_2026',
    'OK_GENERAL_COMMENT',
    'ЕНО, Callibri_Сверка, Отчет  регионы',
    'Общий комментарий обновлён за ' + result.dateFrom + '–' + result.dateTo,
    result.quality.unknown > 0 ? 'Проверить Callibri без класса: ' + result.quality.unknown : 'Дополнительных действий по Callibri нет',
    JSON.stringify(result.current),
    JSON.stringify(result)
  ]);
}

function tdmGcJoinNamed20260727_(rows) {
  return rows.map(function(row) {
    return row.name + ' — ' + row.fact + ', CPA ' + tdmGcRub20260727_(row.cpa || 0);
  }).join('; ');
}

function tdmGcSignedPct20260727_(current, previous) {
  current = Number(current || 0);
  previous = Number(previous || 0);
  if (!previous) return current ? '+100,0%' : '0,0%';
  const value = (current - previous) / previous * 100;
  return (value > 0 ? '+' : '') + value.toFixed(1).replace('.', ',') + '%';
}

function tdmGcInt20260727_(value) {
  return Math.round(Number(value || 0)).toLocaleString('ru-RU');
}

function tdmGcRub20260727_(value) {
  return Number(value || 0).toLocaleString('ru-RU', { minimumFractionDigits: 2, maximumFractionDigits: 2 }) + ' руб.';
}

function tdmGcNum20260727_(value) {
  const text = String(value == null ? '' : value)
    .replace(/[₽р.руб\s\u00A0]/g, '')
    .replace(',', '.')
    .replace(/[^0-9.\-]/g, '');
  return Number(text) || 0;
}

function tdmGcNorm20260727_(value) {
  return String(value == null ? '' : value).toLowerCase().replace(/ё/g, 'е').replace(/[^a-zа-я0-9]+/g, ' ').trim();
}

function tdmGcRuDate20260727_(apiDate) {
  const p = String(apiDate || '').split('-');
  return p.length === 3 ? p[2] + '.' + p[1] + '.' + p[0] : String(apiDate || '');
}

function tdmGcApiDate20260727_(ruDate) {
  const p = String(ruDate || '').split('.');
  return p.length === 3 ? p[2] + '-' + ('0' + p[1]).slice(-2) + '-' + ('0' + p[0]).slice(-2) : '';
}

function tdmGcAnyDateToApi20260727_(value) {
  if (Object.prototype.toString.call(value) === '[object Date]' && !isNaN(value)) {
    return Utilities.formatDate(value, TDM_GENERAL_COMMENT_20260727.timezone, 'yyyy-MM-dd');
  }
  const text = String(value || '').trim();
  if (/^\d{4}-\d{2}-\d{2}$/.test(text)) return text;
  if (/^\d{1,2}\.\d{1,2}\.\d{4}$/.test(text)) return tdmGcApiDate20260727_(text);
  return '';
}

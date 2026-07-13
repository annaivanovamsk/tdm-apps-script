// =====================
// ТДМ — ДБ ежедневный отчёт
// Файл: db.gs
// Версия v21_prev_structure_april_format
// =====================
//
// ОБЩЕЕ ПРАВИЛО ДЛЯ ОТЧЁТОВ:
// - Новый месяц вставляем после предыдущего месяца.
// - Структуру нового месяца берём из предыдущего месяца.
// - Оформление строк дат берём из Апрель 2026, там корректное оформление.
// - Формулы не перезаписываем значениями.
// - Исключение: плановая колонка "Расход План c НДС" при создании месяца очищается и заполняется новыми формулами.
// - Строка ИТОГО всегда тянется до последнего дня текущего месяца.
// - Рабочие дни без заливки, выходные/праздники — как в апрельском шаблоне.
// - Расход План c НДС: формула только в рабочие дни, делитель = количество рабочих дней.
// - Деньги отображаем через "р.", без символа ₽.
// =====================


// =====================
// ПУБЛИЧНЫЕ ФУНКЦИИ
// В выпадающем списке будут только они.
// =====================

function fillTdmDbYesterday() {
  TDM_DB.fillYesterday();
}

function fillTdmDbToYesterday() {
  TDM_DB.fillToYesterday();
}

function fillTdmDbTestOneDay() {
  TDM_DB.fillTestOneDay();
}

// Удали блок Июнь 2026 вручную и запусти эту функцию.
function createTdmDbJune2026Block() {
  TDM_DB.createOrFixMonth('2026-06-01');
}

function archived_createTdmDbTrigger8am() {
  TDM_DB.createTrigger8am();
}

function deleteTdmDbTrigger() {
  TDM_DB.deleteTrigger();
}


// =====================
// ВНУТРЕННИЙ МОДУЛЬ
// =====================

const TDM_DB = {
  config: {
    clientLogin: 'ilhaleontj3v',
    sheetName: 'ДБ',
    timezone: 'Europe/Moscow',

    startDate: '2026-05-01',
    testDate: '2026-05-31',

    // Эталон оформления и структуры для новых месяцев.
    templateMonthDate: '2026-04-01',

    // План бюджета с НДС для новых месяцев.
    monthPlanWithVat: 210000,

    // Праздники / нерабочие дни.
    holidays: [
      '2026-06-12'
    ],

    goals: {
      view3Pages: ['453453800'],      // Просмотр 3х страниц — микроцель
      siteSearch: ['410522803'],      // Поиск по сайту — микроцель
      addToCart: ['504318736'],       // Ecommerce: добавление в корзину
      purchase: ['504318735'],        // Ecommerce: покупка — факт
      jivo: ['575188424'],            // Jivo-сайт: пользователь начал чат — факт
      callibriSpam: [],               // Callibri: Спам — источник пока не подключён
      callibriNonTarget: [],          // Callibri: Нецелевой_Лид — источник пока не подключён
      callibriLeadA: [],              // Callibri: Лид_Квал_A — источник пока не подключён
      callibriLeadC: []               // Callibri: Лид_Квал_C — источник пока не подключён
    }
  },


  // =====================
  // ЗАПУСКИ
  // =====================

  fillYesterday() {
    const yesterday = this.getYesterday_();
    this.fillPeriod_(yesterday, yesterday);
  },

  fillToYesterday() {
    const yesterday = this.getYesterday_();
    this.fillPeriod_(this.config.startDate, yesterday);
  },

  fillTestOneDay() {
    this.fillPeriod_(this.config.testDate, this.config.testDate);
  },

  createOrFixMonth(apiDate) {
    const sheet = this.getSheet_();

    this.ensureMonthBlock_(sheet, apiDate);
    this.fixExistingMonthBlock_(sheet, apiDate);
    this.validateMonth_(sheet, apiDate);

    this.safeMessage_('Блок ' + this.monthTitle_(apiDate) + ' создан / исправлен.');
  },

  createTrigger8am() {
    this.deleteTrigger();

    ScriptApp.newTrigger('fillTdmDbYesterday')
      .timeBased()
      .everyDays(1)
      .atHour(8)
      .inTimezone(this.config.timezone)
      .create();

    this.safeMessage_('Триггер ДБ поставлен: каждый день около 08:00 по Москве.');
  },

  deleteTrigger() {
    const triggers = ScriptApp.getProjectTriggers();

    triggers.forEach(trigger => {
      if (trigger.getHandlerFunction() === 'fillTdmDbYesterday') {
        ScriptApp.deleteTrigger(trigger);
      }
    });

    this.safeMessage_('Триггеры ДБ удалены.');
  },


  // =====================
  // ЗАПОЛНЕНИЕ ДАННЫХ
  // =====================

  fillPeriod_(dateFrom, dateTo) {
    const sheet = this.getSheet_();
    const rowsByDate = this.loadDirectDailyRows_(dateFrom, dateTo);
    const dates = this.getDateRange_(dateFrom, dateTo);

    dates.forEach(apiDate => {
      this.ensureMonthBlock_(sheet, apiDate);

      const headerRow = this.findHeaderRowForMonth_(sheet, apiDate);
      const headerMap = this.getHeaderMap_(sheet, headerRow);
      const sheetRow = this.findDateRow_(sheet, headerMap.date, headerRow, apiDate);

      if (!sheetRow) {
        Logger.log('Дата не найдена в ДБ: ' + apiDate);
        return;
      }

      const item = rowsByDate[apiDate] || this.emptyRow_();

      this.setNumberCell_(sheet, sheetRow, headerMap.impressions, item.impressions);
      this.setNumberCell_(sheet, sheetRow, headerMap.clicks, item.clicks);
      this.setNumberCell_(sheet, sheetRow, headerMap.cost, item.cost);

      this.setGoalCell_(sheet, sheetRow, headerMap.view3Pages, item.view3Pages);
      this.setGoalCell_(sheet, sheetRow, headerMap.siteSearch, item.siteSearch);
      this.setExternalGoalCellIfConfigured_(sheet, sheetRow, headerMap.callibriSpam, item.callibriSpam, this.config.goals.callibriSpam);
      this.setExternalGoalCellIfConfigured_(sheet, sheetRow, headerMap.callibriNonTarget, item.callibriNonTarget, this.config.goals.callibriNonTarget);
      this.setGoalCell_(sheet, sheetRow, headerMap.addToCart, item.addToCart);
      this.setGoalCell_(sheet, sheetRow, headerMap.purchase, item.purchase);
      this.setGoalCell_(sheet, sheetRow, headerMap.jivo, item.jivo);
      this.setExternalGoalCellIfConfigured_(sheet, sheetRow, headerMap.callibriLeadA, item.callibriLeadA, this.config.goals.callibriLeadA);
      this.setExternalGoalCellIfConfigured_(sheet, sheetRow, headerMap.callibriLeadC, item.callibriLeadC, this.config.goals.callibriLeadC);
    });

    Logger.log('ДБ заполнен за период: ' + dateFrom + ' — ' + dateTo);
  },


  // =====================
  // СОЗДАНИЕ МЕСЯЦА
  // =====================

  ensureMonthBlock_(sheet, apiDate) {
    const monthTitle = this.monthTitle_(apiDate);

    if (this.monthExists_(sheet, monthTitle)) {
      return;
    }

    // Вставляем после предыдущего месяца.
    const prevMonthDate = this.addMonths_(apiDate, -1);
    const prevMonthTitle = this.monthTitle_(prevMonthDate);
    const prevTopRow = this.findMonthTopRow_(sheet, prevMonthTitle);
    const prevBottomRow = this.findMonthBottomRow_(sheet, prevTopRow);

    // Копируем структуру из предыдущего месяца.
    // Так не подтягиваются старые лишние колонки из апреля, например Roistat Email / Roistat звонки.
    const templateTitle = prevMonthTitle;
    const templateTopRow = prevTopRow;
    const templateBottomRow = prevBottomRow;

    const copyRows = templateBottomRow - templateTopRow + 1;
    const lastCol = sheet.getLastColumn();

    // Между месяцами оставляем 1 пустую строку.
    const rowAfterPrevBlock = prevBottomRow + 1;

    if (!this.isRowEmpty_(sheet, rowAfterPrevBlock, lastCol)) {
      sheet.insertRowsAfter(prevBottomRow, 1);
    }

    const blankRow = prevBottomRow + 1;

    if (sheet.getMaxRows() < blankRow + copyRows + 5) {
      sheet.insertRowsAfter(
        sheet.getMaxRows(),
        blankRow + copyRows + 5 - sheet.getMaxRows()
      );
    }

    sheet.insertRowsAfter(blankRow, copyRows);

    const destTopRow = blankRow + 1;

    sheet
      .getRange(templateTopRow, 1, copyRows, lastCol)
      .copyTo(
        sheet.getRange(destTopRow, 1, copyRows, lastCol),
        SpreadsheetApp.CopyPasteType.PASTE_NORMAL,
        false
      );

    this.prepareCopiedMonthBlock_(sheet, destTopRow, copyRows, apiDate, templateTitle);
  },

  prepareCopiedMonthBlock_(sheet, destTopRow, copyRows, apiDate, templateTitle) {
    const monthTitle = this.monthTitle_(apiDate);
    const templateTitleNorm = this.normalize_(templateTitle);
    const daysInMonth = this.daysInMonth_(apiDate);
    const monthStart = this.monthStartApi_(apiDate);
    const monthEnd = this.monthEndApi_(apiDate);
    const lastCol = sheet.getLastColumn();

    // Переименовываем месяц во всём скопированном блоке.
    const blockRange = sheet.getRange(destTopRow, 1, copyRows, lastCol);
    const displayValues = blockRange.getDisplayValues();

    for (let r = 0; r < displayValues.length; r++) {
      for (let c = 0; c < displayValues[r].length; c++) {
        if (this.normalize_(displayValues[r][c]) === templateTitleNorm) {
          const cell = sheet.getRange(destTopRow + r, c + 1);

          if (!cell.getFormula()) {
            cell.setValue(monthTitle);
          }
        }
      }
    }

    this.updateMonthSettings_(sheet, destTopRow, copyRows, monthStart, monthEnd, daysInMonth);

    const headerRow = this.findHeaderRowForMonth_(sheet, apiDate);
    const headerMap = this.getHeaderMap_(sheet, headerRow);

    this.clearObsoleteLegacyGoalColumns_(sheet, headerRow, headerMap.date);
    this.updateBudgetCells_(sheet, destTopRow, copyRows, headerMap);

    let totalRow = this.findTotalRow_(sheet, headerRow, headerMap.date);
    let dataStartRow = this.findFirstDateRow_(sheet, headerMap.date, headerRow, totalRow) || headerRow + 1;

    let existingDataRows = totalRow - dataStartRow;

    // Подгоняем количество строк дат под текущий месяц.
    if (existingDataRows > daysInMonth) {
      sheet.deleteRows(dataStartRow + daysInMonth, existingDataRows - daysInMonth);
    }

    if (existingDataRows < daysInMonth) {
      const rowsToAdd = daysInMonth - existingDataRows;
      totalRow = this.findTotalRow_(sheet, headerRow, headerMap.date);

      sheet.insertRowsBefore(totalRow, rowsToAdd);

      const sourceRow = Math.max(dataStartRow, totalRow - 1);

      sheet
        .getRange(sourceRow, 1, 1, lastCol)
        .copyTo(
          sheet.getRange(totalRow, 1, rowsToAdd, lastCol),
          SpreadsheetApp.CopyPasteType.PASTE_NORMAL,
          false
        );
    }

    totalRow = this.findTotalRow_(sheet, headerRow, headerMap.date);
    dataStartRow = this.findFirstDateRow_(sheet, headerMap.date, headerRow, totalRow) || headerRow + 1;

    const dataEndRow = dataStartRow + daysInMonth - 1;

    // Даты месяца.
    const startDate = this.parseApiDate_(monthStart);

    for (let i = 0; i < daysInMonth; i++) {
      const row = dataStartRow + i;
      const date = new Date(startDate);
      date.setDate(startDate.getDate() + i);

      const dateCell = sheet.getRange(row, headerMap.date);
      dateCell.clearContent();
      dateCell.setValue(date);
      dateCell.setNumberFormat('dd.MM.yy');
    }

    this.clearInputCells_(sheet, headerMap, dataStartRow, daysInMonth);
    this.updatePlanColumn_(sheet, headerMap, destTopRow, dataStartRow, daysInMonth, apiDate);
    this.fixTotalRowFormulaRanges_(sheet, totalRow, dataStartRow, dataEndRow);
    this.applyFormattingFromTemplate_(sheet, headerMap, dataStartRow, daysInMonth, apiDate);
    this.updatePlanColumn_(sheet, headerMap, destTopRow, dataStartRow, daysInMonth, apiDate);
  },

  fixExistingMonthBlock_(sheet, apiDate) {
    const monthTitle = this.monthTitle_(apiDate);
    const monthTopRow = this.findMonthTopRow_(sheet, monthTitle);
    const headerRow = this.findHeaderRowForMonth_(sheet, apiDate);
    const headerMap = this.getHeaderMap_(sheet, headerRow);

    this.clearObsoleteLegacyGoalColumns_(sheet, headerRow, headerMap.date);
    this.updateBudgetCells_(sheet, monthTopRow, 120, headerMap);

    const totalRow = this.findTotalRow_(sheet, headerRow, headerMap.date);
    const dataStartRow = this.findFirstDateRow_(sheet, headerMap.date, headerRow, totalRow) || headerRow + 1;
    const daysInMonth = this.daysInMonth_(apiDate);
    const dataEndRow = dataStartRow + daysInMonth - 1;

    this.updatePlanColumn_(sheet, headerMap, monthTopRow, dataStartRow, daysInMonth, apiDate);
    this.fixTotalRowFormulaRanges_(sheet, totalRow, dataStartRow, dataEndRow);
    this.applyFormattingFromTemplate_(sheet, headerMap, dataStartRow, daysInMonth, apiDate);
    this.updatePlanColumn_(sheet, headerMap, monthTopRow, dataStartRow, daysInMonth, apiDate);
  },



  clearObsoleteLegacyGoalColumns_(sheet, headerRow, dateCol) {
    const lastCol = sheet.getLastColumn();
    const headers = sheet.getRange(headerRow, 1, 1, lastCol).getDisplayValues()[0];

    const colsToClear = [];

    headers.forEach((header, index) => {
      const key = this.normalize_(header);

      if (key === 'roistat email' || key === 'roistat звонки') {
        colsToClear.push(index + 1);
      }
    });

    if (!colsToClear.length) return;

    const totalRow = this.findTotalRow_(sheet, headerRow, dateCol);
    const numRows = totalRow - headerRow + 1;

    colsToClear.forEach(col => {
      // Не удаляем столбцы физически, чтобы не сдвинуть весь лист.
      // Очищаем только этот месячный блок.
      sheet.getRange(headerRow, col, numRows, 1)
        .clearContent()
        .setBackground('#ffffff')
        .setBorder(false, false, false, false, false, false);
    });
  },

  // =====================
  // НАСТРОЙКИ МЕСЯЦА / БЮДЖЕТ
  // =====================

  updateMonthSettings_(sheet, destTopRow, copyRows, monthStart, monthEnd, daysInMonth) {
    const maxCol = Math.min(sheet.getLastColumn(), 8);

    for (let row = destTopRow; row < destTopRow + copyRows; row++) {
      const values = sheet.getRange(row, 1, 1, maxCol).getDisplayValues()[0];
      const normalized = values.map(value => this.normalize_(value));

      const labelIndex = normalized.findIndex(value => {
        return value === 'период рк' ||
          value === 'дней всего' ||
          value === 'дней прошло' ||
          value === 'остаток' ||
          value === 'дата отчета';
      });

      if (labelIndex === -1) continue;

      const label = normalized[labelIndex];
      const labelCol = labelIndex + 1;

      if (label === 'период рк') {
        this.setIfNoFormula_(sheet, row, labelCol + 1, this.parseApiDate_(monthStart), 'dd.MM.yy');
        this.setIfNoFormula_(sheet, row, labelCol + 2, this.parseApiDate_(monthEnd), 'dd.MM.yy');
      }

      if (label === 'дней всего') {
        this.setIfNoFormula_(sheet, row, labelCol + 1, daysInMonth, null);
      }

      if (label === 'дней прошло') {
        this.setIfNoFormula_(sheet, row, labelCol + 1, this.daysPassedForMonth_(monthStart), null);
      }

      if (label === 'остаток') {
        const passed = this.daysPassedForMonth_(monthStart);
        this.setIfNoFormula_(sheet, row, labelCol + 1, Math.max(daysInMonth - passed, 0), null);
      }

      if (label === 'дата отчета') {
        this.setIfNoFormula_(sheet, row, labelCol + 1, new Date(), 'dd.MM.yy');
      }
    }
  },

  updateBudgetCells_(sheet, destTopRow, copyRows, headerMap) {
    const maxCol = Math.min(sheet.getLastColumn(), 12);
    const monthPlan = this.config.monthPlanWithVat;
    const plannedCostCol = headerMap && headerMap.plannedCost ? headerMap.plannedCost : 0;

    for (let row = destTopRow; row < destTopRow + copyRows; row++) {
      const values = sheet.getRange(row, 1, 1, maxCol).getDisplayValues()[0];
      const normalized = values.map(value => this.normalize_(value));

      for (let i = 0; i < normalized.length; i++) {
        const text = normalized[i];

        // Верхняя строка "Бюджет с НДС": ставим бюджет строго справа от подписи.
        if (text.indexOf('бюджет') !== -1 && text.indexOf('ндс') !== -1) {
          const targetCol = i + 2;
          this.setBudgetCell_(sheet, row, targetCol, monthPlan);
        }

        // Строка "ПЛАН по МП": бюджет ставим в колонку "Расход План c НДС".
        if (text === 'план по мп') {
          const targetCol = plannedCostCol || 5;
          this.setBudgetCell_(sheet, row, targetCol, monthPlan);
        }
      }
    }
  },

  setBudgetCell_(sheet, row, col, value) {
    const cell = sheet.getRange(row, col);
    cell.clearContent();
    cell.setValue(value);
    cell.setNumberFormat('"р." #,##0.00');
  },

  fixMonthTitleBeforeTable_(sheet, headerRow, dateCol, monthTitle) {
    const titleRow = headerRow - 3;
    const col = dateCol && dateCol > 0 ? dateCol : 2;

    if (titleRow <= 0) return;

    const cell = sheet.getRange(titleRow, col);

    if (!cell.getFormula()) {
      cell.setValue(monthTitle);
    }
  },

  updatePlanColumn_(sheet, headerMap, monthTopRow, dataStartRow, daysInMonth, apiDate) {
    let col = headerMap.plannedCost;

    if (!col || col <= 0) {
      col = this.findPlannedCostColumnFallback_(sheet, dataStartRow);
    }

    if (!col || col <= 0) {
      Logger.log('Колонка "Расход План c НДС" не найдена.');
      return;
    }

    const formulaParts = this.findPlanFormulaParts_(sheet, monthTopRow, col);

    if (!formulaParts) {
      Logger.log('Не нашла ячейку "ПЛАН по МП" для формулы Расход План c НДС.');
      return;
    }

    const workingDays = this.countWorkingDays_(apiDate);
    const monthStart = this.parseApiDate_(this.monthStartApi_(apiDate));

    for (let i = 0; i < daysInMonth; i++) {
      const row = dataStartRow + i;
      const date = new Date(monthStart);
      date.setDate(monthStart.getDate() + i);

      const api = this.formatDate_(date, 'yyyy-MM-dd');
      const cell = sheet.getRange(row, col);

      cell.clearContent();

      if (workingDays > 0 && this.isWorkingDay_(date, api)) {
        cell.setFormula(
          '=' +
          '$' + this.columnToLetter_(formulaParts.planCol) + '$' + formulaParts.planRow +
          '/' +
          workingDays
        );
        cell.setNumberFormat('"р." #,##0.00');
      }
    }
  },


  // =====================
  // ПОИСКИ ПО ЛИСТУ
  // =====================

  getSheet_() {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(this.config.sheetName);

    if (!sheet) {
      throw new Error('Не найден лист: ' + this.config.sheetName);
    }

    return sheet;
  },

  monthExists_(sheet, monthTitle) {
    const lastRow = sheet.getLastRow();
    const maxCol = Math.min(sheet.getLastColumn(), 10);
    const target = this.normalize_(monthTitle);
    const values = sheet.getRange(1, 1, lastRow, maxCol).getDisplayValues();

    for (let r = 0; r < values.length; r++) {
      for (let c = 0; c < values[r].length; c++) {
        if (this.normalize_(values[r][c]) === target) {
          return true;
        }
      }
    }

    return false;
  },

  findMonthTopRow_(sheet, monthTitle) {
    const lastRow = sheet.getLastRow();
    const maxCol = Math.min(sheet.getLastColumn(), 10);
    const target = this.normalize_(monthTitle);
    const values = sheet.getRange(1, 1, lastRow, maxCol).getDisplayValues();

    for (let r = 0; r < values.length; r++) {
      for (let c = 0; c < values[r].length; c++) {
        if (this.normalize_(values[r][c]) === target) {
          return r + 1;
        }
      }
    }

    throw new Error('Не нашла блок месяца: ' + monthTitle);
  },

  findMonthBottomRow_(sheet, monthTopRow) {
    const maxRow = Math.min(sheet.getLastRow(), monthTopRow + 160);
    const maxCol = Math.min(sheet.getLastColumn(), 12);

    for (let row = monthTopRow; row <= maxRow; row++) {
      const values = sheet.getRange(row, 1, 1, maxCol).getDisplayValues()[0];

      if (values.some(value => this.normalize_(value) === 'без ндс')) {
        return row;
      }
    }

    throw new Error('Не нашла нижнюю строку месяца "Без НДС" после строки: ' + monthTopRow);
  },

  findHeaderRowForMonth_(sheet, apiDate) {
    const monthTitle = this.monthTitle_(apiDate);
    const monthTopRow = this.findMonthTopRow_(sheet, monthTitle);
    const maxRow = Math.min(sheet.getLastRow(), monthTopRow + 120);
    const maxCol = sheet.getLastColumn();

    for (let row = monthTopRow; row <= maxRow; row++) {
      const values = sheet.getRange(row, 1, 1, maxCol).getDisplayValues()[0];
      const normalized = values.map(value => this.normalize_(value));

      const hasDate = normalized.some(value => value === 'дата');
      const hasImpressions = normalized.some(value => value === 'показы');
      const hasClicks = normalized.some(value => value === 'клики');

      if (hasDate && hasImpressions && hasClicks) {
        return row;
      }
    }

    throw new Error('Не нашла шапку месяца: ' + monthTitle);
  },

  getHeaderMap_(sheet, headerRow) {
    const headers = sheet.getRange(headerRow, 1, 1, sheet.getLastColumn()).getDisplayValues()[0];
    const map = {};

    headers.forEach((header, index) => {
      const key = this.normalize_(header);

      if (key) {
        map[key] = index + 1;
      }
    });

    return {
      date: this.findExactCol_(map, 'дата'),

      impressions: this.findExactCol_(map, 'показы'),
      clicks: this.findExactCol_(map, 'клики'),
      plannedCost: this.findContainsCol_(map, ['расход план с ндс', 'расход план c ндс']),
      cost: this.findContainsCol_(map, ['расход факт с ндс', 'расход факт c ндс']),

      view3Pages: this.findContainsCol_(map, ['просмотр 3х страниц']),
      siteSearch: this.findContainsCol_(map, ['поиск по сайту']),

      callibriSpam: this.findContainsCol_(map, ['callibri: спам', 'callibri спам']),
      callibriNonTarget: this.findContainsCol_(map, ['callibri: нецелевой_лид', 'callibri нецелевой лид', 'callibri нецелевой_лид']),
      addToCart: this.findContainsCol_(map, ['ecommerce: добавление в корзину', 'ecommerce добавление в корзину']),
      purchase: this.findContainsCol_(map, ['ecommerce: покупка', 'ecommerce покупка']),
      jivo: this.findContainsCol_(map, ['jivo-чат', 'jivo чат', 'jivo']),
      callibriLeadA: this.findContainsCol_(map, ['callibri: лид_квал_a', 'callibri лид квал a', 'callibri лид_квал_a']),
      callibriLeadB: this.findContainsCol_(map, ['callibri: лид_квал_b', 'callibri лид квал b', 'callibri лид_квал_b']),
      callibriLeadC: this.findContainsCol_(map, ['callibri: лид_квал_c', 'callibri лид квал c', 'callibri лид_квал_c'])
    };
  },

  findExactCol_(map, header) {
    return map[this.normalize_(header)] || 0;
  },

  findContainsCol_(map, parts) {
    const normalizedParts = parts.map(part => this.normalize_(part));

    for (const key in map) {
      for (let i = 0; i < normalizedParts.length; i++) {
        if (key.indexOf(normalizedParts[i]) !== -1) {
          return map[key];
        }
      }
    }

    return 0;
  },

  findTotalRow_(sheet, headerRow, dateCol) {
    const maxRow = Math.min(sheet.getLastRow(), headerRow + 160);
    const lastCol = sheet.getLastColumn();

    // В ДБ строка ИТОГО идёт прямо перед "ПЛАН по МП".
    for (let row = headerRow + 1; row <= maxRow; row++) {
      const values = sheet.getRange(row, 1, 1, Math.min(lastCol, 12)).getDisplayValues()[0];

      if (values.some(value => this.normalize_(value) === 'план по мп')) {
        const totalRow = row - 1;

        if (dateCol > 0) {
          const totalCell = sheet.getRange(totalRow, dateCol);

          if (!totalCell.getFormula() && this.normalize_(totalCell.getDisplayValue()) !== 'итого') {
            totalCell.setValue('ИТОГО');
          }
        }

        return totalRow;
      }
    }

    throw new Error('Не нашла строку ИТОГО: нет строки "ПЛАН по МП" после строки ' + headerRow);
  },

  findFirstDateRow_(sheet, dateCol, headerRow, totalRow) {
    for (let row = headerRow + 1; row < totalRow; row++) {
      const raw = sheet.getRange(row, dateCol).getValue();
      const display = String(sheet.getRange(row, dateCol).getDisplayValue() || '').trim();

      if (raw instanceof Date) return row;
      if (/^\d{1,2}\.\d{1,2}(\.\d{2,4})?$/.test(display)) return row;
    }

    return 0;
  },

  findDateRow_(sheet, dateCol, headerRow, apiDate) {
    if (!dateCol || dateCol <= 0) {
      throw new Error('Не найдена колонка "Дата" в ДБ.');
    }

    const totalRow = this.findTotalRow_(sheet, headerRow, dateCol);
    const rowsCount = totalRow - headerRow - 1;

    if (rowsCount <= 0) return 0;

    const range = sheet.getRange(headerRow + 1, dateCol, rowsCount, 1);
    const values = range.getValues();
    const displayValues = range.getDisplayValues();

    const targetLong = this.formatApiDate_(apiDate, 'dd.MM.yyyy');
    const targetLongNoZero = this.formatApiDate_(apiDate, 'dd.M.yyyy');
    const targetShort = this.formatApiDate_(apiDate, 'dd.MM');
    const targetShortNoZero = this.formatApiDate_(apiDate, 'dd.M');

    for (let i = 0; i < values.length; i++) {
      const raw = values[i][0];
      const display = String(displayValues[i][0] || '').trim();

      if (raw instanceof Date) {
        const rawApiDate = this.formatDate_(raw, 'yyyy-MM-dd');

        if (rawApiDate === apiDate) {
          return headerRow + 1 + i;
        }
      }

      if (
        display === targetLong ||
        display === targetLongNoZero ||
        display === targetShort ||
        display === targetShort + '.' ||
        display === targetShortNoZero ||
        display === targetShortNoZero + '.'
      ) {
        return headerRow + 1 + i;
      }
    }

    return 0;
  },

  findPlannedCostColumnFallback_(sheet, dataStartRow) {
    const headerRow = dataStartRow - 1;
    const lastCol = sheet.getLastColumn();
    const headers = sheet.getRange(headerRow, 1, 1, lastCol).getDisplayValues()[0];

    for (let i = 0; i < headers.length; i++) {
      const key = this.normalize_(headers[i]);

      if (
        key.indexOf('расход') !== -1 &&
        key.indexOf('план') !== -1 &&
        key.indexOf('ндс') !== -1
      ) {
        return i + 1;
      }
    }

    return 0;
  },

  findPlanFormulaParts_(sheet, monthTopRow, plannedCostCol) {
    const maxRow = Math.min(sheet.getLastRow(), monthTopRow + 160);

    for (let row = monthTopRow; row <= maxRow; row++) {
      const values = sheet.getRange(row, 1, 1, Math.min(sheet.getLastColumn(), 12)).getDisplayValues()[0];

      for (let c = 0; c < values.length; c++) {
        const text = this.normalize_(values[c]);

        if (text === 'план по мп') {
          return {
            planRow: row,
            planCol: plannedCostCol
          };
        }
      }
    }

    return null;
  },


  // =====================
  // ОФОРМЛЕНИЕ / ФОРМУЛЫ / ПРОВЕРКА
  // =====================

  clearInputCells_(sheet, headerMap, dataStartRow, daysInMonth) {
    const columns = [
      headerMap.impressions,
      headerMap.clicks,
      headerMap.cost,
      headerMap.view3Pages,
      // SAFETY_2026_07_07: Callibri пока внешний источник, не очищаем его при создании/фиксе месяца.
      headerMap.addToCart,
      headerMap.purchase,
      headerMap.jivo
    ].filter(col => col && col > 0);

    columns.forEach(col => {
      for (let i = 0; i < daysInMonth; i++) {
        const cell = sheet.getRange(dataStartRow + i, col);

        if (!cell.getFormula()) {
          cell.clearContent();
        }
      }
    });
  },

  fixTotalRowFormulaRanges_(sheet, totalRow, dataStartRow, dataEndRow) {
    const lastCol = sheet.getLastColumn();
    const range = sheet.getRange(totalRow, 1, 1, lastCol);
    const formulas = range.getFormulas()[0];

    const fixed = formulas.map(formula => {
      if (!formula) return formula;
      return this.fixFormulaRangesToCurrentBlock_(formula, dataStartRow, dataEndRow);
    });

    range.setFormulas([fixed]);
  },

  fixFormulaRangesToCurrentBlock_(formula, dataStartRow, dataEndRow) {
    return String(formula).replace(
      /(\$?[A-Z]{1,3})(\$?\d+):(\$?[A-Z]{1,3})(\$?\d+)/g,
      function(match, col1, row1, col2, row2) {
        const cleanCol1 = col1.replace(/\$/g, '');
        const cleanCol2 = col2.replace(/\$/g, '');

        if (cleanCol1 !== cleanCol2) return match;

        const newRow1 = row1.indexOf('$') === 0 ? '$' + dataStartRow : String(dataStartRow);
        const newRow2 = row2.indexOf('$') === 0 ? '$' + dataEndRow : String(dataEndRow);

        return col1 + newRow1 + ':' + col2 + newRow2;
      }
    );
  },

  applyFormattingFromTemplate_(sheet, headerMap, dataStartRow, daysInMonth, apiDate) {
    const templateDate = this.config.templateMonthDate;
    const templateHeaderRow = this.findHeaderRowForMonth_(sheet, templateDate);
    const templateHeaderMap = this.getHeaderMap_(sheet, templateHeaderRow);
    const templateTotalRow = this.findTotalRow_(sheet, templateHeaderRow, templateHeaderMap.date);
    const templateDataStartRow = this.findFirstDateRow_(sheet, templateHeaderMap.date, templateHeaderRow, templateTotalRow) || templateHeaderRow + 1;

    const samples = this.findTemplateSampleRows_(sheet, templateHeaderMap.date, templateDataStartRow, templateTotalRow);

    const startCol = headerMap.date && headerMap.date > 0 ? headerMap.date : 1;
    const endCol = this.findTableEndCol_(sheet, dataStartRow - 1);
    const width = Math.max(endCol - startCol + 1, 1);

    const monthStart = this.parseApiDate_(this.monthStartApi_(apiDate));

    for (let i = 0; i < daysInMonth; i++) {
      const row = dataStartRow + i;
      const date = new Date(monthStart);
      date.setDate(monthStart.getDate() + i);

      const api = this.formatDate_(date, 'yyyy-MM-dd');
      const sourceRow = this.isWorkingDay_(date, api) ? samples.weekdayRow : samples.weekendRow;

      sheet
        .getRange(sourceRow, startCol, 1, width)
        .copyTo(
          sheet.getRange(row, startCol, 1, width),
          SpreadsheetApp.CopyPasteType.PASTE_FORMAT,
          false
        );
    }
  },

  findTemplateSampleRows_(sheet, dateCol, dataStartRow, totalRow) {
    let weekdayRow = 0;
    let weekendRow = 0;

    for (let row = dataStartRow; row < totalRow; row++) {
      const raw = sheet.getRange(row, dateCol).getValue();
      const display = sheet.getRange(row, dateCol).getDisplayValue();
      const date = this.parseDbDateValue_(raw, display);

      if (!date) continue;

      const api = this.formatDate_(date, 'yyyy-MM-dd');
      const isWorking = this.isWorkingDay_(date, api);

      if (isWorking && !weekdayRow) weekdayRow = row;
      if (!isWorking && !weekendRow) weekendRow = row;

      if (weekdayRow && weekendRow) break;
    }

    if (!weekdayRow) weekdayRow = dataStartRow;
    if (!weekendRow) weekendRow = weekdayRow;

    return {
      weekdayRow: weekdayRow,
      weekendRow: weekendRow
    };
  },

  parseDbDateValue_(raw, display) {
    if (raw instanceof Date) return raw;

    const text = String(display || '').trim();
    const match = text.match(/^(\d{1,2})\.(\d{1,2})(?:\.(\d{2,4}))?\.?$/);

    if (!match) return null;

    const day = Number(match[1]);
    const month = Number(match[2]) - 1;

    let year = new Date().getFullYear();

    if (match[3]) {
      year = Number(match[3]);
      if (year < 100) year += 2000;
    }

    return new Date(year, month, day);
  },

  findTableEndCol_(sheet, headerRow) {
    const lastCol = sheet.getLastColumn();
    const headers = sheet.getRange(headerRow, 1, 1, lastCol).getDisplayValues()[0];

    for (let i = headers.length - 1; i >= 0; i--) {
      if (String(headers[i] || '').trim() !== '') {
        return i + 1;
      }
    }

    return lastCol;
  },

  validateMonth_(sheet, apiDate) {
    const monthTitle = this.monthTitle_(apiDate);
    const headerRow = this.findHeaderRowForMonth_(sheet, apiDate);
    const headerMap = this.getHeaderMap_(sheet, headerRow);
    const totalRow = this.findTotalRow_(sheet, headerRow, headerMap.date);
    const dataStartRow = this.findFirstDateRow_(sheet, headerMap.date, headerRow, totalRow) || headerRow + 1;
    const daysInMonth = this.daysInMonth_(apiDate);
    const expectedTotalRow = dataStartRow + daysInMonth;

    if (totalRow !== expectedTotalRow) {
      throw new Error(
        'Проверка ' + monthTitle + ': строка ИТОГО найдена на ' + totalRow +
        ', а должна быть на ' + expectedTotalRow +
        '. Первая дата: ' + dataStartRow + ', дней в месяце: ' + daysInMonth
      );
    }

    const planCol = headerMap.plannedCost || this.findPlannedCostColumnFallback_(sheet, dataStartRow);
    const monthStart = this.parseApiDate_(this.monthStartApi_(apiDate));

    for (let i = 0; i < daysInMonth; i++) {
      const row = dataStartRow + i;
      const date = new Date(monthStart);
      date.setDate(monthStart.getDate() + i);

      const api = this.formatDate_(date, 'yyyy-MM-dd');
      const cell = sheet.getRange(row, planCol);

      if (this.isWorkingDay_(date, api) && !cell.getFormula()) {
        throw new Error('Проверка ' + monthTitle + ': в плановой колонке нет формулы в рабочий день, строка ' + row);
      }

      if (!this.isWorkingDay_(date, api) && cell.getDisplayValue()) {
        throw new Error('Проверка ' + monthTitle + ': в выходной/праздник план должен быть пустым, строка ' + row);
      }
    }

    Logger.log('Проверка месяца пройдена: ' + monthTitle);
  },


  // =====================
  // ЗАПИСЬ
  // =====================

  setNumberCell_(sheet, row, col, value) {
    if (!col || col <= 0) return;

    const cell = sheet.getRange(row, col);

    if (cell.getFormula()) return;

    cell.setValue(Number(value || 0));
  },

  setGoalCell_(sheet, row, col, value) {
    if (!col || col <= 0) return;

    const cell = sheet.getRange(row, col);

    if (cell.getFormula()) return;

    const number = Number(value || 0);
    cell.setValue(number > 0 ? number : '-');
  },

  setExternalGoalCellIfConfigured_(sheet, row, col, value, sourceIds) {
    // SAFETY_2026_07_07:
    // Если внешний источник Callibri ещё не подключён к этому генератору,
    // не затираем уже записанные значения в ДБ нулями / прочерками.
    if (!sourceIds || !sourceIds.length) return;
    this.setGoalCell_(sheet, row, col, value);
  },

  setIfNoFormula_(sheet, row, col, value, numberFormat) {
    const cell = sheet.getRange(row, col);

    if (cell.getFormula()) return;

    cell.setValue(value);

    if (numberFormat) {
      cell.setNumberFormat(numberFormat);
    }
  },


  // =====================
  // ДИРЕКТ
  // =====================

  loadDirectDailyRows_(dateFrom, dateTo) {
    const allGoalIds = this.getAllGoalIds_();

    const reportBody = {
      params: {
        SelectionCriteria: {
          DateFrom: dateFrom,
          DateTo: dateTo
        },
        Goals: allGoalIds.map(id => Number(id)),
        AttributionModels: ['AUTO'],
        FieldNames: [
          'Date',
          'Impressions',
          'Clicks',
          'Cost',
          'Conversions'
        ],
        ReportName: 'tdm_db_' + dateFrom + '_' + dateTo + '_' + Utilities.getUuid(),
        ReportType: 'ACCOUNT_PERFORMANCE_REPORT',
        DateRangeType: 'CUSTOM_DATE',
        Format: 'TSV',
        IncludeVAT: 'YES',
        IncludeDiscount: 'NO'
      }
    };

    const text = this.requestDirectReport_(reportBody);
    return this.parseDirectDailyTsv_(text);
  },

  requestDirectReport_(reportBody) {
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
      payload: JSON.stringify(reportBody),
      headers: {
        Authorization: 'Bearer ' + token,
        'Client-Login': this.config.clientLogin,
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

      if (code === 200) return text;

      if (code === 201 || code === 202) {
        Utilities.sleep(5000);
        continue;
      }

      throw new Error('Ошибка Директа. Код: ' + code + '\n' + text);
    }

    throw new Error('Отчёт Директа не успел сформироваться.');
  },

  parseDirectDailyTsv_(tsvText) {
    const text = String(tsvText || '').trim();

    if (!text) return {};

    const lines = text.split(/\r?\n/).filter(line => String(line).trim() !== '');

    const headerIndex = lines.findIndex(line => {
      const cols = line.split('\t');
      return cols.indexOf('Date') !== -1 && cols.indexOf('Cost') !== -1;
    });

    if (headerIndex === -1) {
      throw new Error('Не найдена строка заголовков в отчёте Директа.');
    }

    const headers = lines[headerIndex]
      .split('\t')
      .map(header => String(header).replace(/^\uFEFF/, '').trim());

    const idxDate = this.requireColumn_(headers, 'Date');
    const idxImpressions = this.requireColumn_(headers, 'Impressions');
    const idxClicks = this.requireColumn_(headers, 'Clicks');
    const idxCost = this.requireColumn_(headers, 'Cost');

    const goalIndexes = {};

    this.getAllGoalIds_().forEach(goalId => {
      goalIndexes[goalId] = this.findGoalColumn_(headers, goalId);

      if (goalIndexes[goalId] === -1) {
        Logger.log('Не найден столбец цели в отчёте Директа: ' + goalId);
      }
    });

    const result = {};

    for (let i = headerIndex + 1; i < lines.length; i++) {
      const row = lines[i].split('\t');
      const apiDate = String(row[idxDate] || '').trim();

      if (!apiDate) continue;

      if (!result[apiDate]) {
        result[apiDate] = this.emptyRow_();
      }

      result[apiDate].impressions += this.toNumber_(row[idxImpressions]);
      result[apiDate].clicks += this.toNumber_(row[idxClicks]);
      result[apiDate].cost += this.toNumber_(row[idxCost]);

      result[apiDate].view3Pages += this.getGoalSum_(row, goalIndexes, this.config.goals.view3Pages);
      result[apiDate].siteSearch += this.getGoalSum_(row, goalIndexes, this.config.goals.siteSearch);
      result[apiDate].callibriSpam += this.getGoalSum_(row, goalIndexes, this.config.goals.callibriSpam);
      result[apiDate].callibriNonTarget += this.getGoalSum_(row, goalIndexes, this.config.goals.callibriNonTarget);
      result[apiDate].addToCart += this.getGoalSum_(row, goalIndexes, this.config.goals.addToCart);
      result[apiDate].purchase += this.getGoalSum_(row, goalIndexes, this.config.goals.purchase);
      result[apiDate].jivo += this.getGoalSum_(row, goalIndexes, this.config.goals.jivo);
      result[apiDate].callibriLeadA += this.getGoalSum_(row, goalIndexes, this.config.goals.callibriLeadA);
      result[apiDate].callibriLeadC += this.getGoalSum_(row, goalIndexes, this.config.goals.callibriLeadC);
    }

    return result;
  },

  getGoalSum_(row, goalIndexes, goalIds) {
    return goalIds.reduce((sum, goalId) => {
      const idx = goalIndexes[goalId];

      if (idx === -1 || idx === undefined) return sum;

      return sum + this.toNumber_(row[idx]);
    }, 0);
  },

  getAllGoalIds_() {
    let result = [];

    Object.keys(this.config.goals).forEach(key => {
      result = result.concat(this.config.goals[key]);
    });

    return result;
  },

  findGoalColumn_(headers, goalId) {
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
  },


  // =====================
  // ОБЩЕЕ
  // =====================

  emptyRow_() {
    return {
      impressions: 0,
      clicks: 0,
      cost: 0,
      view3Pages: 0,
      siteSearch: 0,
      callibriSpam: 0,
      callibriNonTarget: 0,
      addToCart: 0,
      purchase: 0,
      jivo: 0,
      callibriLeadA: 0,
      callibriLeadC: 0
    };
  },

  requireColumn_(headers, columnName) {
    const index = headers.indexOf(columnName);

    if (index === -1) {
      throw new Error('Не найден столбец "' + columnName + '". Заголовки: ' + headers.join(' | '));
    }

    return index;
  },

  getYesterday_() {
    const now = new Date();
    const yesterday = new Date(now);
    yesterday.setDate(now.getDate() - 1);

    return this.formatDate_(yesterday, 'yyyy-MM-dd');
  },

  getDateRange_(dateFrom, dateTo) {
    const start = this.parseApiDate_(dateFrom);
    const end = this.parseApiDate_(dateTo);
    const dates = [];

    const current = new Date(start);

    while (current <= end) {
      dates.push(this.formatDate_(current, 'yyyy-MM-dd'));
      current.setDate(current.getDate() + 1);
    }

    return dates;
  },

  parseApiDate_(apiDate) {
    const parts = String(apiDate).split('-').map(Number);
    return new Date(parts[0], parts[1] - 1, parts[2]);
  },

  formatApiDate_(apiDate, format) {
    const date = this.parseApiDate_(apiDate);
    return this.formatDate_(date, format);
  },

  formatDate_(date, format) {
    return Utilities.formatDate(date, this.config.timezone, format);
  },

  monthTitle_(apiDate) {
    const date = this.parseApiDate_(apiDate);
    const monthNames = [
      'Январь', 'Февраль', 'Март', 'Апрель', 'Май', 'Июнь',
      'Июль', 'Август', 'Сентябрь', 'Октябрь', 'Ноябрь', 'Декабрь'
    ];

    return monthNames[date.getMonth()] + ' ' + date.getFullYear();
  },

  monthStartApi_(apiDate) {
    const date = this.parseApiDate_(apiDate);

    return this.formatDate_(
      new Date(date.getFullYear(), date.getMonth(), 1),
      'yyyy-MM-dd'
    );
  },

  monthEndApi_(apiDate) {
    const date = this.parseApiDate_(apiDate);

    return this.formatDate_(
      new Date(date.getFullYear(), date.getMonth() + 1, 0),
      'yyyy-MM-dd'
    );
  },

  addMonths_(apiDate, diff) {
    const date = this.parseApiDate_(apiDate);
    date.setMonth(date.getMonth() + diff);

    return this.formatDate_(date, 'yyyy-MM-dd');
  },

  daysInMonth_(apiDate) {
    const date = this.parseApiDate_(apiDate);
    return new Date(date.getFullYear(), date.getMonth() + 1, 0).getDate();
  },

  daysPassedForMonth_(monthStartApi) {
    const monthStart = this.parseApiDate_(monthStartApi);
    const today = this.parseApiDate_(this.formatDate_(new Date(), 'yyyy-MM-dd'));

    if (today.getFullYear() < monthStart.getFullYear()) return 0;
    if (today.getFullYear() === monthStart.getFullYear() && today.getMonth() < monthStart.getMonth()) return 0;

    const daysInMonth = new Date(monthStart.getFullYear(), monthStart.getMonth() + 1, 0).getDate();

    if (today.getFullYear() === monthStart.getFullYear() && today.getMonth() === monthStart.getMonth()) {
      return Math.max(Math.min(today.getDate() - 1, daysInMonth), 0);
    }

    return daysInMonth;
  },

  countWorkingDays_(apiDate) {
    const daysInMonth = this.daysInMonth_(apiDate);
    const start = this.parseApiDate_(this.monthStartApi_(apiDate));
    let count = 0;

    for (let i = 0; i < daysInMonth; i++) {
      const date = new Date(start);
      date.setDate(start.getDate() + i);

      const api = this.formatDate_(date, 'yyyy-MM-dd');

      if (this.isWorkingDay_(date, api)) {
        count++;
      }
    }

    return count;
  },

  isWorkingDay_(date, apiDate) {
    const day = date.getDay();

    if (day === 0 || day === 6) return false;
    if (this.config.holidays.indexOf(apiDate) !== -1) return false;

    return true;
  },

  toNumber_(value) {
    if (value === null || value === undefined) return 0;

    const text = String(value)
      .replace(/\s/g, '')
      .replace(',', '.')
      .replace(/[^\d.-]/g, '');

    const number = Number(text);

    return isNaN(number) ? 0 : number;
  },

  normalize_(value) {
    return String(value || '')
      .toLowerCase()
      .replace(/ё/g, 'е')
      .replace(/[«»"']/g, '')
      .replace(/\s+/g, ' ')
      .trim();
  },

  columnToLetter_(column) {
    let temp = '';
    let letter = '';

    while (column > 0) {
      temp = (column - 1) % 26;
      letter = String.fromCharCode(temp + 65) + letter;
      column = (column - temp - 1) / 26;
    }

    return letter;
  },

  isRowEmpty_(sheet, row, lastCol) {
    if (row > sheet.getMaxRows()) return true;

    const values = sheet.getRange(row, 1, 1, lastCol).getDisplayValues()[0];

    return values.every(value => String(value || '').trim() === '');
  },

  safeMessage_(message) {
    Logger.log(message);

    try {
      SpreadsheetApp.getUi().alert(message);
    } catch (e) {
      Logger.log('UI-сообщение не показано: ' + e.message);
    }
  }
};

// =====================
// ТДМ_2026 — единый автооркестратор отчётов
// Создано: 2026-07-07
// Не пишет в формулы и ИТОГО напрямую: вызывает существующие функции и логирует статус в «Рабочие».
// =====================

function tdmAutoDailyReports20260707() {
  // Старый оркестратор повторно запускал ДБ и несколько устаревших отчётов.
  // Оставлен как безопасная заглушка для старых ручных ссылок.
  return {
    ok: true,
    mode: 'LEGACY_DAILY_DISABLED',
    canonicalDbHandler: 'fillTdmDbYesterday',
    canonicalMondayHandler: 'tdmRunVerifiedMondayReports20260713'
  };
}

function archived_tdmBackfillDbSearchSiteJuly0607_20260708() {
  TDM_DB.fillPeriod_('2026-07-06', '2026-07-07');
  return { ok: true, period: '2026-07-06 — 2026-07-07', fields: ['Поиск по сайту'] };
}

function archived_tdmFixCallibriManualJuly020607_20260708() {
  const ss = SpreadsheetApp.openById('1zXDDfiRYHJE34iOJg-QfAC0bvoTKMIAuVMfsJS0z5Dg');
  const db = ss.getSheetByName('ДБ');
  const check = ss.getSheetByName('Callibri_Сверка');

  if (!db) throw new Error('Не найден лист ДБ');
  if (!check) throw new Error('Не найден лист Callibri_Сверка');

  // ДБ: 02.07 возвращаем Callibri A/C, 06.07 вносим Spam из Callibri с utm_source=ya.
  db.getRange('R291').setValue(1);
  db.getRange('S291').setValue(2);
  db.getRange('M295').setValue(1);
  db.getRange('M296').setValue('-');
  db.getRange('N296').setValue('-');
  db.getRange('R296').setValue('-');
  db.getRange('S296').setValue('-');

  // Callibri_Сверка: фиксируем реальные статусы из Callibri, включая utm_source=ya.
  check.getRange('A5:H5').setValues([['02.07.2026', 3, 1, 2, 0, 0, 'Записано', 'Клиент проставил статусы в Callibri']]);
  check.getRange('A9:H9').setValues([['07.07.2026', 1, 0, 0, 0, 0, 'Записано', 'Есть yandex/ya cpc без явного класса — не считаем лидом']]);
  check.getRange('A10:H10').setValues([['06.07.2026', 1, 0, 0, 1, 0, 'Записано', 'utm_source=ya, utm_medium=cpc; статус Спам']]);

  return { ok: true, fixed: ['02.07 A=1 C=2', '06.07 Spam=1', '07.07 без класса, не лид'] };
}

function tdmUpdateRegionReportCurrentMonth20260707() {
  TDM_REGION_AUTO_20260707.updateCurrentMonthToYesterday();
}

function tdmCheckCallibriPublic20260707() {
  const today = new Date();
  const yesterday = new Date(today);
  yesterday.setDate(today.getDate() - 1);
  const dateFrom = Utilities.formatDate(yesterday, 'Europe/Moscow', 'yyyy-MM-dd');
  const dateTo = dateFrom;
  return tdmGptCheckCallibriAggregate_({
    dateFrom: dateFrom,
    dateTo: dateTo,
    mode: 'yesterday_daily_sync'
  });
}

function tdmDailyCallibriSync20260708() {
  const period = tdmCallibriRolling7DayPeriod_();
  const snapshot = tdmGptCallibriAggregateYandexCpcByCampaign_(period.dateFrom, period.dateTo);
  const summary = tdmGptCheckCallibriAggregate_({
    dateFrom: period.dateFrom,
    dateTo: period.dateTo,
    mode: 'rolling_7_days_daily_sync',
    callibriSnapshot: snapshot
  });
  const writeResult = tdmApplyCallibriAggregateToDbPeriod_(snapshot, period.dateFrom, period.dateTo);

  return {
    ok: true,
    mode: 'daily_callibri_rolling_7_days',
    period: period,
    aggregate: summary,
    write: writeResult
  };
}

function tdmCheckCallibriLast3Workdays20260708() {
  // Совместимый read-only вход. Старое имя оставлено для существующих вызовов,
  // но фактически проверяется то же скользящее окно 7 дней.
  return tdmCheckCallibriRolling7Days20260713();
}

function tdmCheckCallibriRolling7Days20260713() {
  const period = tdmCallibriRolling7DayPeriod_();
  return tdmGptCheckCallibriAggregate_({
    dateFrom: period.dateFrom,
    dateTo: period.dateTo,
    mode: 'rolling_7_days_check'
  });
}

function archived_tdmCheckCallibriLast3WorkdaysBeforeRolling7_20260708() {
  const timezone = 'Europe/Moscow';
  const today = new Date();
  const days = [];
  const cursor = new Date(today);
  cursor.setDate(cursor.getDate() - 1);

  while (days.length < 3) {
    const day = cursor.getDay();
    if (day !== 0 && day !== 6) {
      days.push(Utilities.formatDate(cursor, timezone, 'yyyy-MM-dd'));
    }
    cursor.setDate(cursor.getDate() - 1);
  }

  const dateFrom = days[days.length - 1];
  const dateTo = days[0];

  return tdmGptCheckCallibriAggregate_({
    dateFrom: dateFrom,
    dateTo: dateTo,
    mode: 'last_3_workdays_recheck'
  });
}

function tdmCallibriRolling7DayPeriod_() {
  const timezone = 'Europe/Moscow';
  const dateToValue = new Date();
  dateToValue.setDate(dateToValue.getDate() - 1);
  const dateFromValue = new Date(dateToValue);
  dateFromValue.setDate(dateToValue.getDate() - 6);
  return {
    dateFrom: Utilities.formatDate(dateFromValue, timezone, 'yyyy-MM-dd'),
    dateTo: Utilities.formatDate(dateToValue, timezone, 'yyyy-MM-dd')
  };
}

function tdmCallibriApiDates_(dateFrom, dateTo) {
  const result = [];
  let cursor = tdmGptParseApiDate_(dateFrom);
  const end = tdmGptParseApiDate_(dateTo);
  while (cursor <= end) {
    result.push(tdmGptFormatApiDate_(cursor));
    cursor = new Date(cursor);
    cursor.setDate(cursor.getDate() + 1);
  }
  return result;
}

function tdmApplyCallibriAggregateToDbPeriod_(callibriSnapshot, dateFrom, dateTo) {
  if (!callibriSnapshot || typeof callibriSnapshot !== 'object') {
    throw new Error('ДБ: не передан снимок Callibri. Повторный запрос API внутри записи запрещён.');
  }

  const rows = Object.keys(callibriSnapshot).sort().map(function(key) {
    return callibriSnapshot[key];
  });
  const dates = tdmCallibriApiDates_(dateFrom, dateTo);
  const write = TDM_CALLIBRI_STABLE_20260709.writeSelectedDates_([{ rows: rows }], dates);
  return { ok: true, period: { dateFrom: dateFrom, dateTo: dateTo }, dates: dates, write: write };
}

function archived_tdmUpdateDbCallibriLast3Workdays20260708() {
  return TDM_DB_CALLIBRI_20260708.writeLast3WorkdaysFromRecheckSheet();
}

const TDM_DB_CALLIBRI_20260708 = {
  spreadsheetId: '1zXDDfiRYHJE34iOJg-QfAC0bvoTKMIAuVMfsJS0z5Dg',
  dbSheetName: 'ДБ',
  checkSheetName: 'Callibri_Сверка',
  timezone: 'Europe/Moscow',

  writeLast3WorkdaysFromRecheckSheet() {
    const dates = this.yesterdayAndLast3Workdays_();
    const ss = SpreadsheetApp.openById(this.spreadsheetId);
    const dbSheet = ss.getSheetByName(this.dbSheetName);
    const checkSheet = ss.getSheetByName(this.checkSheetName);

    if (!dbSheet) throw new Error('Не найден лист ДБ.');
    if (!checkSheet) throw new Error('Не найден лист Callibri_Сверка.');

    const checkByDate = this.readCallibriCheck_(checkSheet);
    const result = [];

    dates.forEach(apiDate => {
      const item = checkByDate[this.ruDate_(apiDate)];
      if (!item) {
        result.push(apiDate + ': нет строки в Callibri_Сверка, ДБ не тронут');
        return;
      }

      this.writeDbDate_(dbSheet, apiDate, item);
      result.push(apiDate + ': записано Spam=' + item.spam + ', NonTarget=' + item.nonTarget + ', A=' + item.leadA + ', C=' + item.leadC);
    });

    return { ok: true, mode: 'last_3_workdays_db_write', result: result };
  },

  writeFreshApiResultsToCheckAndDb(apiResults) {
    const ss = SpreadsheetApp.openById(this.spreadsheetId);
    const dbSheet = ss.getSheetByName(this.dbSheetName);
    const checkSheet = ss.getSheetByName(this.checkSheetName);

    if (!dbSheet) throw new Error('Не найден лист ДБ.');
    if (!checkSheet) throw new Error('Не найден лист Callibri_Сверка.');

    const byDate = this.aggregateFreshApiResults_(apiResults);
    const dates = this.yesterdayAndLast3Workdays_();
    const result = [];

    dates.forEach(apiDate => {
      const item = byDate[apiDate] || this.emptyCallibriItem_();
      this.upsertCheckRow_(checkSheet, apiDate, item);
      this.writeDbDate_(dbSheet, apiDate, item);
      result.push(apiDate + ': сверка и ДБ обновлены; Total=' + item.totalYandexCpc + ', Spam=' + item.spam + ', NonTarget=' + item.nonTarget + ', A=' + item.leadA + ', C=' + item.leadC + ', Unknown=' + item.unknownClass);
    });

    return { ok: true, mode: 'fresh_api_to_check_and_db', result: result };
  },

  aggregateFreshApiResults_(apiResults) {
    const result = {};
    const seenRows = {};

    (apiResults || []).forEach(apiResult => {
      const rows = apiResult && apiResult.rows ? apiResult.rows : [];
      rows.forEach(row => {
        const apiDate = String(row.date || '').trim();
        if (!/^\d{4}-\d{2}-\d{2}$/.test(apiDate)) return;

        const rowKey = [
          apiDate,
          String(row.campaignKey || ''),
          this.num_(row.totalYandexCpc),
          this.num_(row.callibriLeadA),
          this.num_(row.callibriLeadB),
          this.num_(row.callibriLeadC),
          this.num_(row.callibriSpam),
          this.num_(row.callibriNonTarget),
          this.num_(row.unknownClass)
        ].join('|');

        if (seenRows[rowKey]) return;
        seenRows[rowKey] = true;

        if (!result[apiDate]) result[apiDate] = this.emptyCallibriItem_();

        result[apiDate].totalYandexCpc += this.num_(row.totalYandexCpc);
        result[apiDate].leadA += this.num_(row.callibriLeadA);
        result[apiDate].leadB += this.num_(row.callibriLeadB);
        result[apiDate].leadC += this.num_(row.callibriLeadC);
        result[apiDate].spam += this.num_(row.callibriSpam);
        result[apiDate].nonTarget += this.num_(row.callibriNonTarget);
        result[apiDate].unknownClass += this.num_(row.unknownClass);
      });
    });

    return result;
  },

  emptyCallibriItem_() {
    return {
      totalYandexCpc: 0,
      leadA: 0,
      leadB: 0,
      leadC: 0,
      spam: 0,
      nonTarget: 0,
      unknownClass: 0
    };
  },

  upsertCheckRow_(sheet, apiDate, item) {
    const ruDate = this.ruDate_(apiDate);
    const row = this.findCheckRowByDate_(sheet, ruDate) || this.findCheckInsertRow_(sheet);
    const comment = item.unknownClass > 0
      ? 'Есть yandex/cpc без явного класса — не считаем лидом'
      : (item.totalYandexCpc > 0 ? 'Статусы обновлены из Callibri; скользящие 7 дней перепроверены; B=' + item.leadB : 'Нет рекламных yandex/cpc; B=' + item.leadB);

    sheet.getRange(row, 1, 1, 8).setValues([[
      ruDate,
      item.totalYandexCpc,
      item.leadA,
      item.leadC,
      item.spam,
      item.nonTarget,
      'Записано',
      comment
    ]]);
  },

  findCheckRowByDate_(sheet, ruDate) {
    const values = sheet.getRange(1, 1, sheet.getLastRow(), 1).getDisplayValues();
    for (let i = 1; i < values.length; i++) {
      if (String(values[i][0] || '').trim() === ruDate) return i + 1;
    }
    return 0;
  },

  findCheckInsertRow_(sheet) {
    const values = sheet.getRange(1, 1, sheet.getLastRow(), 1).getDisplayValues();
    for (let i = 1; i < values.length; i++) {
      const text = String(values[i][0] || '').trim().toLowerCase();
      if (text === 'итого') {
        sheet.insertRowsBefore(i + 1, 1);
        return i + 1;
      }
    }
    return Math.max(sheet.getLastRow() + 1, 2);
  },

  readCallibriCheck_(sheet) {
    const values = sheet.getDataRange().getDisplayValues();
    const result = {};

    for (let r = 1; r < values.length; r++) {
      const row = values[r];
      const date = String(row[0] || '').trim();
      if (!/^\d{2}\.\d{2}\.\d{4}$/.test(date)) continue;

      result[date] = {
        spam: this.num_(row[4]),
        nonTarget: this.num_(row[5]),
        leadA: this.num_(row[2]),
        leadC: this.num_(row[3])
      };
    }

    return result;
  },

  writeDbDate_(sheet, apiDate, item) {
    const headerRow = TDM_DB.findHeaderRowForMonth_(sheet, apiDate);
    const headerMap = TDM_DB.getHeaderMap_(sheet, headerRow);
    const sheetRow = TDM_DB.findDateRow_(sheet, headerMap.date, headerRow, apiDate);

    if (!sheetRow) throw new Error('Не найдена дата в ДБ: ' + apiDate);
    if (this.num_(item.leadB) > 0 && !headerMap.callibriLeadB) {
      throw new Error('Callibri Lead B=' + item.leadB + ', но в ДБ нет отдельной колонки Лид_Квал_B. Значение не прибавлено к A.');
    }

    this.setGoalsBatch_(sheet, sheetRow, [
      { col: headerMap.callibriSpam, value: item.spam, name: 'Callibri Spam' },
      { col: headerMap.callibriNonTarget, value: item.nonTarget, name: 'Callibri NonTarget' },
      { col: headerMap.callibriLeadA, value: item.leadA, name: 'Callibri Lead A' },
      { col: headerMap.callibriLeadB, value: item.leadB, name: 'Callibri Lead B' },
      { col: headerMap.callibriLeadC, value: item.leadC, name: 'Callibri Lead C' }
    ]);
  },

  setGoalsBatch_(sheet, row, goals) {
    const safeGoals = (goals || [])
      .filter(goal => goal && goal.col > 0)
      .sort((a, b) => a.col - b.col);

    if (!safeGoals.length) {
      throw new Error('В ДБ не найдены Callibri-колонки для записи, строка ' + row);
    }

    safeGoals.forEach(goal => {
      const formula = sheet.getRange(row, goal.col).getFormula();
      if (formula) {
        throw new Error(goal.name + ': ячейка содержит формулу, строка ' + row + ', колонка ' + goal.col);
      }
    });

    let group = [];
    const flushGroup = () => {
      if (!group.length) return;

      const startCol = group[0].col;
      const values = [group.map(goal => {
        const number = Number(goal.value || 0);
        return number > 0 ? number : '-';
      })];

      sheet.getRange(row, startCol, 1, group.length).setValues(values);
      group = [];
    };

    safeGoals.forEach(goal => {
      if (!group.length || goal.col === group[group.length - 1].col + 1) {
        group.push(goal);
        return;
      }

      flushGroup();
      group.push(goal);
    });

    flushGroup();
  },

  setGoal_(sheet, row, col, value) {
    this.setGoalsBatch_(sheet, row, [{ col: col, value: value, name: 'Callibri goal' }]);
  },

  yesterdayAndLast3Workdays_() {
    const result = [];
    const seen = {};
    const yesterday = new Date();
    yesterday.setDate(yesterday.getDate() - 1);

    const addDate = date => {
      const apiDate = Utilities.formatDate(date, this.timezone, 'yyyy-MM-dd');
      if (!seen[apiDate]) {
        seen[apiDate] = true;
        result.push(apiDate);
      }
    };

    // Вчера обновляем каждый день, даже если это выходной.
    addDate(yesterday);

    const cursor = new Date(yesterday);
    while (result.length < 4) {
      const day = cursor.getDay();
      if (day !== 0 && day !== 6) {
        addDate(cursor);
      }
      cursor.setDate(cursor.getDate() - 1);
    }

    return result;
  },

  ruDate_(apiDate) {
    return Utilities.formatDate(TDM_DB.parseApiDate_(apiDate), this.timezone, 'dd.MM.yyyy');
  },

  num_(value) {
    return TDM_DB.toNumber_(value);
  }
};

function tdmAutoWeeklyMoReports20260707() {
  return tdmRunVerifiedMondayReports20260713();
}

function archived_tdmInstallAutoReportTriggers20260707() {
  return tdmStabilizeAutomationTriggers20260713();
}

const TDM_AUTO_20260707 = {
  spreadsheetId: '1zXDDfiRYHJE34iOJg-QfAC0bvoTKMIAuVMfsJS0z5Dg',
  workSheetName: 'Рабочие',
  timezone: 'Europe/Moscow',
  dailyHandler: 'tdmAutoDailyReports20260707',
  weeklyHandler: 'tdmAutoWeeklyMoReports20260707',

  daily() {
    const startedAt = new Date();
    const results = [];

    results.push(this.safeCall_('fillTdmDbYesterday'));
    results.push(this.safeCall_('tdmDailyCallibriSync20260708'));
    // Региональный контур возвращён к детальному листу «Регионы_города».
    // Упрощённый отчёт из названий РК не запускаем в daily runtime.
    results.push(this.safeCall_('tdmUpdateRegionsCitiesReport'));
    results.push(this.safeCall_('updateRegionsCitiesReport'));
    results.push(this.safeCall_('tdmUpdateWeeklyDynamicCampaignLists20260708'));
    results.push(this.safeCall_('tdmInstallAllDynamicCommentsQualC20260708'));

    this.writeStatus_('daily', startedAt, results);
  },

  weekly() {
    const startedAt = new Date();
    const results = [];

    results.push(this.safeCall_('tdmRunVerifiedMondayReports20260713'));

    this.writeStatus_('weekly', startedAt, results);
  },

  installTriggers() {
    this.deleteOwnTriggers_();

    ScriptApp.newTrigger(this.dailyHandler)
      .timeBased()
      .everyDays(1)
      .atHour(8)
      .nearMinute(20)
      .inTimezone(this.timezone)
      .create();

    ScriptApp.newTrigger(this.weeklyHandler)
      .timeBased()
      .onWeekDay(ScriptApp.WeekDay.MONDAY)
      .atHour(9)
      .nearMinute(20)
      .inTimezone(this.timezone)
      .create();

    this.writeStatus_('install', new Date(), [
      { name: this.dailyHandler, ok: true, message: 'ежедневно около 08:20 МСК' },
      { name: this.weeklyHandler, ok: true, message: 'понедельник около 09:20 МСК' }
    ]);
  },

  deleteOwnTriggers_() {
    const handlers = [this.dailyHandler, this.weeklyHandler];

    ScriptApp.getProjectTriggers().forEach(trigger => {
      if (handlers.indexOf(trigger.getHandlerFunction()) !== -1) {
        ScriptApp.deleteTrigger(trigger);
      }
    });
  },

  safeCall_(functionName) {
    const startedAt = new Date();

    try {
      const fn = this.getFunctionByName_(functionName);

      if (typeof fn !== 'function') {
        return { name: functionName, ok: false, message: 'function_not_found', seconds: 0 };
      }

      fn();

      return {
        name: functionName,
        ok: true,
        message: 'ok',
        seconds: Math.round((new Date() - startedAt) / 1000)
      };
    } catch (e) {
      return {
        name: functionName,
        ok: false,
        message: String(e && e.message ? e.message : e),
        seconds: Math.round((new Date() - startedAt) / 1000)
      };
    }
  },

  getFunctionByName_(functionName) {
    return Function('return (typeof ' + functionName + ' === "function") ? ' + functionName + ' : null;')();
  },

  writeStatus_(mode, startedAt, results) {
    const ss = SpreadsheetApp.openById(this.spreadsheetId);
    const sheet = ss.getSheetByName(this.workSheetName) || ss.insertSheet(this.workSheetName);
    const now = new Date();
    const okCount = results.filter(item => item.ok).length;
    const failCount = results.length - okCount;
    const status = failCount ? 'FAIL_PARTIAL' : 'OK';
    const details = results.map(item => {
      return item.name + ': ' + (item.ok ? 'OK' : 'ERROR') + (item.message ? ' — ' + item.message : '');
    }).join('\n');

    const row = [
      Utilities.formatDate(now, this.timezone, 'yyyy-MM-dd HH:mm:ss'),
      mode,
      status,
      okCount,
      failCount,
      Utilities.formatDate(startedAt, this.timezone, 'yyyy-MM-dd HH:mm:ss'),
      Math.round((now - startedAt) / 1000),
      details
    ];

    sheet.getRange(Math.max(sheet.getLastRow() + 1, 6), 1, 1, row.length).setValues([row]);
  }
};

function tdmCallibriDailyHardSync20260709() {
  return TDM_CALLIBRI_DAILY_20260709.run();
}

function tdmInstallCallibriDailyHardTrigger20260709() {
  // Старый hard-триггер заменён стабильным rolling-7-days контуром.
  return tdmInstallCallibriStableTriggers20260709();
}

const TDM_CALLIBRI_DAILY_20260709 = {
  spreadsheetId: '1zXDDfiRYHJE34iOJg-QfAC0bvoTKMIAuVMfsJS0z5Dg',
  workSheetName: 'Рабочие',
  timezone: 'Europe/Moscow',
  handler: 'tdmCallibriDailyHardSync20260709',

  run() {
    const startedAt = new Date();
    let status = 'OK';
    let message = 'ok';
    let data = null;

    try {
      data = tdmDailyCallibriSync20260708();
    } catch (e) {
      status = 'ERROR';
      message = String(e && e.message ? e.message : e);
    }

    this.writeStatus_(startedAt, status, message, data);

    if (status !== 'OK') {
      throw new Error(message);
    }

    return { ok: true, mode: 'callibri_daily_hard_sync', data: data };
  },

  installTriggers() {
    this.deleteOwnTriggers_();

    ScriptApp.newTrigger(this.handler)
      .timeBased()
      .everyDays(1)
      .atHour(9)
      .nearMinute(10)
      .inTimezone(this.timezone)
      .create();

    ScriptApp.newTrigger(this.handler)
      .timeBased()
      .everyDays(1)
      .atHour(12)
      .nearMinute(10)
      .inTimezone(this.timezone)
      .create();

    this.writeStatus_(new Date(), 'INSTALL_OK', 'Callibri-only triggers: 09:10 + 12:10 МСК', null);

    return { ok: true, handler: this.handler, schedule: ['09:10 МСК', '12:10 МСК'] };
  },

  deleteOwnTriggers_() {
    ScriptApp.getProjectTriggers().forEach(trigger => {
      if (trigger.getHandlerFunction() === this.handler) {
        ScriptApp.deleteTrigger(trigger);
      }
    });
  },

  writeStatus_(startedAt, status, message, data) {
    const ss = SpreadsheetApp.openById(this.spreadsheetId);
    const sheet = ss.getSheetByName(this.workSheetName) || ss.insertSheet(this.workSheetName);
    const now = new Date();
    const details = data ? JSON.stringify(data).slice(0, 45000) : message;

    sheet.getRange(Math.max(sheet.getLastRow() + 1, 6), 1, 1, 8).setValues([[
      Utilities.formatDate(now, this.timezone, 'yyyy-MM-dd HH:mm:ss'),
      'callibri_daily_hard_sync',
      status,
      status === 'OK' || status === 'INSTALL_OK' ? 1 : 0,
      status === 'OK' || status === 'INSTALL_OK' ? 0 : 1,
      Utilities.formatDate(startedAt, this.timezone, 'yyyy-MM-dd HH:mm:ss'),
      Math.round((now - startedAt) / 1000),
      details
    ]]);
  }
};

function tdmCallibriDailyYesterdayOnly20260709() {
  return TDM_CALLIBRI_STABLE_20260709.runYesterdayOnly();
}

function tdmCallibriEvery3DaysRecheck20260709() {
  return {
    ok: true,
    mode: 'legacy_recheck_skipped',
    message: 'Отдельная трёхдневная перепроверка отключена: её полностью заменяет ежедневное окно 7 дней.'
  };
}

function tdmCallibriRetryYesterdayUntilSuccess20260709() {
  return TDM_CALLIBRI_STABLE_20260709.retryYesterdayUntilSuccessToday();
}

function tdmInstallCallibriStableTriggers20260709() {
  return TDM_CALLIBRI_STABLE_20260709.installStableTriggers();
}

function tdmInstallCallibriRetryToday20260709() {
  return TDM_CALLIBRI_STABLE_20260709.installRetryTodayAfter2h();
}

const TDM_CALLIBRI_STABLE_20260709 = {
  spreadsheetId: '1zXDDfiRYHJE34iOJg-QfAC0bvoTKMIAuVMfsJS0z5Dg',
  workSheetName: 'Рабочие',
  timezone: 'Europe/Moscow',
  dailyHandler: 'tdmCallibriDailyYesterdayOnly20260709',
  recheckHandler: 'tdmCallibriEvery3DaysRecheck20260709',
  retryHandler: 'tdmCallibriRetryYesterdayUntilSuccess20260709',
  oldHardHandler: 'tdmCallibriDailyHardSync20260709',

  runYesterdayOnly() {
    const startedAt = new Date();
    try {
      const result = tdmDailyCallibriSync20260708();
      this.writeStatus_(startedAt, 'callibri_rolling_7_days', 'OK', result);
      return result;
    } catch (e) {
      const message = String(e && e.message ? e.message : e);
      this.writeStatus_(startedAt, 'callibri_rolling_7_days', 'ERROR', message);
      throw new Error(message);
    }
  },

  runLast3WorkdaysOnly() {
    return { ok: true, mode: 'legacy_recheck_skipped', coveredBy: this.dailyHandler };
  },

  retryYesterdayUntilSuccessToday() {
    const startedAt = new Date();
    try {
      const result = this.runYesterdayOnly();
      this.deleteTriggersFor_([this.retryHandler]);
      this.writeStatus_(startedAt, 'callibri_retry_until_success_today', 'OK_STOP_RETRY', result);
      return result;
    } catch (e) {
      const message = String(e && e.message ? e.message : e);
      this.writeStatus_(startedAt, 'callibri_retry_until_success_today', 'ERROR_WILL_RETRY_IF_TODAY', message);
      this.scheduleNextRetryIfToday_();
      throw new Error(message);
    }
  },

  installStableTriggers() {
    this.deleteTriggersFor_([this.oldHardHandler, this.dailyHandler, this.recheckHandler]);

    ScriptApp.newTrigger(this.dailyHandler)
      .timeBased()
      .everyDays(1)
      .atHour(9)
      .nearMinute(10)
      .inTimezone(this.timezone)
      .create();

    const result = { ok: true, schedule: ['daily rolling-7-days 09:10 МСК'], disabledLegacyHandler: this.recheckHandler };
    this.writeStatus_(new Date(), 'callibri_stable_triggers_install', 'INSTALL_OK', result);
    return result;
  },

  installRetryTodayAfter2h() {
    this.deleteTriggersFor_([this.retryHandler]);

    ScriptApp.newTrigger(this.retryHandler)
      .timeBased()
      .after(2 * 60 * 60 * 1000)
      .create();

    const result = { ok: true, retry: 'after 2 hours, then every 2 hours today until success', handler: this.retryHandler };
    this.writeStatus_(new Date(), 'callibri_retry_today_install', 'INSTALL_OK', result);
    return result;
  },

  scheduleNextRetryIfToday_() {
    if (!this.canRetryToday_()) return;
    this.deleteTriggersFor_([this.retryHandler]);
    ScriptApp.newTrigger(this.retryHandler)
      .timeBased()
      .after(2 * 60 * 60 * 1000)
      .create();
  },

  canRetryToday_() {
    const now = new Date();
    const hour = Number(Utilities.formatDate(now, this.timezone, 'H'));
    return hour < 22;
  },

  writeSelectedDates_(apiResults, apiDates) {
    const ss = SpreadsheetApp.openById(this.spreadsheetId);
    const dbSheet = ss.getSheetByName(TDM_DB_CALLIBRI_20260708.dbSheetName);
    const checkSheet = ss.getSheetByName(TDM_DB_CALLIBRI_20260708.checkSheetName);

    if (!dbSheet) throw new Error('Не найден лист ДБ.');
    if (!checkSheet) throw new Error('Не найден лист Callibri_Сверка.');

    const byDate = TDM_DB_CALLIBRI_20260708.aggregateFreshApiResults_(apiResults || []);
    const result = [];

    (apiDates || []).forEach(apiDate => {
      const item = byDate[apiDate] || TDM_DB_CALLIBRI_20260708.emptyCallibriItem_();
      TDM_DB_CALLIBRI_20260708.upsertCheckRow_(checkSheet, apiDate, item);
      TDM_DB_CALLIBRI_20260708.writeDbDate_(dbSheet, apiDate, item);
      result.push(apiDate + ': written Total=' + item.totalYandexCpc + ', Spam=' + item.spam + ', NonTarget=' + item.nonTarget + ', A=' + item.leadA + ', B=' + item.leadB + ', C=' + item.leadC + ', Unknown=' + item.unknownClass);
    });

    return result;
  },

  deleteTriggersFor_(handlers) {
    const allowed = handlers || [];
    ScriptApp.getProjectTriggers().forEach(trigger => {
      if (allowed.indexOf(trigger.getHandlerFunction()) !== -1) {
        ScriptApp.deleteTrigger(trigger);
      }
    });
  },

  yesterdayApi_() {
    const date = new Date();
    date.setDate(date.getDate() - 1);
    return Utilities.formatDate(date, this.timezone, 'yyyy-MM-dd');
  },

  last3WorkdaysApi_() {
    const result = [];
    const cursor = new Date();
    cursor.setDate(cursor.getDate() - 1);

    while (result.length < 3) {
      const day = cursor.getDay();
      if (day !== 0 && day !== 6) {
        result.push(Utilities.formatDate(cursor, this.timezone, 'yyyy-MM-dd'));
      }
      cursor.setDate(cursor.getDate() - 1);
    }

    return result;
  },

  writeStatus_(startedAt, mode, status, payload) {
    const ss = SpreadsheetApp.openById(this.spreadsheetId);
    const sheet = ss.getSheetByName(this.workSheetName) || ss.insertSheet(this.workSheetName);
    const now = new Date();
    const details = typeof payload === 'string' ? payload : JSON.stringify(payload).slice(0, 45000);

    sheet.getRange(Math.max(sheet.getLastRow() + 1, 6), 1, 1, 8).setValues([[
      Utilities.formatDate(now, this.timezone, 'yyyy-MM-dd HH:mm:ss'),
      mode,
      status,
      status.indexOf('OK') !== -1 || status === 'INSTALL_OK' ? 1 : 0,
      status.indexOf('OK') !== -1 || status === 'INSTALL_OK' ? 0 : 1,
      Utilities.formatDate(startedAt, this.timezone, 'yyyy-MM-dd HH:mm:ss'),
      Math.round((now - startedAt) / 1000),
      details
    ]]);
  }
};

const TDM_REGION_AUTO_20260707 = {
  spreadsheetId: '1zXDDfiRYHJE34iOJg-QfAC0bvoTKMIAuVMfsJS0z5Dg',
  sheetName: 'Отчет  регионы',
  timezone: 'Europe/Moscow',

  updateCurrentMonthToYesterday() {
    const today = new Date();
    const yesterday = new Date(today);
    yesterday.setDate(today.getDate() - 1);
    const dateFrom = Utilities.formatDate(new Date(yesterday.getFullYear(), yesterday.getMonth(), 1), this.timezone, 'yyyy-MM-dd');
    const dateTo = Utilities.formatDate(yesterday, this.timezone, 'yyyy-MM-dd');
    const rows = this.loadRows_(dateFrom, dateTo);
    const byRegion = this.aggregate_(rows);
    this.write_(dateFrom, dateTo, byRegion);
  },

  loadRows_(dateFrom, dateTo) {
    const goals = ['453453800', '504318736', '504318735', '575188424'];
    const body = {
      params: {
        SelectionCriteria: { DateFrom: dateFrom, DateTo: dateTo },
        Goals: goals.map(Number),
        AttributionModels: ['AUTO'],
        FieldNames: ['CampaignName', 'Impressions', 'Clicks', 'Cost', 'Conversions'],
        ReportName: 'tdm_region_auto_' + dateFrom + '_' + dateTo + '_' + Utilities.getUuid(),
        ReportType: 'CAMPAIGN_PERFORMANCE_REPORT',
        DateRangeType: 'CUSTOM_DATE',
        Format: 'TSV',
        IncludeVAT: 'YES',
        IncludeDiscount: 'NO'
      }
    };
    const text = TDM_DB.requestDirectReport_(body);
    return this.parseTsv_(text, goals);
  },

  parseTsv_(text, goals) {
    const lines = String(text || '').trim().split(/\r?\n/).filter(Boolean);
    const headerIndex = lines.findIndex(line => line.indexOf('CampaignName') !== -1 && line.indexOf('Cost') !== -1);
    if (headerIndex === -1) throw new Error('Региональный отчёт: не найдены заголовки Direct TSV.');
    const headers = lines[headerIndex].split('\t').map(h => String(h).replace(/^\uFEFF/, '').trim());
    const idx = name => headers.indexOf(name);
    const goalIdx = {};
    goals.forEach(goal => {
      goalIdx[goal] = headers.findIndex(h => String(h).indexOf('Conversions_' + goal + '_') === 0 || String(h).indexOf(goal) !== -1);
    });
    const rows = [];
    for (let i = headerIndex + 1; i < lines.length; i++) {
      const cols = lines[i].split('\t');
      if (!cols[idx('CampaignName')]) continue;
      rows.push({
        campaign: cols[idx('CampaignName')],
        impressions: TDM_DB.toNumber_(cols[idx('Impressions')]),
        clicks: TDM_DB.toNumber_(cols[idx('Clicks')]),
        cost: TDM_DB.toNumber_(cols[idx('Cost')]),
        view3: this.goal_(cols, goalIdx['453453800']),
        addToCart: this.goal_(cols, goalIdx['504318736']),
        purchase: this.goal_(cols, goalIdx['504318735']),
        jivo: this.goal_(cols, goalIdx['575188424'])
      });
    }
    return rows;
  },

  goal_(cols, index) {
    return index === -1 || index === undefined ? 0 : TDM_DB.toNumber_(cols[index]);
  },

  aggregate_(rows) {
    const result = {};
    rows.forEach(row => {
      const region = this.region_(row.campaign);
      if (!result[region]) result[region] = { impressions: 0, clicks: 0, cost: 0, view3: 0, leads: 0 };
      result[region].impressions += row.impressions;
      result[region].clicks += row.clicks;
      result[region].cost += row.cost;
      result[region].view3 += row.view3;
      result[region].leads += row.addToCart + row.purchase + row.jivo;
    });
    return result;
  },

  region_(campaignName) {
    const name = TDM_DB.normalize_(campaignName);
    if (name.indexOf('spb-rf') !== -1 || name.indexOf('rf_spb') !== -1 || name.indexOf('rf-spb') !== -1) return 'СПБ+РФ';
    if (name.indexOf('_kz') !== -1 || name.indexOf('-kz') !== -1 || name.indexOf(' kz') !== -1) return 'Казахстан';
    if (name.indexOf('_spb') !== -1 || name.indexOf('-spb') !== -1 || name.indexOf(' spb') !== -1) return 'СПБ';
    if (name.indexOf('_rf') !== -1 || name.indexOf('-rf') !== -1 || name.indexOf(' rf') !== -1) return 'РФ';
    return 'Не определено';
  },

  write_(dateFrom, dateTo, byRegion) {
    const ss = SpreadsheetApp.openById(this.spreadsheetId);
    const sheet = ss.getSheetByName(this.sheetName) || ss.insertSheet(this.sheetName);
    const order = ['СПБ', 'РФ', 'Казахстан', 'СПБ+РФ', 'Не определено'];
    const values = [];
    values.push(['Отчёт по регионам из названия РК — ' + this.ruDate_(dateFrom) + '–' + this.ruDate_(dateTo), '', '', '', '', '', '', '', '', '', '', '', '', '']);
    values.push(['Логика: регион берётся из названия кампании. Лиды = Ecommerce добавление в корзину + Ecommerce покупка + Jivo. Callibri отдельно не добавлен, пока API Callibri не проходит в ежедневном контуре.', '', '', '', '', '', '', '', '', '', '', '', '', '']);
    values.push(['', '', '', '', '', '', '', '', '', '', '', '', '', '']);
    values.push(['Регион', 'Показы', 'Клики', 'CTR', 'Расход, ₽', 'Просмотр 3х страниц', 'Отказы', 'Глубина', 'Лиды без Callibri', 'Доля лидов', 'CR по лидам', 'CPA по лидам', 'Оценка', 'Комментарий']);
    const total = { impressions: 0, clicks: 0, cost: 0, view3: 0, leads: 0 };
    order.forEach(region => {
      const item = byRegion[region];
      if (!item) return;
      total.impressions += item.impressions;
      total.clicks += item.clicks;
      total.cost += item.cost;
      total.view3 += item.view3;
      total.leads += item.leads;
    });
    order.forEach(region => {
      const item = byRegion[region];
      if (!item) return;
      values.push([region, item.impressions, item.clicks, this.pct_(item.clicks, item.impressions), item.cost, item.view3, '', '', item.leads, this.pct_(item.leads, total.leads), this.pct_(item.leads, item.clicks), item.leads ? item.cost / item.leads : 0, item.leads ? 'Контроль' : 'Нет лидов', region + ': ' + item.leads + ' лидов без Callibri, расход ' + Math.round(item.cost) + ' ₽. Callibri-классы подключить отдельно.']);
    });
    values.push(['Итого', total.impressions, total.clicks, this.pct_(total.clicks, total.impressions), total.cost, total.view3, '', '', total.leads, '100,00%', this.pct_(total.leads, total.clicks), total.leads ? total.cost / total.leads : 0, '', 'Итого без Callibri за период ' + this.ruDate_(dateFrom) + '–' + this.ruDate_(dateTo) + '.']);
    sheet.getRange(1, 1, Math.max(sheet.getMaxRows(), 30), 14).breakApart();
    sheet.clearContents();
    sheet.getRange(1, 1, values.length, 14).setValues(values);
    sheet.getRange(4, 1, 1, 14).setFontWeight('bold');
    sheet.autoResizeColumns(1, 14);
  },

  pct_(a, b) {
    return b ? Utilities.formatString('%.2f%%', (a / b) * 100).replace('.', ',') : '0,00%';
  },

  ruDate_(apiDate) {
    return Utilities.formatDate(TDM_DB.parseApiDate_(apiDate), this.timezone, 'dd.MM.yyyy');
  }
};

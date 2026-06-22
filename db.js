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

function createTdmDbTrigger8am() {
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
      view3Pages: ['453453800'],      // Просмотр 3х страниц
      emailClick: ['388393244'],      // Клик: По email адресу
      phoneClick: ['432293138'],      // Клик: По номеру телефона
      talkMeOnline: ['432303787'],    // TalkMe: Клиент написал в чат онлайн
      addToCart: ['504318736'],       // Ecommerce: добавление в корзину
      purchase: ['504318735'],        // Ecommerce: покупка
      talkMe: ['432309471'],          // TalkMe: Клиент написал в чат офлайн
      email: ['541674471'],           // Email tracking Roistat
      calls: ['512458209']            // Звонок
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
      this.setGoalCell_(sheet, sheetRow, headerMap.emailClick, item.emailClick);
      this.setGoalCell_(sheet, sheetRow, headerMap.phoneClick, item.phoneClick);
      this.setGoalCell_(sheet, sheetRow, headerMap.talkMeOnline, item.talkMeOnline);
      this.setGoalCell_(sheet, sheetRow, headerMap.addToCart, item.addToCart);
      this.setGoalCell_(sheet, sheetRow, headerMap.purchase, item.purchase);
      this.setGoalCell_(sheet, sheetRow, headerMap.talkMe, item.talkMe);
      this.setGoalCell_(sheet, sheetRow, headerMap.email, item.email);
      this.setGoalCell_(sheet, sheetRow, headerMap.calls, item.calls);
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

    this.clearObsoleteRoistatColumns_(sheet, headerRow, headerMap.date);
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

    this.clearObsoleteRoistatColumns_(sheet, headerRow, headerMap.date);
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



  clearObsoleteRoistatColumns_(sheet, headerRow, dateCol) {
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

      emailClick: this.findContainsCol_(map, ['клик: по email адресу', 'клик по email адресу']),
      phoneClick: this.findContainsCol_(map, ['клик: по номеру телефона', 'клик по номеру телефона']),
      talkMeOnline: this.findContainsCol_(map, ['talkme: клиент написал в чат онлайн', 'talkme клиент написал в чат онлайн', 'talkme: клиент написал в чат', 'talkme клиент написал в чат']),
      addToCart: this.findContainsCol_(map, ['ecommerce: добавление в корзину', 'ecommerce добавление в корзину']),
      purchase: this.findContainsCol_(map, ['ecommerce: покупка', 'ecommerce покупка']),

      talkMe: this.findExactCol_(map, 'talk me'),
      email: this.findExactCol_(map, 'email'),
      calls: this.findExactCol_(map, 'звонки')
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
      headerMap.emailClick,
      headerMap.phoneClick,
      headerMap.talkMeOnline,
      headerMap.addToCart,
      headerMap.purchase,
      headerMap.talkMe,
      headerMap.email,
      headerMap.calls
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
      result[apiDate].emailClick += this.getGoalSum_(row, goalIndexes, this.config.goals.emailClick);
      result[apiDate].phoneClick += this.getGoalSum_(row, goalIndexes, this.config.goals.phoneClick);
      result[apiDate].talkMeOnline += this.getGoalSum_(row, goalIndexes, this.config.goals.talkMeOnline);
      result[apiDate].addToCart += this.getGoalSum_(row, goalIndexes, this.config.goals.addToCart);
      result[apiDate].purchase += this.getGoalSum_(row, goalIndexes, this.config.goals.purchase);
      result[apiDate].talkMe += this.getGoalSum_(row, goalIndexes, this.config.goals.talkMe);
      result[apiDate].email += this.getGoalSum_(row, goalIndexes, this.config.goals.email);
      result[apiDate].calls += this.getGoalSum_(row, goalIndexes, this.config.goals.calls);
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
      emailClick: 0,
      phoneClick: 0,
      talkMeOnline: 0,
      addToCart: 0,
      purchase: 0,
      talkMe: 0,
      email: 0,
      calls: 0
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

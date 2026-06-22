// Safe patch for ENO weekly report.
// Purpose: run ENO report against the exact TDM spreadsheet by ID, not active spreadsheet.
// This keeps the existing eno.js logic intact and adds safe entry points for manual check and trigger migration.

const TDME_SPREADSHEET_ID_SAFE = '1zXDDfiRYHJE34iOJg-QfAC0bvoTKMIAuVMfsJS0z5Dg';

function fillTdmEnoPreviousFullWeekFixed() {
  try {
    const period = tdmePreviousFullWeek_();
    const ss = SpreadsheetApp.openById(TDME_SPREADSHEET_ID_SAFE);
    tdmeFillReportBySpreadsheet_(ss, period.dateFrom, period.dateTo, null);
  } catch (e) {
    tdmeNotifyError_('ЕНО не обновился — fixed openById', e);
    throw e;
  }
}

function fillTdmEnoMissingWeekJune15to21() {
  const ss = SpreadsheetApp.openById(TDME_SPREADSHEET_ID_SAFE);
  tdmeFillReportBySpreadsheet_(ss, '2026-06-15', '2026-06-21', null);
}

function createTdmEnoWeeklyTriggerMonday9amFixed() {
  const triggers = ScriptApp.getProjectTriggers();

  triggers.forEach(trigger => {
    const handler = trigger.getHandlerFunction();
    if (handler === 'fillTdmEnoPreviousFullWeek' || handler === 'fillTdmEnoPreviousFullWeekFixed') {
      ScriptApp.deleteTrigger(trigger);
    }
  });

  ScriptApp.newTrigger('fillTdmEnoPreviousFullWeekFixed')
    .timeBased()
    .onWeekDay(ScriptApp.WeekDay.MONDAY)
    .atHour(9)
    .inTimezone(TDME_TIMEZONE)
    .create();
}

function tdmeFillReportBySpreadsheet_(ss, dateFrom, dateTo, forcedTopRow) {
  const sheet = ss.getSheetByName(TDME_SHEET_NAME);

  if (!sheet) {
    throw new Error('Не найден лист: ' + TDME_SHEET_NAME);
  }

  const rows = tdmeLoadCampaignRows_(dateFrom, dateTo);
  rows.sort((a, b) => b.cost - a.cost);

  const monthRows = tdmeLoadCampaignRows_(tdmeMonthStart_(dateTo), dateTo);

  const title = tdmeTitle_(dateFrom, dateTo);
  const topRow = forcedTopRow || tdmeNextTopRow_(sheet, title);

  tdmeCopyPreviousReportBlock_(sheet, topRow, title);

  let block = tdmeGetBlock_(sheet, topRow);
  tdmeRemoveExtraUniqueCallColumn_(sheet, block);

  block = tdmeGetBlock_(sheet, topRow);
  block = tdmeFitCampaignRows_(sheet, block, rows.length);

  const header = tdmeHeaderMap_(sheet, block.headerRow);
  const dataStartRow = block.headerRow + 1;
  const dataRowsCount = Math.max(rows.length, 1);
  const dataEndRow = dataStartRow + dataRowsCount - 1;

  block.totalRow = tdmeForceTotalRowAfterData_(sheet, header, dataEndRow, block.totalRow);
  tdmeApplyUpdatedGoalColumns_(sheet, block.headerRow, block.totalRow, header);

  tdmeSetPlain_(sheet, block.topRow, 1, title);

  tdmeClearDataInputCells_(sheet, header, dataStartRow, dataRowsCount);
  tdmeFillRows_(sheet, header, dataStartRow, rows);

  tdmeEnsureTotalLabel_(sheet, header, block.totalRow);
  tdmeFixTotalRowFormulaRanges_(sheet, block.totalRow, dataStartRow, dataEndRow);
  tdmeFixTopSummaryFormulas_(sheet, block.topRow, block.totalRow);

  tdmeFormatBounceColumn_(sheet, header, dataStartRow, dataRowsCount, block.totalRow);
  tdmeHighlightGreenRows_(sheet, header, dataStartRow, dataRowsCount, rows);

  const comments = tdmeBuildComments_(rows, monthRows, dateFrom, dateTo);
  const commentRow = tdmeFindCommentRow_(sheet, block.totalRow) || block.totalRow + 3;
  tdmeWriteComments_(sheet, commentRow, comments);

  tdmeValidateReportBlock_(sheet, block, dataStartRow, dataEndRow, title);
}

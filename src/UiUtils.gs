function showProgressToast(message, title, timeoutSeconds) {
  SpreadsheetApp.getActiveSpreadsheet().toast(message, title || 'Progress', timeoutSeconds || 5);
  Logger.log(`Toast: [${title || 'Progress'}] ${message}`);
}
function showStandardDialog(title, message) {
  SpreadsheetApp.getUi().alert(title, message, SpreadsheetApp.getUi().ButtonSet.OK);
  Logger.log(`Standard Dialog: [${title}] ${message}`);
}
function buildSummaryMessage(items) {
  return items.join('\n');
}

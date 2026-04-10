function showProgressToast(message, title, timeoutSeconds) {
  SpreadsheetApp.getActiveSpreadsheet().toast(message, title || 'Progress', timeoutSeconds || 5);
  Logger.log(`Toast: [${title || 'Progress'}] ${message}`);
}
function showStandardDialog(title, message) {
  SpreadsheetApp.getUi().alert(title, message, SpreadsheetApp.getUi().ButtonSet.OK);
  Logger.log(`Standard Dialog: [${title}] ${message}`);
}
/**
 * Shows a modal dialog with plain-text lines and an optional clickable link.
 * @param {string}   title      - Dialog title bar text.
 * @param {string[]} textLines  - Array of strings, each rendered as a paragraph.
 * @param {string}   [linkUrl]  - If provided, renders a clickable anchor at the bottom.
 * @param {string}   [linkLabel]- Display text for the link (defaults to the URL).
 */
function showHtmlSummaryDialog(title, textLines, linkUrl, linkLabel) {
  function esc(s) {
    return String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
  }
  var body = textLines.map(function(l) { return '<p style="margin:0 0 8px">' + esc(l) + '</p>'; }).join('');
  if (linkUrl) {
    body += '<p style="margin-top:14px"><a href="' + esc(linkUrl) + '" target="_blank" style="color:#1a73e8">'
          + esc(linkLabel || linkUrl) + '</a></p>';
  }
  var html = HtmlService.createHtmlOutput(
    '<html><body style="font-family:Arial,sans-serif;font-size:13px;padding:16px;color:#202124">'
    + body + '</body></html>'
  ).setWidth(460).setHeight(300);
  SpreadsheetApp.getUi().showModalDialog(html, title);
  Logger.log('HTML Dialog: [' + title + ']');
}

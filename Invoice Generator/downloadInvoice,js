/**
 * @OnlyCurrentDoc
 */
function downloadInvoice() {
  // get invoice id from Invoice Template sheet cell D9
  const ss = SpreadsheetApp.getActive();
  const sheet = ss.getSheetByName('Invoice Template'); // exact name, adjust if different
  const invoiceId = (sheet.getRange('D9').getValue() || '').toString().trim();
  if (!invoiceId) {
    SpreadsheetApp.getUi().alert('No invoice id found in D9');
    return;
  }

  // your working webapp base (without invoice param)
  const webAppBase = 'LINK TO WEBAPP BASE';

  // build the full URL (encoded)
  const url = webAppBase + '?invoice=' + encodeURIComponent(invoiceId);

  // create template, inject url, show modal
  const tpl = HtmlService.createTemplateFromFile('download');
  tpl.url = url;
  const html = tpl.evaluate()
    .setWidth(10)
    .setHeight(10);
  SpreadsheetApp.getUi().showModalDialog(html, 'Downloading...');
}

/**
 * @OnlyCurrentDoc
 */
function generateInvoicePDF() {
  const ss = SpreadsheetApp.getActive();
  const templateSheet = ss.getSheetByName("Invoice Template");
  const invoiceId = templateSheet.getRange("d9").getValue().trim();                          // D9 is the cell on the Invoice Template Sheet with the Invoice_Id dropdown
  const customerName = templateSheet.getRange("A10").getValue().trim();                      // A10 is the cell on the Invoice Template Sheet with the customer name
  if (!invoiceId) return;

  const folderId = "FOLDER_ID-FROM-GOOGLE_DRIVE";
  const folder = DriveApp.getFolderById(folderId);

  const url = ss.getUrl().replace(/edit$/, '') + 
    'export?format=pdf&portrait=true&size=A4&sheetnames=false' +
    '&printtitle=false&pagenumbers=false&gridlines=false&fzr=false' +
    '&gid=' + templateSheet.getSheetId();

  const token = ScriptApp.getOAuthToken();

  const response = UrlFetchApp.fetch(url, {
    headers: { Authorization: 'Bearer ' + token },
    muteHttpExceptions: true
  });

  const blob = response.getBlob().setName(customerName + " " + invoiceId + ".pdf");
  folder.createFile(blob);

  return true;  // clean termination
}

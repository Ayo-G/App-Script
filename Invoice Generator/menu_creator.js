/**
 * @OnlyCurrentDoc
 */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("Invoice PDF")
    .addItem("Save Invoice to Drive", "generateInvoicePDF")
    .addItem("Download Invoice", "downloadInvoice")
    .addToUi();
}

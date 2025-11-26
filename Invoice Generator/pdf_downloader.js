/**
 * @OnlyCurrentDoc
 */
// Web app: generate PDF from a sheet and trigger browser download via base64 data URI.
// Deploy as "Web app" and set access appropriately (Anyone with link).

function doGet(e) {
  try {
    const ssId = "SHEET_ID";         // <-- replace with your sheet ID
    const invoiceParam = e.parameter.invoice || ""; // invoice id passed in URL ?invoice=INV-001
    if (!invoiceParam) {
      return HtmlService.createHtmlOutput("Missing invoice parameter. Use ?invoice=INV-001");
    }

    const ss = SpreadsheetApp.openById(ssId);
    const templateSheet = ss.getSheetByName("Invoice Template");
    if (!templateSheet) {
      return HtmlService.createHtmlOutput("Invoice Template sheet not found.");
    }

    // Build export URL
    const exportUrl = ss.getUrl().replace(/edit$/,'') +
      'export?format=pdf' +
      '&size=A4' +
      '&portrait=true' +
      '&fitw=true' +               // fit to width
      '&sheetnames=false' +
      '&printtitle=false' +
      '&pagenumbers=false' +
      '&gridlines=false' +
      '&fzr=false' +
      '&gid=' + templateSheet.getSheetId();

    const token = ScriptApp.getOAuthToken();
    const response = UrlFetchApp.fetch(exportUrl, {
      headers: { Authorization: 'Bearer ' + token },
      muteHttpExceptions: true
    });

    const blob = response.getBlob();
    const pdfBytes = blob.getBytes();
    const base64 = Utilities.base64Encode(pdfBytes);
    const safeFilename = invoiceParam.replace(/[^a-zA-Z0-9_\-\.]/g, "_") + ".pdf";

    // HTML page that auto-downloads the file using a data URI
    const html = `
      <!doctype html>
      <html>
        <head>
          <meta charset="utf-8">
          <title>Download ${safeFilename}</title>
        </head>
        <body>
          <a id="dl" href="data:application/pdf;base64,${base64}" download="${safeFilename}">Download</a>
          <script>
            // Auto-click the download link
            (function(){
              const a = document.getElementById('dl');
              if (a) {
                a.click();
                // optional: close window after delay (works if opened in a new tab)
                setTimeout(()=>{ try { window.close(); } catch(e){} }, 1000);
              }
            })();
          </script>
        </body>
      </html>
    `;

    return HtmlService.createHtmlOutput(html)
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);

  } catch (err) {
    return HtmlService.createHtmlOutput("Error: " + (err.message || err));
  }
}

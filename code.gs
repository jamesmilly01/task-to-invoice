const TEMPLATE_DOC_ID = '1LfppvDXpHc5hR_bMKSuIq9BXfjsLfNuunwj-9UjZ-DE';
const SPREADSHEET_ID = '1OTWAiQ96x0A3pim7M8p2XY_tE_z3UO7YmbY1QBcIZfA';
const MASTER_FOLDER_NAME = 'Task Invoices';
const GST_RATE = 0.18; 

function withRetry(fn, retries = 3, delay = 2000) {
  for (let i = 0; i < retries; i++) {
    try {
      return fn();
    } catch (e) {
      if (i === retries - 1) throw e;
      Logger.log(`Retry ${i + 1}: ${e.message}`);
      Utilities.sleep(delay);
    }
  }
}

function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('Task-to-Invoice Portal')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function saveTaskAndGenerateInvoice(data) {
  try {
    const sheet = withRetry(() =>
      SpreadsheetApp.openById(SPREADSHEET_ID).getActiveSheet()
    );

    const conversionRate = getCurrencyRate(data.currency);
    const subtotal       = data.hours * data.rate * conversionRate;
    const gstAmount      = subtotal * GST_RATE;
    const grandTotal     = subtotal + gstAmount;

    const invoiceLink = withRetry(() =>
      generateInvoice(data, subtotal, gstAmount, grandTotal)
    );

    withRetry(() =>
      sheet.appendRow([
        new Date(),
        data.taskName,
        data.hours,
        data.rate,
        data.clientName,
        data.clientEmail,
        data.currency,
        subtotal.toFixed(2),
        gstAmount.toFixed(2),
        grandTotal.toFixed(2),
        invoiceLink
      ])
    );

    sendInvoiceEmail(data.clientEmail, data.clientName, invoiceLink);

    return { success: true, link: invoiceLink };

  } catch (e) {
    return { success: false, error: e.message };
  }
}

function getCurrencyRate(currency) {
  try {
    if (currency === 'INR') return 1;

    const url      = `https://open.er-api.com/v6/latest/${currency}`;
    const response = UrlFetchApp.fetch(url);
    const json     = JSON.parse(response.getContentText());

    if (json.result === 'success') {
      return json.rates['INR'] || 1;
    } else {
      throw new Error('Invalid API response');
    }
  } catch (e) {
    throw new Error('Currency API failed: ' + e.message);
  }
}

function generateInvoice(data, subtotal, gstAmount, grandTotal) {
  const templateFile = withRetry(() =>
    DriveApp.getFileById(TEMPLATE_DOC_ID)
  );
  const folder = withRetry(() =>
    getOrCreateClientFolder(data.clientName)
  );

  const today    = Utilities.formatDate(
    new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd'
  );
  const invoiceNo = getNextInvoiceNumber();
  const fileName  = `Invoice_${today}_${data.clientName}`;

  const newDoc = withRetry(() =>
    templateFile.makeCopy(fileName, folder)
  );
  const doc  = withRetry(() =>
    DocumentApp.openById(newDoc.getId())
  );
  const body = doc.getBody();

  body.replaceText('{{CLIENT_NAME}}',    data.clientName);
  body.replaceText('{{CLIENT_EMAIL}}',   data.clientEmail);
  body.replaceText('{{TASK_NAME}}',      data.taskName);
  body.replaceText('{{HOURS}}',          data.hours.toString());
  body.replaceText('{{RATE}}',           data.rate.toString());
  body.replaceText('{{TOTAL}}',          subtotal.toFixed(2));
  body.replaceText('{{GST_AMOUNT}}',     gstAmount.toFixed(2));
  body.replaceText('{{GRAND_TOTAL}}',    grandTotal.toFixed(2));
  body.replaceText('{{DATE}}',           today);
  body.replaceText('{{CURRENCY}}',       data.currency);
  body.replaceText('{{INVOICE_NUMBER}}', invoiceNo);

  doc.saveAndClose();

  const pdfBlob = withRetry(() =>
    DriveApp.getFileById(newDoc.getId()).getAs('application/pdf')
  );
  pdfBlob.setName(fileName + '.pdf');

  const pdfFile = withRetry(() => folder.createFile(pdfBlob));

  withRetry(() =>
    pdfFile.setSharing(
      DriveApp.Access.ANYONE_WITH_LINK,
      DriveApp.Permission.VIEW
    )
  );

  newDoc.setTrashed(true);
  return pdfFile.getUrl();
}

function sendInvoiceEmail(clientEmail, clientName, invoiceLink) {
  try {
    const subject = `Your Invoice is Ready — ${clientName}`;
    const body    = `Dear ${clientName},\n\n` +
      `Please find your invoice attached below:\n${invoiceLink}\n\n` +
      `Thank you for your business!\n\nRegards,\nTask-to-Invoice Portal`;

    GmailApp.sendEmail(clientEmail, subject, body);
  } catch (e) {
    Logger.log('Email failed: ' + e.message);
  }
}

function getNextInvoiceNumber() {
  const sheet  = SpreadsheetApp.openById(SPREADSHEET_ID).getActiveSheet();
  const lastRow = sheet.getLastRow();
  return 'INV-' + String(lastRow).padStart(3, '0');
}

function getOrCreateClientFolder(clientName) {
  const folders      = DriveApp.getFoldersByName(MASTER_FOLDER_NAME);
  const masterFolder = folders.hasNext()
    ? folders.next()
    : DriveApp.createFolder(MASTER_FOLDER_NAME);

  const clientFolders = masterFolder.getFoldersByName(clientName);
  return clientFolders.hasNext()
    ? clientFolders.next()
    : masterFolder.createFolder(clientName);
}

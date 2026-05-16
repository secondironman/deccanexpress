// ============================================================
//  DECCAN EXPRESS — Google Apps Script
//  Paste this entire file into Extensions → Apps Script
//  in your "Deccan Express Orders" Google Sheet
// ============================================================

// ── IMPORTANT: Replace this with your actual Google Sheet ID ──
// Find it in the URL: docs.google.com/spreadsheets/d/YOUR_SHEET_ID/edit
const SHEET_ID = '1oS9qT4zgKh8LHgvii1Xog78R57_S2E1JE0MCk72a7Gw';

// Sheet tab name (the tab at the bottom of your spreadsheet)
const SHEET_NAME = 'Orders';

// ── This function runs when the form POSTs an order ──
function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);
    appendOrder(data);
    return ContentService
      .createTextOutput(JSON.stringify({ status: 'success', ref: data.ref }))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ status: 'error', message: err.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// ── Also handles GET (for testing in browser) ──
function doGet(e) {
  return ContentService
    .createTextOutput(JSON.stringify({ status: 'Deccan Express order endpoint is live ✅' }))
    .setMimeType(ContentService.MimeType.JSON);
}

// ── Writes a row to the sheet ──
function appendOrder(data) {
  const ss    = SpreadsheetApp.openById(SHEET_ID);
  const sheet = ss.getSheetByName(SHEET_NAME);

  // Auto-create headers if sheet is empty
  if (sheet.getLastRow() === 0) {
    sheet.appendRow([
      'Order Ref',
      'Submitted At',
      'Customer Name',
      'Phone',
      'Address',
      'Delivery Date',
      'Time Slot',
      'Biryani Qty',
      'Tea Qty',
      'Cola Qty',
      'Total',
      'Special Notes',
      'Payment Status'
    ]);

    // Style the header row
    const headerRange = sheet.getRange(1, 1, 1, 13);
    headerRange.setBackground('#f0c840');
    headerRange.setFontColor('#1a0800');
    headerRange.setFontWeight('bold');
    headerRange.setFontSize(11);
    sheet.setFrozenRows(1);
  }

  // Append the order row
  sheet.appendRow([
    data.ref          || '',
    data.timestamp    || new Date().toLocaleString(),
    data.name         || '',
    data.phone        || '',
    data.address      || '',
    data.deliveryDate || '',
    data.deliveryTime || '',
    data.biryani      || 0,
    data.tea          || 0,
    data.cola         || 0,
    data.total        || '',
    data.notes        || '',
    'PENDING PAYMENT'   // default — change to PAID once you receive Revolut
  ]);

  // Auto-resize columns for readability
  sheet.autoResizeColumns(1, 13);

  // ── Optional: Send yourself an email notification ──
  // Uncomment and replace with your email to get an email per order:
  //
  // MailApp.sendEmail({
  //   to: 'your@email.com',
  //   subject: '🍛 New Deccan Express Order — ' + data.ref,
  //   body:
  //     'New Pre-Order Received!\n\n' +
  //     'Ref:      ' + data.ref + '\n' +
  //     'Name:     ' + data.name + '\n' +
  //     'Phone:    ' + data.phone + '\n' +
  //     'Address:  ' + data.address + '\n' +
  //     'Date:     ' + data.deliveryDate + '\n' +
  //     'Time:     ' + data.deliveryTime + '\n' +
  //     'Biryani:  ' + data.biryani + '\n' +
  //     'Tea:      ' + data.tea + '\n' +
  //     'Cola:     ' + data.cola + '\n' +
  //     'TOTAL:    ' + data.total + '\n' +
  //     'Notes:    ' + data.notes + '\n\n' +
  //     'Awaiting Revolut payment from customer.'
  // });
}

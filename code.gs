function doGet(e) {
  const type = e.parameter.type || 'invoices';
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(type);

  if (!sheet) {
    return ContentService
      .createTextOutput(JSON.stringify({ status: 'error', message: 'Sheet not found: ' + type }))
      .setMimeType(ContentService.MimeType.JSON);
  }

  const values = sheet.getDataRange().getValues();
  if (values.length < 2) {
    return ContentService.createTextOutput(JSON.stringify([])).setMimeType(ContentService.MimeType.JSON);
  }

  const headers = values[0];
  const rows = values.slice(1).map(row => {
    const obj = {};
    headers.forEach((header, index) => {
      obj[header] = row[index];
    });
    return obj;
  });

  return ContentService.createTextOutput(JSON.stringify(rows)).setMimeType(ContentService.MimeType.JSON);
}

function doPost(e) {
  const payload = JSON.parse(e.postData.contents || '{}');
  const action = payload.action;

  if (action === 'saveCustomer') return saveCustomer(payload.data);
  if (action === 'saveInvoice') return saveInvoice(payload.data);
  if (action === 'deleteCustomer') return deleteCustomer(payload.data);
  if (action === 'deleteInvoice') return deleteInvoice(payload.data);
  if (action === 'saveLog') return saveLog(payload.data);
  if (action === 'clearLogs') return clearLogs();
  if (action === 'saveSettings') return saveSettings(payload.data);
  if (action === 'getSettings') return getSettings();

  return ContentService
    .createTextOutput(JSON.stringify({ status: 'error', message: 'Unknown action' }))
    .setMimeType(ContentService.MimeType.JSON);
}

function getSheet(name, headers) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
  }

  const existing = sheet.getDataRange().getValues();
  if (existing.length === 0) {
    sheet.appendRow(headers);
    return sheet;
  }

  const firstRow = existing[0];
  headers.forEach((header, index) => {
    if (!firstRow[index]) {
      firstRow[index] = header;
    }
  });
  sheet.getRange(1, 1, 1, headers.length).setValues([firstRow]);
  return sheet;
}

function saveCustomer(data) {
  const sheet = getSheet('customers', ['id', 'name', 'company', 'phone', 'email']);
  const rows = sheet.getDataRange().getValues();
  const rowIndex = rows.findIndex(row => row[0] === data.id);

  const newRow = [
    data.id,
    data.name || '',
    data.company || '',
    data.phone || '',
    data.email || ''
  ];

  if (rowIndex !== -1) {
    sheet.getRange(rowIndex + 1, 1, 1, newRow.length).setValues([newRow]);
  } else {
    sheet.appendRow(newRow);
  }

  return ContentService.createTextOutput(JSON.stringify({ status: 'ok' })).setMimeType(ContentService.MimeType.JSON);
}

function saveInvoice(data) {
  const sheet = getSheet('invoices', ['id', 'number', 'customerId', 'amount', 'status', 'issueDate', 'dueDate', 'remindersSent']);
  const rows = sheet.getDataRange().getValues();
  const rowIndex = rows.findIndex(row => row[0] === data.id);

  const reminders = Array.isArray(data.remindersSent) ? data.remindersSent.join('|') : '';

  const newRow = [
    data.id,
    data.number || '',
    data.customerId || '',
    Number(data.amount || 0),
    data.status || 'Belum Lunas',
    data.issueDate || '',
    data.dueDate || '',
    reminders
  ];

  if (rowIndex !== -1) {
    sheet.getRange(rowIndex + 1, 1, 1, newRow.length).setValues([newRow]);
  } else {
    sheet.appendRow(newRow);
  }

  return ContentService.createTextOutput(JSON.stringify({ status: 'ok' })).setMimeType(ContentService.MimeType.JSON);
}

function deleteCustomer(data) {
  const sheet = getSheet('customers', ['id', 'name', 'company', 'phone', 'email']);
  const rows = sheet.getDataRange().getValues();
  const rowIndex = rows.findIndex(row => row[0] === data.id);

  if (rowIndex !== -1) {
    sheet.deleteRow(rowIndex + 1);
  }

  return ContentService.createTextOutput(JSON.stringify({ status: 'ok' })).setMimeType(ContentService.MimeType.JSON);
}

function deleteInvoice(data) {
  const sheet = getSheet('invoices', ['id', 'number', 'customerId', 'amount', 'status', 'issueDate', 'dueDate', 'remindersSent']);
  const rows = sheet.getDataRange().getValues();
  const rowIndex = rows.findIndex(row => row[0] === data.id);

  if (rowIndex !== -1) {
    sheet.deleteRow(rowIndex + 1);
  }

  return ContentService.createTextOutput(JSON.stringify({ status: 'ok' })).setMimeType(ContentService.MimeType.JSON);
}

function saveLog(data) {
  const sheet = getSheet('logs', ['id', 'time', 'channel', 'recipient', 'invoice', 'message', 'status']);
  sheet.appendRow([
    data.id || 'log_' + Date.now(),
    data.time || '',
    data.channel || '',
    data.recipient || '',
    data.invoice || '',
    data.message || '',
    data.status || 'Terkirim'
  ]);

  return ContentService.createTextOutput(JSON.stringify({ status: 'ok' })).setMimeType(ContentService.MimeType.JSON);
}

function clearLogs() {
  const sheet = getSheet('logs', ['id', 'time', 'channel', 'recipient', 'invoice', 'message', 'status']);
  const values = sheet.getDataRange().getValues();

  if (values.length > 1) {
    sheet.getRange(2, 1, values.length - 1, 7).clearContent();
  }

  return ContentService.createTextOutput(JSON.stringify({ status: 'ok' })).setMimeType(ContentService.MimeType.JSON);
}

function saveSettings(data) {
  const sheet = getSheet('settings', ['key', 'value']);
  const rows = sheet.getDataRange().getValues();
  const rowIndex = rows.findIndex(row => row[0] === data.key);

  if (rowIndex !== -1) {
    sheet.getRange(rowIndex + 1, 1, 1, 2).setValues([[data.key, data.value]]);
  } else {
    sheet.appendRow([data.key, data.value]);
  }

  return ContentService.createTextOutput(JSON.stringify({ status: 'ok' })).setMimeType(ContentService.MimeType.JSON);
}

function getSettings() {
  const sheet = getSheet('settings', ['key', 'value']);
  const values = sheet.getDataRange().getValues();
  if (values.length < 2) {
    return ContentService.createTextOutput(JSON.stringify({})).setMimeType(ContentService.MimeType.JSON);
  }

  const data = {};
  values.slice(1).forEach(row => {
    const key = row[0];
    const value = row[1];
    if (key) data[key] = value;
  });

  return ContentService.createTextOutput(JSON.stringify(data)).setMimeType(ContentService.MimeType.JSON);
}

// AutoDashboard — Google Apps Script Backend
// In Google Sheets: Erweiterungen → Apps Script → diesen Code einfügen → deployen
// Deploy als: Web App | Ausführen als: Ich | Zugriff: Jeder
//
// Alle Operationen laufen über GET (vermeidet Browser-CORS-Probleme mit POST-Redirects)
// Aufruf-Muster:
//   Lesen:    ?action=read&sheet=Tanken
//   Hinzuf.:  ?action=add&sheet=Tanken&data={"datum":"2024-01-01",...}
//   Update:   ?action=update&sheet=Tanken&row=5&data={...}
//   Löschen:  ?action=delete&sheet=Tanken&row=5

const SPREADSHEET_ID = '1RexNif9j0xDXS0mV1YT6JvG0fAhWXHUGMAXB0sMRSxA';

function doGet(e) {
  try {
    const action = e.parameter.action || 'read';
    const sheet  = e.parameter.sheet;

    if (!sheet) return respond({ error: 'Parameter "sheet" fehlt' });

    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const ws = ss.getSheetByName(sheet);
    if (!ws) return respond({ error: `Sheet "${sheet}" nicht gefunden` });

    if (action === 'read') {
      return respond({ success: true, data: readSheet(ws) });
    }

    if (action === 'add') {
      const data = e.parameter.data ? JSON.parse(e.parameter.data) : {};
      const row  = appendRow(ws, data);
      return respond({ success: true, row });
    }

    if (action === 'batchAdd') {
      const rows = e.parameter.data ? JSON.parse(e.parameter.data) : [];
      if (!Array.isArray(rows)) return respond({ error: '"data" muss ein Array sein' });
      initHeaders(ws, rows[0]);
      const lastCol = ws.getLastColumn();
      const headers = ws.getRange(1, 1, 1, lastCol).getValues()[0];
      const matrix  = rows.map(d => headers.map(h => (d[h] !== undefined ? d[h] : '')));
      ws.getRange(ws.getLastRow() + 1, 1, matrix.length, headers.length).setValues(matrix);
      return respond({ success: true, count: rows.length });
    }

    if (action === 'update') {
      const row  = parseInt(e.parameter.row);
      const data = e.parameter.data ? JSON.parse(e.parameter.data) : {};
      if (!row) return respond({ error: 'Parameter "row" fehlt' });
      updateRow(ws, row, data);
      return respond({ success: true });
    }

    if (action === 'delete') {
      const row = parseInt(e.parameter.row);
      if (!row) return respond({ error: 'Parameter "row" fehlt' });
      ws.deleteRow(row);
      return respond({ success: true });
    }

    return respond({ error: `Unbekannte Aktion: ${action}` });

  } catch (err) {
    return respond({ error: err.message });
  }
}

// ---------------------------------------------------------------------------

function isDate(v) {
  return Object.prototype.toString.call(v) === '[object Date]';
}

function readSheet(ws) {
  const lastRow = ws.getLastRow();
  const lastCol = ws.getLastColumn();
  if (lastRow < 2 || lastCol < 1) return [];

  const headers = ws.getRange(1, 1, 1, lastCol).getValues()[0];
  const rows    = ws.getRange(2, 1, lastRow - 1, lastCol).getValues();
  const tz      = Session.getScriptTimeZone();

  return rows.map((row, i) => {
    const obj = { _row: i + 2 };
    headers.forEach((h, j) => {
      obj[h] = isDate(row[j])
        ? Utilities.formatDate(row[j], tz, 'yyyy-MM-dd')
        : row[j];
    });
    return obj;
  });
}

function initHeaders(ws, sample) {
  if (ws.getLastRow() === 0 && sample) {
    ws.appendRow(Object.keys(sample));
  }
}

function appendRow(ws, data) {
  initHeaders(ws, data);
  const lastCol = ws.getLastColumn();
  const headers = ws.getRange(1, 1, 1, lastCol).getValues()[0];
  const row     = headers.map(h => (data[h] !== undefined ? data[h] : ''));
  ws.appendRow(row);
  return ws.getLastRow();
}

function updateRow(ws, rowNumber, data) {
  const lastCol = ws.getLastColumn();
  const headers = ws.getRange(1, 1, 1, lastCol).getValues()[0];
  const row     = headers.map(h => (data[h] !== undefined ? data[h] : ''));
  ws.getRange(rowNumber, 1, 1, row.length).setValues([row]);
}

function respond(data) {
  return ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}

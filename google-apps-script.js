// ============================================================
//  GOOGLE APPS SCRIPT - AutoVenda Pro
//  IMPORTANTE: Ao implantar selecione:
//  - Executar como: Eu
//  - Quem tem acesso: Qualquer pessoa (até anônimos)
// ============================================================

const SHEET_NAME = 'Clientes';
const EVOLUTION_URL = 'https://evolution-api-production-da04e.up.railway.app';
const EVOLUTION_INSTANCE = 'autopecas';
const INSTANCE_TOKEN = '7450F3DF72D9-4FC3-B24B-1AEED19657CC';

function getSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
    const headers = ['id','empresa','cnpj','contato','telefone','email','whatsapp','obs','lastContact','lastPurchase','purchaseCount','createdAt','history'];
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold');
    sheet.setFrozenRows(1);
  }
  return sheet;
}

function createResponse(data) {
  const output = ContentService.createTextOutput(JSON.stringify(data));
  output.setMimeType(ContentService.MimeType.JSON);
  return output;
}

function doGet(e) {
  try {
    const action = e.parameter ? e.parameter.action : null;

    // Proxy WhatsApp
    if (action === 'sendWhatsApp') {
      const number = (e.parameter.number || '').replace(/\D/g, '');
      const message = decodeURIComponent(e.parameter.message || '');

      if (!number || !message) {
        return createResponse({ status: 'error', message: 'Parametros invalidos' });
      }

      const url = `${EVOLUTION_URL}/message/sendText/${EVOLUTION_INSTANCE}`;
      const options = {
        method: 'post',
        contentType: 'application/json',
        headers: { 'apikey': INSTANCE_TOKEN },
        payload: JSON.stringify({
          number: number + '@s.whatsapp.net',
          text: message,
          delay: 1000
        }),
        muteHttpExceptions: true
      };

      const response = UrlFetchApp.fetch(url, options);
      const code = response.getResponseCode();

      if (code === 200 || code === 201) {
        return createResponse({ status: 'ok', code: code });
      } else {
        return createResponse({ status: 'error', code: code, body: response.getContentText() });
      }
    }

    // Buscar clientes
    const sheet = getSheet();
    const data = sheet.getDataRange().getValues();
    const headers = data[0];
    const clients = [];

    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      if (!row[0]) continue;
      const client = {};
      headers.forEach((h, j) => { client[h] = row[j]; });
      try { client.history = JSON.parse(client.history || '[]'); } catch(e) { client.history = []; }
      clients.push(client);
    }

    return createResponse({ status: 'ok', clients: clients });

  } catch(err) {
    return createResponse({ status: 'error', message: err.toString() });
  }
}

function doPost(e) {
  try {
    let clients = [];
    // Aceita text/plain (enviado pelo browser para evitar CORS preflight)
    if (e.postData && e.postData.contents) {
      const data = JSON.parse(e.postData.contents);
      if (data.action === 'save') {
        clients = data.clients || [];
      }
    } else if (e.parameters && e.parameters.clients) {
      clients = JSON.parse(e.parameters.clients[0]);
    }
    saveClients(clients);
    return createResponse({ status: 'ok', saved: clients.length });
  } catch(err) {
    return createResponse({ status: 'error', message: err.toString() });
  }
}

function saveClients(clients) {
  const sheet = getSheet();
  const lastRow = sheet.getLastRow();
  if (lastRow > 1) {
    sheet.getRange(2, 1, lastRow - 1, 13).clearContent();
  }
  if (!clients || clients.length === 0) return;
  const rows = clients.map(c => [
    c.id || '', c.empresa || '', c.cnpj || '', c.contato || '',
    c.telefone || '', c.email || '', c.whatsapp || '', c.obs || '',
    c.lastContact || '', c.lastPurchase || '', c.purchaseCount || 0,
    c.createdAt || '', JSON.stringify(c.history || [])
  ]);
  sheet.getRange(2, 1, rows.length, 13).setValues(rows);
}

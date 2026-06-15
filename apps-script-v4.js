// ============================================
// Apps Script — Apontamento de Apara por OP
// v4 — OPs_v2: estrutura dinâmica
// ============================================
// Cole ABAIXO do código existente no editor,
// ou substitua tudo se quiser um script limpo.
//
// Para atualizar uma implantação existente:
// Implantar > Gerenciar implantações > lápis
// > Nova versão > Implantar

var SHEET_V2 = 'OPs_v2';

var HEADERS_V2 = [
  'timestamp',
  'data',
  'cliente',
  'produto',
  'codigo',
  'tipo_op',        // 'box_pouch' | 'simples'
  'maquinas',       // string separada por vírgulas
  'qtd_unidades',
  'metragem_principal',
  'total_util',
  'total_apara',
  'total_pct',
  'total_pct_sr',
  'etapas_json',    // JSON com todas as etapas detalhadas
  'meta_json'       // dimensões refile formatação + inputs_json
];

// ─── POST ───
function doPost(e) {
  try {
    var payload = JSON.parse(e.postData.contents);
    var version = payload._version || 'v1';

    if (version === 'v2') {
      return handleV2Post(payload);
    } else {
      // fallback para código v1/v3 existente
      return handleV1Post(payload);
    }
  } catch (err) {
    return respond({ status: 'error', message: err.toString() });
  }
}

// ─── GET ───
function doGet(e) {
  try {
    var version = e && e.parameter && e.parameter.v ? e.parameter.v : 'v1';
    if (version === 'v2') {
      return handleV2Get(e);
    } else {
      return handleV1Get(e);
    }
  } catch (err) {
    return respond({ status: 'error', message: err.toString() });
  }
}

// ══════════════════════════════════════════
// V2 handlers
// ══════════════════════════════════════════

function handleV2Post(payload) {
  var sheet = getOrCreateSheetV2();
  var action = payload._action || 'save';

  var row = [];
  for (var i = 0; i < HEADERS_V2.length; i++) {
    var h = HEADERS_V2[i];
    if (h === 'timestamp') {
      if (action === 'update' && payload._row) {
        var orig = sheet.getRange(payload._row, 1).getValue();
        row.push(orig || new Date().toISOString());
      } else {
        row.push(new Date().toISOString());
      }
    } else if (h === 'etapas_json') {
      row.push(payload.etapas_json || '[]');
    } else if (h === 'meta_json') {
      row.push(payload.meta_json || '{}');
    } else {
      row.push(payload[h] !== undefined ? payload[h] : '');
    }
  }

  if (action === 'update' && payload._row && payload._row >= 2) {
    sheet.getRange(payload._row, 1, 1, row.length).setValues([row]);
    return respond({ status: 'ok', message: 'OP atualizada com sucesso' });
  } else {
    sheet.appendRow(row);
    return respond({ status: 'ok', message: 'OP salva com sucesso' });
  }
}

function handleV2Get(e) {
  var sheet = getOrCreateSheetV2();
  var lastRow = sheet.getLastRow();

  if (lastRow <= 1) {
    return respond({ status: 'ok', ops: [] });
  }

  var numCols = sheet.getLastColumn();
  var data = sheet.getRange(2, 1, lastRow - 1, numCols).getValues();
  var headers = sheet.getRange(1, 1, 1, numCols).getValues()[0];

  var ops = [];
  for (var i = 0; i < data.length; i++) {
    var obj = { _row: i + 2 };
    for (var j = 0; j < headers.length; j++) {
      var val = data[i][j];
      if (val instanceof Date) {
        val = Utilities.formatDate(val, Session.getScriptTimeZone(), 'yyyy-MM-dd');
      }
      obj[String(headers[j])] = val;
    }
    ops.push(obj);
  }

  return respond({ status: 'ok', ops: ops });
}

function getOrCreateSheetV2() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_V2);

  if (!sheet) {
    sheet = ss.insertSheet(SHEET_V2);
    sheet.getRange(1, 1, 1, HEADERS_V2.length).setValues([HEADERS_V2]);
    sheet.getRange(1, 1, 1, HEADERS_V2.length).setFontWeight('bold');
    sheet.setFrozenRows(1);
    // Largura das colunas JSON
    sheet.setColumnWidth(HEADERS_V2.indexOf('etapas_json') + 1, 400);
    sheet.setColumnWidth(HEADERS_V2.indexOf('meta_json') + 1, 300);
  } else {
    // Migração: adicionar colunas faltantes
    var lastCol = sheet.getLastColumn();
    var existing = lastCol > 0
      ? sheet.getRange(1, 1, 1, lastCol).getValues()[0].map(String)
      : [];
    for (var i = 0; i < HEADERS_V2.length; i++) {
      if (existing.indexOf(HEADERS_V2[i]) === -1) {
        sheet.getRange(1, lastCol + 1).setValue(HEADERS_V2[i]);
        sheet.getRange(1, lastCol + 1).setFontWeight('bold');
        lastCol++;
      }
    }
  }

  return sheet;
}

// ══════════════════════════════════════════
// V1 handlers (mantidos para compatibilidade)
// ══════════════════════════════════════════

var SHEET_NAME = 'OPs';

var HEADERS_V1 = [
  'timestamp', 'data', 'cliente', 'produto', 'codigo',
  'f_imp_apara', 'f_imp_pct', 'f_l1_apara', 'f_l1_pct',
  'f_l2_apara', 'f_l2_pct', 'f_reb_apara', 'f_reb_pct_cr',
  'f_reb_pct_sr', 'f_reb_refile', 'f_util_geral', 'f_apara_total', 'f_pct_geral',
  's_imp_apara', 's_imp_pct', 's_l1_apara', 's_l1_pct',
  's_l2_apara', 's_l2_pct', 's_reb_apara', 's_reb_pct_cr',
  's_reb_pct_sr', 's_reb_refile', 's_util_geral', 's_apara_total', 's_pct_geral',
  'fmt_qtd', 'fmt_apara', 'fmt_pct', 'fmt_refile_kg',
  'total_util', 'total_apara', 'total_pct',
  'inputs_json'
];

function handleV1Post(payload) {
  var sheet = getOrCreateSheetV1();
  var action = payload._action || 'save';

  var row = [];
  for (var i = 0; i < HEADERS_V1.length; i++) {
    var h = HEADERS_V1[i];
    if (h === 'timestamp') {
      if (action === 'update' && payload._row) {
        var orig = sheet.getRange(payload._row, 1).getValue();
        row.push(orig || new Date().toISOString());
      } else {
        row.push(new Date().toISOString());
      }
    } else if (h === 'inputs_json') {
      row.push(payload._inputs_json || '{}');
    } else {
      row.push(payload[h] !== undefined ? payload[h] : '');
    }
  }

  if (action === 'update' && payload._row && payload._row >= 2) {
    sheet.getRange(payload._row, 1, 1, row.length).setValues([row]);
    return respond({ status: 'ok', message: 'OP atualizada com sucesso' });
  } else {
    sheet.appendRow(row);
    return respond({ status: 'ok', message: 'OP salva com sucesso' });
  }
}

function handleV1Get(e) {
  var sheet = getOrCreateSheetV1();
  var lastRow = sheet.getLastRow();

  if (lastRow <= 1) {
    return respond({ status: 'ok', ops: [] });
  }

  var numCols = sheet.getLastColumn();
  var data = sheet.getRange(2, 1, lastRow - 1, numCols).getValues();
  var headers = sheet.getRange(1, 1, 1, numCols).getValues()[0];

  var ops = [];
  for (var i = 0; i < data.length; i++) {
    var obj = { _row: i + 2 };
    for (var j = 0; j < headers.length; j++) {
      var val = data[i][j];
      if (val instanceof Date) {
        val = Utilities.formatDate(val, Session.getScriptTimeZone(), 'yyyy-MM-dd');
      }
      obj[String(headers[j])] = val;
    }
    ops.push(obj);
  }

  return respond({ status: 'ok', ops: ops });
}

function getOrCreateSheetV1() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_NAME);

  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
    sheet.getRange(1, 1, 1, HEADERS_V1.length).setValues([HEADERS_V1]);
    sheet.getRange(1, 1, 1, HEADERS_V1.length).setFontWeight('bold');
    sheet.setFrozenRows(1);
  } else {
    var lastCol = sheet.getLastColumn();
    var existing = lastCol > 0
      ? sheet.getRange(1, 1, 1, lastCol).getValues()[0].map(String)
      : [];
    for (var i = 0; i < HEADERS_V1.length; i++) {
      if (existing.indexOf(HEADERS_V1[i]) === -1) {
        sheet.getRange(1, lastCol + 1).setValue(HEADERS_V1[i]);
        sheet.getRange(1, lastCol + 1).setFontWeight('bold');
        lastCol++;
      }
    }
  }

  return sheet;
}

// ─── Utilitário ───
function respond(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

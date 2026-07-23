// ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
//  LUMENGRID — Financeiro · Google Apps Script Backend
//  Planilha: https://docs.google.com/spreadsheets/d/1b5bp7uPF2jDsR2i9CvGVdyhfkyuqYm-FyKevn3o8foM
//  Cole em: script.google.com → Novo Projeto → Colar → Implantar como Web App
// ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

const SHEET_ID_FIN = '1b5bp7uPF2jDsR2i9CvGVdyhfkyuqYm-FyKevn3o8foM';

// ── Cabeçalhos ────────────────────────────────────────────────
const HDR_PROJETOS = [
  'Nº Contrato','Cliente','Valor Venda (R$)','Data','Vendedor','Parceiro','Observação','Criado em'
];
const HDR_DESPESAS = [
  'ID','Tipo','Data','Nº Contrato','Categoria','Descrição','Valor (R$)','Lançado por','Criado em'
];

// ── CORS ──────────────────────────────────────────────────────
function addCors(output) {
  return output
    .setHeader('Access-Control-Allow-Origin', '*')
    .setHeader('Access-Control-Allow-Methods', 'GET,POST')
    .setHeader('Access-Control-Allow-Headers', 'Content-Type');
}

// ── GET ───────────────────────────────────────────────────────
function doGet(e) {
  const action = (e && e.parameter && e.parameter.action) || 'get_data';
  let result;
  try {
    if (action === 'get_data')     result = getData();
    else if (action === 'get_projetos') result = { projetos: getProjetos() };
    else result = { error: 'Ação desconhecida: ' + action };
  } catch(err) {
    result = { error: err.message };
  }
  return addCors(ContentService
    .createTextOutput(JSON.stringify(result))
    .setMimeType(ContentService.MimeType.JSON));
}

// ── POST ──────────────────────────────────────────────────────
function doPost(e) {
  let result;
  try {
    const body   = JSON.parse(e.postData.contents);
    const action = body.action;
    if      (action === 'save_despesa') result = saveDespesa(body);
    else if (action === 'save_receita') result = saveDespesa(body); // mesmo fluxo
    else if (action === 'save_projeto') result = saveProjeto(body);
    else if (action === 'vincular')     result = vincularContrato(body);
    else result = { error: 'Ação desconhecida: ' + action };
  } catch(err) {
    result = { error: err.message };
  }
  return addCors(ContentService
    .createTextOutput(JSON.stringify(result))
    .setMimeType(ContentService.MimeType.JSON));
}

// ── GET DATA (tudo) ───────────────────────────────────────────
function getData() {
  return {
    projetos:  getProjetos(),
    despesas:  getDespesas(),
  };
}

// ── GARANTIR ABA ──────────────────────────────────────────────
function garantirAba(ss, nome, headers) {
  let sheet = ss.getSheetByName(nome);
  if (!sheet) {
    sheet = ss.insertSheet(nome);
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    sheet.getRange(1, 1, 1, headers.length)
      .setBackground('#E8641A')
      .setFontColor('#FFFFFF')
      .setFontWeight('bold');
    sheet.setFrozenRows(1);
  }
  return sheet;
}

// ── PROJETOS ──────────────────────────────────────────────────
function getProjetos() {
  const ss    = SpreadsheetApp.openById(SHEET_ID_FIN);
  const sheet = ss.getSheetByName('Projetos');
  if (!sheet || sheet.getLastRow() < 2) return [];

  const data    = sheet.getDataRange().getValues();
  const headers = data[0];
  return data.slice(1).filter(r => r[0]).map(row => {
    const obj = {};
    headers.forEach((h, i) => obj[h] = row[i]);
    return {
      contrato:  obj['Nº Contrato']    || '',
      cliente:   obj['Cliente']         || '',
      valorVenda:parseFloat(obj['Valor Venda (R$)']) || 0,
      data:      obj['Data']            || '',
      vendedor:  obj['Vendedor']        || '',
      parceiro:  obj['Parceiro']        || '',
      obs:       obj['Observação']      || '',
    };
  });
}

function saveProjeto(body) {
  const ss    = SpreadsheetApp.openById(SHEET_ID_FIN);
  const sheet = garantirAba(ss, 'Projetos', HDR_PROJETOS);

  // Verifica duplicata
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]).trim() === String(body.contrato || '').trim()) {
      // Atualiza valor se já existe
      sheet.getRange(i + 1, 3).setValue(body.valorVenda || 0);
      return { ok: true, updated: true };
    }
  }

  sheet.appendRow([
    body.contrato   || '',
    body.cliente    || '',
    body.valorVenda || 0,
    body.data       || new Date().toLocaleDateString('pt-BR'),
    body.vendedor   || '',
    body.parceiro   || '',
    body.obs        || '',
    new Date().toLocaleString('pt-BR'),
  ]);
  return { ok: true, created: true };
}

// ── DESPESAS ──────────────────────────────────────────────────
function getDespesas() {
  const ss    = SpreadsheetApp.openById(SHEET_ID_FIN);
  const sheet = ss.getSheetByName('Despesas');
  if (!sheet || sheet.getLastRow() < 2) return [];

  const data    = sheet.getDataRange().getValues();
  const headers = data[0];
  return data.slice(1).filter(r => r[0]).map(row => {
    const obj = {};
    headers.forEach((h, i) => obj[h] = row[i]);
    return {
      id:         String(obj['ID']           || ''),
      tipo:       obj['Tipo']                 || 'despesa',
      data:       obj['Data']                 || '',
      contrato:   String(obj['Nº Contrato']  || ''),
      categoria:  obj['Categoria']            || '',
      descricao:  obj['Descrição']            || '',
      valor:      parseFloat(obj['Valor (R$)']) || 0,
      lancadoPor: obj['Lançado por']          || '',
      criadoEm:   obj['Criado em']            || '',
    };
  }).reverse(); // mais recentes primeiro
}

function saveDespesa(body) {
  const ss    = SpreadsheetApp.openById(SHEET_ID_FIN);
  const sheet = garantirAba(ss, 'Despesas', HDR_DESPESAS);

  sheet.appendRow([
    body.id          || Utilities.getUuid(),
    body.tipo        || 'despesa',
    body.data        || new Date().toLocaleDateString('pt-BR'),
    body.contrato    || '',
    body.categoria   || '',
    body.descricao   || '',
    parseFloat(body.valor) || 0,
    body.lancadoPor  || '',
    new Date().toLocaleString('pt-BR'),
  ]);
  return { ok: true };
}

// ── VINCULAR CONTRATO ─────────────────────────────────────────
function vincularContrato(body) {
  const ss    = SpreadsheetApp.openById(SHEET_ID_FIN);
  const sheet = ss.getSheetByName('Despesas');
  if (!sheet) return { error: 'Aba Despesas não encontrada' };

  const data     = sheet.getDataRange().getValues();
  const headers  = data[0];
  const idCol    = headers.indexOf('ID');
  const contCol  = headers.indexOf('Nº Contrato') + 1;

  for (let i = 1; i < data.length; i++) {
    if (String(data[i][idCol]) === String(body.id)) {
      sheet.getRange(i + 1, contCol).setValue(body.contrato || '');
      return { ok: true };
    }
  }
  return { error: 'Despesa não encontrada: ' + body.id };
}

// ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
// COMO IMPLANTAR:
// 1. Acesse script.google.com → Novo projeto
// 2. Cole todo este código
// 3. Clique em Implantar → Nova implantação
// 4. Tipo: App da Web
// 5. Executar como: Você (sua conta Google)
// 6. Quem tem acesso: Qualquer pessoa
// 7. Clique em Implantar → copie a URL
// 8. Cole a URL em financeiro.html → aba Config
// ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

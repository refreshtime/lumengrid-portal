// ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
//  LUMENGRID — Financeiro · Google Apps Script Backend
//  Planilha: https://docs.google.com/spreadsheets/d/1b5bp7uPF2jDsR2i9CvGVdyhfkyuqYm-FyKevn3o8foM
//  Cole em: script.google.com → Novo Projeto → Colar → Implantar como Web App
//  Executar como: Você · Quem acessa: Qualquer pessoa
// ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

const SHEET_ID_FIN = '1b5bp7uPF2jDsR2i9CvGVdyhfkyuqYm-FyKevn3o8foM';

// ── Cabeçalhos (correspondem exatamente ao dashboard.html) ────
const HDR_RECEITAS = [
  'ID','Data','Cliente','CPF/CNPJ','Endereço','Tipo','kVp','Módulos',
  'Inversor','Bateria','Valor Total (R$)','Valor Recebido (R$)',
  'Forma Pagamento','Condições','Consultor','Status','Observações','Nº Contrato','Criado em'
];
const HDR_DESPESAS = [
  'ID','Data','Descrição','Categoria','Valor (R$)','Método',
  'Observação','Nº Contrato','Projeto ID','Criado em'
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
    if (action === 'get_data') result = getData();
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
    else if (action === 'save_receita') result = saveReceita(body);
    else if (action === 'vincular')     result = vincularDespesa(body);
    else result = { error: 'Ação desconhecida: ' + action };
  } catch(err) {
    result = { error: err.message };
  }
  return addCors(ContentService
    .createTextOutput(JSON.stringify(result))
    .setMimeType(ContentService.MimeType.JSON));
}

// ── GET DATA ──────────────────────────────────────────────────
function getData() {
  return {
    receitas: getReceitas(),
    despesas: getDespesas(),
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

// ── RECEITAS ──────────────────────────────────────────────────
function getReceitas() {
  const ss    = SpreadsheetApp.openById(SHEET_ID_FIN);
  const sheet = ss.getSheetByName('Receitas');
  if (!sheet || sheet.getLastRow() < 2) return [];

  const data    = sheet.getDataRange().getValues();
  const headers = data[0];
  return data.slice(1).filter(r => r[0]).map(row => {
    const o = {};
    headers.forEach((h, i) => o[h] = row[i]);
    return {
      id:         String(o['ID']                  || ''),
      data:       fmtIsoDate(o['Data']            || ''),
      cliente:    o['Cliente']                     || '',
      doc:        o['CPF/CNPJ']                    || '',
      ender:      o['Endereço']                    || '',
      tipo:       o['Tipo']                        || '',
      kvp:        o['kVp']                         || '',
      modulos:    o['Módulos']                     || '',
      inversor:   o['Inversor']                    || '',
      bateria:    o['Bateria']                     || '',
      valor:      parseFloat(o['Valor Total (R$)'])   || 0,
      recebido:   parseFloat(o['Valor Recebido (R$)']) || 0,
      pagamento:  o['Forma Pagamento']             || '',
      condicoes:  o['Condições']                   || '',
      consultor:  o['Consultor']                   || '',
      status:     o['Status']                      || '',
      obs:        o['Observações']                 || '',
      numContrato:String(o['Nº Contrato']          || ''),
    };
  });
}

function saveReceita(body) {
  const ss    = SpreadsheetApp.openById(SHEET_ID_FIN);
  const sheet = garantirAba(ss, 'Receitas', HDR_RECEITAS);

  // Evita duplicata por ID
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]) === String(body.id || '')) {
      return { ok: true, skipped: true };
    }
  }

  sheet.appendRow([
    body.id          || Utilities.getUuid(),
    body.data        || '',
    body.cliente     || '',
    body.doc         || '',
    body.ender       || '',
    body.tipo        || '',
    body.kvp         || '',
    body.modulos     || '',
    body.inversor    || '',
    body.bateria     || '',
    parseFloat(body.valor)     || 0,
    parseFloat(body.recebido)  || 0,
    body.pagamento   || '',
    body.condicoes   || '',
    body.consultor   || '',
    body.status      || '',
    body.obs         || '',
    body.numContrato || '',
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
    const o = {};
    headers.forEach((h, i) => o[h] = row[i]);
    return {
      id:          String(o['ID']           || ''),
      data:        fmtIsoDate(o['Data']   || ''),
      desc:        o['Descrição']          || '',
      cat:         o['Categoria']          || '',
      valor:       parseFloat(o['Valor (R$)']) || 0,
      metodo:      o['Método']             || '',
      obs:         o['Observação']         || '',
      numContrato: String(o['Nº Contrato']|| ''),
      projetoId:   String(o['Projeto ID'] || ''),
    };
  }).reverse();
}

function saveDespesa(body) {
  const ss    = SpreadsheetApp.openById(SHEET_ID_FIN);
  const sheet = garantirAba(ss, 'Despesas', HDR_DESPESAS);

  // Evita duplicata por ID
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]) === String(body.id || '')) {
      return { ok: true, skipped: true };
    }
  }

  sheet.appendRow([
    body.id        || Utilities.getUuid(),
    body.data      || '',
    body.desc      || '',
    body.cat       || '',
    parseFloat(body.valor) || 0,
    body.metodo    || '',
    body.obs         || '',
    body.numContrato || '',
    body.projetoId   || '',
    new Date().toLocaleString('pt-BR'),
  ]);
  return { ok: true, created: true };
}

// ── VINCULAR DESPESA A PROJETO ────────────────────────────────
function vincularDespesa(body) {
  const ss    = SpreadsheetApp.openById(SHEET_ID_FIN);
  const sheet = ss.getSheetByName('Despesas');
  if (!sheet) return { error: 'Aba Despesas não encontrada' };

  const data    = sheet.getDataRange().getValues();
  const headers = data[0];
  const idCol   = headers.indexOf('ID');
  const ncCol   = headers.indexOf('Nº Contrato') + 1;
  const pidCol  = headers.indexOf('Projeto ID') + 1;

  for (let i = 1; i < data.length; i++) {
    if (String(data[i][idCol]) === String(body.id || '')) {
      if (ncCol  > 0) sheet.getRange(i + 1, ncCol).setValue(body.numContrato || '');
      if (pidCol > 0) sheet.getRange(i + 1, pidCol).setValue(body.projetoId  || '');
      return { ok: true };
    }
  }
  return { error: 'Despesa não encontrada: ' + body.id };
}

// ── HELPER: converte Date objeto ou string p/ YYYY-MM-DD ──────
function fmtIsoDate(v) {
  if (!v) return '';
  if (v instanceof Date) return v.toISOString().slice(0, 10);
  return String(v).slice(0, 10);
}

// ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
// COMO IMPLANTAR:
// 1. Acesse script.google.com → Novo projeto
// 2. Cole todo este código
// 3. Clique em Implantar → Nova implantação
// 4. Tipo: App da Web
// 5. Executar como: Você (sua conta Google)
// 6. Quem tem acesso: Qualquer pessoa
// 7. Clique em Implantar → copie a URL gerada
// 8. No dashboard.html → botão engrenagem (Config) → cole a URL → Salvar
// ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

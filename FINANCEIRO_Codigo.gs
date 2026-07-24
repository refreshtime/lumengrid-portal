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
  'Observação','Lançado por','Nº Contrato','Projeto ID','Criado em'
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

  atualizarResumos();
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
      obs:         o['Observação']          || '',
      lancadoPor:  o['Lançado por']        || '',
      numContrato: String(o['Nº Contrato'] || ''),
      projetoId:   String(o['Projeto ID']  || ''),
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

  atualizarResumos();
  sheet.appendRow([
    body.id        || Utilities.getUuid(),
    body.data      || '',
    body.desc      || '',
    body.cat       || '',
    parseFloat(body.valor) || 0,
    body.metodo    || '',
    body.obs         || '',
    body.lancadoPor  || '',
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

// ── RESUMOS AUTOMÁTICOS ───────────────────────────────────────
function atualizarResumos() {
  try {
    const ss       = SpreadsheetApp.openById(SHEET_ID_FIN);
    const recs     = getReceitas();
    const desps    = getDespesas();
    gerarCompetencia(ss, recs, desps);
    gerarFluxoCaixa(ss, recs, desps);
  } catch(e) {
    Logger.log('atualizarResumos erro: ' + e.message);
  }
}

function gerarCompetencia(ss, recs, desps) {
  let sheet = ss.getSheetByName('📊 Competência');
  if (!sheet) { sheet = ss.insertSheet('📊 Competência'); }
  sheet.clearContents();

  const hdrs = ['Mês','Receitas (R$)','Despesas (R$)','Resultado (R$)','Margem %','Acumulado (R$)'];
  const hdrRow = sheet.getRange(1, 1, 1, hdrs.length);
  hdrRow.setValues([hdrs]).setBackground('#1a7a3c').setFontColor('#fff').setFontWeight('bold');
  sheet.setFrozenRows(1);

  const meses = [...new Set([
    ...recs.map(r => (r.data||'').slice(0,7)),
    ...desps.map(d => (d.data||'').slice(0,7))
  ])].filter(Boolean).sort();

  let acum = 0;
  const rows = meses.map(m => {
    const rec  = recs .filter(r => (r.data||'').startsWith(m)).reduce((s,r) => s+r.valor, 0);
    const desp = desps.filter(d => (d.data||'').startsWith(m)).reduce((s,d) => s+d.valor, 0);
    const res  = rec - desp;
    const mg   = rec > 0 ? (res / rec * 100).toFixed(1) + '%' : '—';
    acum += res;
    return [nomeMes(m), rec, desp, res, mg, acum];
  });

  // Linha de total
  const totRec  = recs .reduce((s,r) => s+r.valor, 0);
  const totDesp = desps.reduce((s,d) => s+d.valor, 0);
  rows.push(['TOTAL', totRec, totDesp, totRec-totDesp, totRec>0?((totRec-totDesp)/totRec*100).toFixed(1)+'%':'—', '']);

  if (rows.length > 0) {
    sheet.getRange(2, 1, rows.length, hdrs.length).setValues(rows);
    // Formata valores numéricos
    const fmt = '#,##0.00';
    [[2,2],[2,3],[2,4],[2,6]].forEach(([c1,c2]) =>
      sheet.getRange(2, c1, rows.length, c2-c1+1)
           .setNumberFormat('R$ ' + fmt));
    // Linha de total em negrito
    sheet.getRange(rows.length+1, 1, 1, hdrs.length).setFontWeight('bold').setBackground('#f0f0f0');
    // Resultado: verde se positivo, vermelho se negativo
    rows.forEach((r, i) => {
      const cell = sheet.getRange(i+2, 4);
      if (typeof r[3] === 'number') cell.setFontColor(r[3] >= 0 ? '#1a7a3c' : '#c0392b');
    });
  }
  sheet.autoResizeColumns(1, hdrs.length);
}

function gerarFluxoCaixa(ss, recs, desps) {
  let sheet = ss.getSheetByName('💰 Fluxo de Caixa');
  if (!sheet) { sheet = ss.insertSheet('💰 Fluxo de Caixa'); }
  sheet.clearContents();

  const hdrs = ['Mês','Entradas Recebidas (R$)','Receita a Receber (R$)','Saídas (R$)','Saldo do Mês (R$)','Saldo Acumulado (R$)'];
  sheet.getRange(1, 1, 1, hdrs.length)
       .setValues([hdrs]).setBackground('#1a4a7a').setFontColor('#fff').setFontWeight('bold');
  sheet.setFrozenRows(1);

  const meses = [...new Set([
    ...recs .map(r => (r.data||'').slice(0,7)),
    ...desps.map(d => (d.data||'').slice(0,7))
  ])].filter(Boolean).sort();

  let acum = 0;
  const rows = meses.map(m => {
    const recebido   = recs .filter(r => (r.data||'').startsWith(m)).reduce((s,r) => s+(r.recebido||0), 0);
    const aReceber   = recs .filter(r => (r.data||'').startsWith(m)).reduce((s,r) => s+Math.max(0,(r.valor||0)-(r.recebido||0)), 0);
    const saidas     = desps.filter(d => (d.data||'').startsWith(m)).reduce((s,d) => s+d.valor, 0);
    const saldo      = recebido - saidas;
    acum += saldo;
    return [nomeMes(m), recebido, aReceber, saidas, saldo, acum];
  });

  const totRec = recs .reduce((s,r) => s+(r.recebido||0), 0);
  const totAR  = recs .reduce((s,r) => s+Math.max(0,(r.valor||0)-(r.recebido||0)), 0);
  const totSai = desps.reduce((s,d) => s+d.valor, 0);
  rows.push(['TOTAL', totRec, totAR, totSai, totRec-totSai, '']);

  if (rows.length > 0) {
    sheet.getRange(2, 1, rows.length, hdrs.length).setValues(rows);
    const fmt = '#,##0.00';
    sheet.getRange(2, 2, rows.length, 5).setNumberFormat('R$ ' + fmt);
    sheet.getRange(rows.length+1, 1, 1, hdrs.length).setFontWeight('bold').setBackground('#f0f0f0');
    rows.forEach((r, i) => {
      const cell = sheet.getRange(i+2, 5);
      if (typeof r[4] === 'number') cell.setFontColor(r[4] >= 0 ? '#1a7a3c' : '#c0392b');
    });
  }
  sheet.autoResizeColumns(1, hdrs.length);
}

function nomeMes(m) {
  if (!m) return '—';
  const [y, mo] = m.split('-');
  const ns = ['Jan','Fev','Mar','Abr','Mai','Jun','Jul','Ago','Set','Out','Nov','Dez'];
  return (ns[+mo-1]||mo) + '/' + y;
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

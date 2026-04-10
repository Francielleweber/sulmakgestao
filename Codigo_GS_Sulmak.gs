// ═══════════════════════════════════════════════════════
//  SULMAK LOCAÇÕES — Google Apps Script
//  Cole este código no Apps Script da sua planilha
//  (Extensions > Apps Script)
// ═══════════════════════════════════════════════════════

const SHEET_ID = SpreadsheetApp.getActiveSpreadsheet().getId();

// Nomes das abas — altere se necessário
const ABA_FUNCIONARIOS = 'Funcionarios';
const ABA_LANCAMENTOS  = 'Lancamentos';
const ABA_FERIAS       = 'Ferias';

// ── Cabeçalhos das abas ──────────────────────────────────
const COLS_FUNC = [
  'id','nome','cpf','dataNascimento','cargo','salario','insalubridade','adicionalSalario','ats',
  'salarioFamilia','auxilioCreche','auxilioCombustivel',
  'dataAdmissao','dataDemissao',
  'endereco','telefone','email','pis','ctps',
  'banco','agencia','conta','tipoConta','pix','obs'
];

const COLS_LANC = [
  // — Identificação —
  'id', 'funcionarioId', 'funcionarioNome',
  // — Datas / Período —
  'mes', 'competencia', 'dataLancamento',
  // — Classificação —
  'categoria', 'tipo',
  // — Valor principal —
  'valor', 'unidade',
  // — Hora Extra —
  'qtdHoras', 'valorHora',
  // — Vale Refeição —
  'diasVR', 'valorDiaVR', 'totalVR',
  // — Empréstimo —
  'totalEmprestimo', 'totalParcelas', 'parcelaAtual',
  // — Falta —
  'qtdFaltas',
  // — Genérico —
  'quantidade', 'justificado',
  // — Notas e Anexo —
  'obs', 'anexoNome', 'anexoBase64'
];

const COLS_FERIAS = [
  // — Identificação —
  'id', 'funcionarioId', 'funcionarioNome',
  // — Datas —
  'dataLancamento', 'periodoInicio', 'periodoFim',
  // — Verbas de Férias —
  'salarioFerias', 'insalubridadeFerias', 'adicionalFerias', 'atsFerias',
  'auxilioCombustivelFerias', 'auxilioCrecheFerias', 'salarioFamiliaFerias'
];

// ═══════════════════════════════════════════════════════
//  ENTRY POINT — recebe chamadas do front-end
// ═══════════════════════════════════════════════════════
function doPost(e) {
  const headers = {
    'Access-Control-Allow-Origin': '*',
    'Content-Type': 'application/json'
  };

  try {
    const body   = JSON.parse(e.postData.contents);
    const action = body.action;
    let result;

    if      (action === 'ping')              result = { msg: 'API Sulmak online' };
    else if (action === 'getAll')            result = getAll();
    else if (action === 'saveFuncionario')   result = saveFuncionario(body.data);
    else if (action === 'deleteFuncionario') result = deleteFuncionario(body.id);
    else if (action === 'saveLancamento')    result = saveLancamento(body.data);
    else if (action === 'deleteLancamento')  result = deleteLancamento(body.id);
    else if (action === 'saveFerias')        result = saveFerias(body.data);
    else if (action === 'salvarAnexo')       result = salvarAnexoNoDrive(body.nome, body.base64, body.mime);
    else throw new Error('Ação desconhecida: ' + action);

    return ContentService
      .createTextOutput(JSON.stringify({ ok: true, ...result }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ ok: false, error: err.message }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// Necessário para que o CORS preflight funcione
function doGet(e) {
  return ContentService
    .createTextOutput(JSON.stringify({ ok: true, msg: 'API Sulmak online' }))
    .setMimeType(ContentService.MimeType.JSON);
}

// ═══════════════════════════════════════════════════════
//  getAll — retorna funcionários + lançamentos + férias
// ═══════════════════════════════════════════════════════
function getAll() {
  const funcionarios = lerAba(ABA_FUNCIONARIOS, COLS_FUNC);
  const lancamentos  = lerAba(ABA_LANCAMENTOS,  COLS_LANC);
  const ferias       = lerAba(ABA_FERIAS,        COLS_FERIAS);
  return { funcionarios, lancamentos, ferias };
}

// ═══════════════════════════════════════════════════════
//  saveFuncionario — cria ou atualiza
//
//  Usa o cabeçalho REAL da planilha para mapear colunas
//  por NOME, evitando dados em colunas erradas quando a
//  planilha teve colunas adicionadas por versão antiga.
// ═══════════════════════════════════════════════════════
function saveFuncionario(data) {
  const sheet = getOrCreateSheet(ABA_FUNCIONARIOS, COLS_FUNC);

  const cabecalho = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0].map(String);
  const rowData   = cabecalho.map(col => data[col] !== undefined ? data[col] : '');

  if (!data.id) {
    data.id = gerarId();
    rowData[cabecalho.indexOf('id')] = data.id;
    sheet.appendRow(rowData);
  } else {
    const linha = encontrarLinha(sheet, data.id);
    if (linha) {
      sheet.getRange(linha, 1, 1, rowData.length).setValues([rowData]);
    } else {
      sheet.appendRow(rowData);
    }
  }
  return { id: data.id };
}

// ═══════════════════════════════════════════════════════
//  deleteFuncionario
// ═══════════════════════════════════════════════════════
function deleteFuncionario(id) {
  const sheet = getOrCreateSheet(ABA_FUNCIONARIOS, COLS_FUNC);
  const linha = encontrarLinha(sheet, id);
  if (!linha) throw new Error('Funcionário não encontrado: ' + id);
  sheet.deleteRow(linha);
  return { deleted: id };
}

// ═══════════════════════════════════════════════════════
//  saveLancamento — cria ou atualiza (upsert)
//  Se o ID existir na planilha → atualiza a linha.
//  Se o ID NÃO existir → insere como novo registro.
// ═══════════════════════════════════════════════════════
function saveLancamento(data) {
  const sheet = getOrCreateSheet(ABA_LANCAMENTOS, COLS_LANC);

  const cabecalho = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0].map(String);
  const rowData   = cabecalho.map(col => data[col] !== undefined ? data[col] : '');

  if (!data.id) {
    data.id = gerarId();
    rowData[cabecalho.indexOf('id')] = data.id;
    sheet.appendRow(rowData);
  } else {
    const linha = encontrarLinha(sheet, data.id);
    if (linha) {
      sheet.getRange(linha, 1, 1, rowData.length).setValues([rowData]);
    } else {
      sheet.appendRow(rowData);
    }
  }
  return { id: data.id };
}

// ═══════════════════════════════════════════════════════
//  deleteLancamento
// ═══════════════════════════════════════════════════════
function deleteLancamento(id) {
  const sheet = getOrCreateSheet(ABA_LANCAMENTOS, COLS_LANC);
  const linha = encontrarLinha(sheet, id);
  if (!linha) throw new Error('Lançamento não encontrado: ' + id);
  sheet.deleteRow(linha);
  return { deleted: id };
}

// ═══════════════════════════════════════════════════════
//  saveFerias — cria ou atualiza registro de férias
//
//  A lógica de "criar novo vs atualizar" é controlada
//  pelo front-end (HTML):
//    • Período novo  → front envia um ID novo → aqui INSERE
//    • Mesmo período → front envia o ID existente → aqui ATUALIZA
//
//  Dessa forma a aba Ferias funciona como histórico
//  completo: cada período diferente fica em uma linha
//  própria, nunca sobrescrevendo períodos anteriores.
// ═══════════════════════════════════════════════════════
function saveFerias(data) {
  const sheet = getOrCreateSheet(ABA_FERIAS, COLS_FERIAS);

  const cabecalho = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0].map(String);
  const rowData   = cabecalho.map(col => data[col] !== undefined ? data[col] : '');

  if (!data.id) {
    // Sem ID → novo registro
    data.id = gerarId();
    rowData[cabecalho.indexOf('id')] = data.id;
    sheet.appendRow(rowData);
  } else {
    const linha = encontrarLinha(sheet, data.id);
    if (linha) {
      // ID já existe → atualiza (mesmo período, valores corrigidos)
      sheet.getRange(linha, 1, 1, rowData.length).setValues([rowData]);
    } else {
      // ID não encontrado → insere como novo (período novo)
      sheet.appendRow(rowData);
    }
  }
  return { id: data.id };
}

// ═══════════════════════════════════════════════════════
//  UTILITÁRIOS
// ═══════════════════════════════════════════════════════

/** Lê uma aba e retorna array de objetos */
function lerAba(nomAba, colunas) {
  const sheet = getOrCreateSheet(nomAba, colunas);
  const dados = sheet.getDataRange().getValues();
  if (dados.length <= 1) return []; // só cabeçalho
  const cabecalho = dados[0];
  return dados.slice(1).map(row => {
    const obj = {};
    cabecalho.forEach((col, i) => { obj[col] = row[i]; });
    return obj;
  }).filter(obj => obj.id !== '' && obj.id !== undefined);
}

/** Retorna a aba, criando-a (com cabeçalho) se não existir.
 *  Se já existir mas o cabeçalho estiver desatualizado, adiciona colunas novas no final. */
function getOrCreateSheet(nome, colunas) {
  const ss  = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(nome);

  if (!sheet) {
    sheet = ss.insertSheet(nome);
    sheet.appendRow(colunas);
    const range = sheet.getRange(1, 1, 1, colunas.length);
    range.setFontWeight('bold');
    range.setBackground('#1B4F8A');
    range.setFontColor('#FFFFFF');
    sheet.setFrozenRows(1);
  } else {
    const cabecalhoAtual  = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0].map(String);
    const colunasFaltando = colunas.filter(c => !cabecalhoAtual.includes(c));
    if (colunasFaltando.length > 0) {
      colunasFaltando.forEach(col => {
        const novaColuna = sheet.getLastColumn() + 1;
        sheet.getRange(1, novaColuna).setValue(col);
        sheet.getRange(1, novaColuna).setFontWeight('bold').setBackground('#1B4F8A').setFontColor('#FFFFFF');
      });
    }
  }

  return sheet;
}

/** Encontra o número da linha pelo campo 'id' (coluna A = índice 0) */
function encontrarLinha(sheet, id) {
  const dados = sheet.getDataRange().getValues();
  for (let i = 1; i < dados.length; i++) {
    if (String(dados[i][0]) === String(id)) return i + 1; // linha é 1-based
  }
  return null;
}

// ═══════════════════════════════════════════════════════
//  migrarCabecalhos — rode UMA VEZ após atualizar o script
//  Insere colunas novas na posição correta conforme COLS_LANC
// ═══════════════════════════════════════════════════════
function migrarCabecalhos() {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(ABA_LANCAMENTOS);
  if (!sheet) { Logger.log('Aba Lancamentos não encontrada.'); return; }

  COLS_LANC.forEach((col, idxDesejado) => {
    const cabecalhoAtual = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0].map(String);
    if (!cabecalhoAtual.includes(col)) {
      const insertAt = idxDesejado + 1;
      sheet.insertColumnBefore(insertAt);
      const cell = sheet.getRange(1, insertAt);
      cell.setValue(col);
      cell.setFontWeight('bold').setBackground('#1B4F8A').setFontColor('#FFFFFF');
      Logger.log('Coluna inserida: ' + col + ' na posição ' + insertAt);
    }
  });

  Logger.log('Migração de Lançamentos concluída!');
}

// ═══════════════════════════════════════════════════════
//  migrarCabecalhosFuncionarios — rode UMA VEZ para
//  reorganizar as colunas da aba Funcionarios na ordem
//  correta de COLS_FUNC.
// ═══════════════════════════════════════════════════════
function migrarCabecalhosFuncionarios() {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(ABA_FUNCIONARIOS);
  if (!sheet) { Logger.log('Aba Funcionarios não encontrada.'); return; }

  COLS_FUNC.forEach((col, idxDesejado) => {
    const cabecalhoAtual = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0].map(String);
    if (!cabecalhoAtual.includes(col)) {
      const insertAt = idxDesejado + 1;
      sheet.insertColumnBefore(insertAt);
      const cell = sheet.getRange(1, insertAt);
      cell.setValue(col);
      cell.setFontWeight('bold').setBackground('#1B4F8A').setFontColor('#FFFFFF');
      Logger.log('Coluna inserida: ' + col + ' na posição ' + insertAt);
    }
  });

  Logger.log('Migração de Funcionários concluída!');
}

// ═══════════════════════════════════════════════════════
//  migrarCabecalhosFerias — rode UMA VEZ para criar /
//  reorganizar a aba Ferias com todas as colunas corretas.
// ═══════════════════════════════════════════════════════
function migrarCabecalhosFerias() {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  let   sheet = ss.getSheetByName(ABA_FERIAS);

  if (!sheet) {
    // Aba não existe → cria do zero
    sheet = ss.insertSheet(ABA_FERIAS);
    sheet.appendRow(COLS_FERIAS);
    const range = sheet.getRange(1, 1, 1, COLS_FERIAS.length);
    range.setFontWeight('bold').setBackground('#1B4F8A').setFontColor('#FFFFFF');
    sheet.setFrozenRows(1);
    Logger.log('Aba Ferias criada com sucesso!');
    return;
  }

  // Aba existe → insere colunas faltando na posição correta
  COLS_FERIAS.forEach((col, idxDesejado) => {
    const cabecalhoAtual = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0].map(String);
    if (!cabecalhoAtual.includes(col)) {
      const insertAt = idxDesejado + 1;
      sheet.insertColumnBefore(insertAt);
      const cell = sheet.getRange(1, insertAt);
      cell.setValue(col);
      cell.setFontWeight('bold').setBackground('#1B4F8A').setFontColor('#FFFFFF');
      Logger.log('Coluna inserida: ' + col + ' na posição ' + insertAt);
    }
  });

  Logger.log('Migração de Férias concluída!');
}

/** Gera um ID único */
function gerarId() {
  return new Date().getTime().toString(36) + Math.random().toString(36).substring(2, 6);
}

// ═══════════════════════════════════════════════════════
//  salvarAnexoNoDrive — salva arquivo no Google Drive
//  e retorna o link de visualização
// ═══════════════════════════════════════════════════════
function salvarAnexoNoDrive(nome, base64, mime) {
  const PASTA_NOME = 'Sulmak_Anexos';
  let folder;
  const folders = DriveApp.getFoldersByName(PASTA_NOME);
  if (folders.hasNext()) {
    folder = folders.next();
  } else {
    folder = DriveApp.createFolder(PASTA_NOME);
  }

  const bytes = Utilities.base64Decode(base64);
  const blob  = Utilities.newBlob(bytes, mime || 'application/octet-stream', nome);
  const file  = folder.createFile(blob);

  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);

  return { url: file.getUrl(), fileId: file.getId() };
}

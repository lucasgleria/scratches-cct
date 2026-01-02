/** =========================
 * BACKEND (Apps Script)
 * =========================
 */

/** CONFIGURAÇÕES GERAIS */
const IMPORT_SPREADSHEETS = [
  { id: '1Qrzq3NatjRtLE8CiQbMiWRHvwFgUKA5ymimoR6JAsV0', name: 'CCT Teste' },
  { id: '1qgP2RIXiA5cjO-EdSUjti11r6jBXVJR0PeMotHoBAA4', name: 'CCT Teste 2' }
];

const EXPORT_SPREADSHEETS = [
  { id: 'EXPORT_ID_1', name: 'Planilha de Exportação 1' },
  { id: 'EXPORT_ID_2', name: 'Planilha de Exportação 2' }
];

const MULTI_JOIN = '; ';

/** CONSTANTES DE ESTRUTURA E ESTILO */
// Índices baseados em 0 (Array)
const COLS = { MAWB: 0, HOUSE: 1, REF: 2, CONS: 3, ENT: 4, DTA: 5, PREV: 6, RESP: 7, OBS: 8 };

const COLORS = {
  LIGHT_BLUE: '#cfe2f3',
  REGULAR_BLUE: '#9fc5e8',
  GREEN: '#d9ead3',
  ORANGE: '#f6b26b',
  YELLOW_HIGHLIGHT: '#fff2cc',
  NO_COLOR: null
};

const FRIDGE_CODES = ["FRI", "FRO", "COL", "ERT", "CRT"];

// -----------------------------------------------------------------------------
// UTILITÁRIOS & HELPERS
// -----------------------------------------------------------------------------

function _asError(message, detail) {
  return { ok: false, message, detail: detail || null };
}

function _asOk(payload) {
  return { ok: true, payload };
}

function _norm(v) {
  return v == null ? '' : String(v).trim();
}

function _normHouse(v) {
  return _norm(v).toUpperCase();
}

/** * Helper centralizado para obter a aba de trabalho (3ª aba).
 * Evita repetição de código de abertura e validação de planilha.
 */
function _getWorksheet(spreadsheetId) {
  if (!spreadsheetId) throw new Error('ID da planilha não fornecido.');
  const ss = SpreadsheetApp.openById(spreadsheetId);
  const sheets = ss.getSheets();
  if (sheets.length < 3) throw new Error('A planilha selecionada não tem pelo menos 3 abas.');
  return { ws: sheets[2], ssName: ss.getName() };
}

// -----------------------------------------------------------------------------
// VALIDAÇÃO
// -----------------------------------------------------------------------------

function validateMAWB(mawb) {
  let normalized = _norm(mawb);
  if (!normalized) return { valid: false, message: 'MAWB é obrigatório' };
  
  normalized = normalized.replace(/-/g, '');
  
  if (!/^\d+$/.test(normalized)) return { valid: false, message: 'MAWB deve conter apenas números' };
  if (normalized.length !== 11) return { valid: false, message: 'MAWB deve ter exatamente 11 dígitos' };
  
  return { valid: true };
}

function validateHOUSE(house) {
  const normalized = _norm(house);
  if (!normalized) return { valid: false, message: 'HOUSE é obrigatório' };
  if (normalized.length > 11) return { valid: false, message: 'HOUSE contém mais de 11 caracteres' };
  return { valid: true };
}

// -----------------------------------------------------------------------------
// WEB APP (HANDLERS)
// -----------------------------------------------------------------------------

function doGet() {
  return HtmlService.createTemplateFromFile('index')
    .evaluate()
    .setTitle('Sistema de Gestão MAWB/HOUSE')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function getHtmlForView(viewName) {
  if (['importacao', 'exportacao'].includes(viewName)) {
    return HtmlService.createTemplateFromFile(viewName).evaluate().getContent();
  }
  return '<div>View não encontrada</div>';
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

function getSpreadsheets(viewName) {
  try {
    const map = { 'importacao': IMPORT_SPREADSHEETS, 'exportacao': EXPORT_SPREADSHEETS };
    return map[viewName] ? _asOk(map[viewName]) : _asError('View desconhecida: ' + viewName);
  } catch (err) {
    return _asError('Erro ao carregar planilhas', err.message);
  }
}

// -----------------------------------------------------------------------------
// TRIGGERS & LÓGICA DE CORES
// -----------------------------------------------------------------------------

function onEdit(e) {
  const range = e.range;
  const sheet = e.source.getActiveSheet();
  const editedCol = range.getColumn();
  const editedRow = range.getRow();

  // Colunas (1-based): ENT(5), DTA(6), OBS(9) -> correspondem aos índices do COLS + 1
  const TRIGGER_COLS = [COLS.ENT + 1, COLS.DTA + 1, COLS.OBS + 1];

  if (TRIGGER_COLS.includes(editedCol) && range.getNumRows() === 1) {
    if (editedRow === 1) {
      const firstCell = _norm(sheet.getRange(1, 1).getValue()).toUpperCase();
      if (firstCell.includes('MAWB')) return;
    }

    const rowData = sheet.getRange(editedRow, 1, 1, 9).getValues()[0];
    const colors = getRowColors(rowData, e.source.getName());
    sheet.getRange(editedRow, 1, 1, colors.length).setBackgrounds([colors]);
  }
}

function getRowColors(rowData, sheetName) {
  const colors = new Array(COLS.OBS + 1).fill(COLORS.NO_COLOR);

  const entregaValue = _norm(rowData[COLS.ENT]).toUpperCase();
  const dtaValue = _norm(rowData[COLS.DTA]).toUpperCase();
  const obsValue = _norm(rowData[COLS.OBS]).toUpperCase();

  // Regra 1: Exportação
  if (entregaValue === 'EXPORTAÇÃO') {
    return colors.fill(COLORS.LIGHT_BLUE);
  }

  // Regra 2: DTA
  if (dtaValue && !dtaValue.startsWith('GRU') && !dtaValue.startsWith('VCP')) {
    colors.fill(COLORS.GREEN, COLS.MAWB, COLS.RESP + 1); // Preenche do início até RESP
    if (sheetName === 'CCT Teste' && dtaValue === 'SSA') {
      colors[COLS.DTA] = COLORS.ORANGE;
    }
  }

  // Regra 3: Carga de Geladeira
  if (FRIDGE_CODES.includes(obsValue)) {
    colors[COLS.RESP] = COLORS.REGULAR_BLUE;
    colors[COLS.OBS] = COLORS.REGULAR_BLUE;
  }

  return colors;
}

function _updateMawbBlockColoring(sheet, mawb) {
  const mawbNorm = _norm(mawb);
  if (!mawbNorm) return;

  const textFinder = sheet.createTextFinder(mawbNorm).matchEntireCell(true);
  const occurrences = textFinder.findAll();
  if (occurrences.length === 0) return;

  const firstRow = occurrences[0].getRowIndex();
  const lastRow = occurrences[occurrences.length - 1].getRowIndex();
  const blockRange = sheet.getRange(firstRow, 1, lastRow - firstRow + 1, 9);
  const blockValues = blockRange.getValues();

  // Verifica múltiplos houses
  const uniqueHouses = new Set();
  blockValues.forEach(row => {
    const house = _normHouse(row[COLS.HOUSE]);
    if (house) uniqueHouses.add(house);
  });
  const isMultiHouse = uniqueHouses.size > 1;

  const sheetName = sheet.getParent().getName();
  const newBackgrounds = blockValues.map(rowData => {
    const rowColors = getRowColors(rowData, sheetName);
    if (isMultiHouse) {
      rowColors[COLS.MAWB] = COLORS.YELLOW_HIGHLIGHT;
    }
    return rowColors;
  });

  blockRange.setBackgrounds(newBackgrounds);
}

// -----------------------------------------------------------------------------
// LÓGICA DE NEGÓCIOS (LEITURA)
// -----------------------------------------------------------------------------

function checkHouseExists(spreadsheetId, house) {
  try {
    if (!house) return _asError('Código HOUSE não fornecido.');
    
    const { ws } = _getWorksheet(spreadsheetId);
    
    const finder = ws.getRange("B:B").createTextFinder(_normHouse(house)).matchEntireCell(true);
    const occurrences = finder.findAll();

    return _asOk(occurrences.length > 0);
  } catch (err) {
    return _asError('Falha ao verificar a existência do HOUSE.', err.message);
  }
}

function getHousesFromSpreadsheet(spreadsheetId) {
  try {
    const { ws } = _getWorksheet(spreadsheetId);
    const lastRow = ws.getLastRow();
    
    if (lastRow < 2) return _asOk([]);

    const houseColumn = ws.getRange(2, 2, lastRow - 1, 1).getValues();
    const uniqueHouses = [...new Set(houseColumn.map(row => _normHouse(row[0])).filter(Boolean))];
    uniqueHouses.sort();

    return _asOk(uniqueHouses);
  } catch (err) {
    return _asError('Falha ao buscar HOUSEs da planilha.', err.message);
  }
}

function getDataForHouse(spreadsheetId, house) {
  try {
    if (!house) return _asError('Código HOUSE não fornecido.');
    
    let ws;
    try {
      ws = _getWorksheet(spreadsheetId).ws;
    } catch (e) {
      return _asError('Falha ao abrir a planilha.', e.message);
    }

    // Busca última ocorrência
    const finder = ws.getRange("B:B").createTextFinder(_normHouse(house)).matchEntireCell(true);
    const occurrences = finder.findAll().reverse();
    
    if (occurrences.length === 0) return _asError(`HOUSE "${house}" não encontrado na planilha.`);
    
    const latestRowIndex = occurrences[0].getRow();
    const rowData = ws.getRange(latestRowIndex, 1, 1, 9).getValues()[0];

    const splitField = (val) => _norm(val).split(MULTI_JOIN).filter(Boolean);

    const data = {
      mawb: _norm(rowData[COLS.MAWB]),
      house: _norm(rowData[COLS.HOUSE]),
      refs: splitField(rowData[COLS.REF]),
      consignees: splitField(rowData[COLS.CONS]),
      entregas: splitField(rowData[COLS.ENT]),
      dtas: splitField(rowData[COLS.DTA]),
      previsoes: splitField(rowData[COLS.PREV]),
      responsaveis: splitField(rowData[COLS.RESP]),
      observacoes: splitField(rowData[COLS.OBS]),
    };

    return _asOk(data);

  } catch (err) {
    return _asError('Ocorreu um erro inesperado em getDataForHouse.', `${err.message} stack: ${err.stack}`);
  }
}

// -----------------------------------------------------------------------------
// LÓGICA DE NEGÓCIOS (ESCRITA)
// -----------------------------------------------------------------------------

function saveEntries(payload) {
  try {
    if (!payload || typeof payload !== 'object') return _asError('Payload inválido.');

    const {
      spreadsheetId, mawb: rawMawb, houses, isEditMode = false,
      refs = [], consignees = [], entregas = [], dtas = [], 
      previsoes = [], responsaveis = [], observacoes = []
    } = payload;

    const mawb = _norm(rawMawb).replace(/-/g, '');
    if (!Array.isArray(houses) || houses.length === 0) return _asError('Adicione ao menos um HOUSE.');

    // Validações
    const mawbVal = validateMAWB(mawb);
    if (!mawbVal.valid) return _asError(mawbVal.message);

    for (const h of houses) {
      const hVal = validateHOUSE(h);
      if (!hVal.valid) return _asError(`HOUSE "${h}": ${hVal.message}`);
    }

    const { ws } = _getWorksheet(spreadsheetId);

    // Helpers locais para lógica de merge
    const joinVals = (vals) => (vals.length ? vals.join(MULTI_JOIN) : '');
    const mergeCell = (oldVal, newVal, isEdit) => {
      const a = _norm(oldVal);
      const b = _norm(newVal);
      if (isEdit) return b;
      if (!a) return b;
      if (!b) return a;
      const parts = a.split(MULTI_JOIN).map(_norm);
      return parts.includes(b) ? a : `${a}${MULTI_JOIN}${b}`;
    };

    const mawbNorm = _norm(mawb);

    // 1. Identificar bloco existente do MAWB
    const finder = ws.createTextFinder(mawbNorm).matchEntireCell(true);
    const occurrences = finder.findAll();
    
    let table = [];
    let firstDataRow = 1;
    let endRow = 0;

    if (occurrences.length > 0) {
      firstDataRow = occurrences[0].getRowIndex();
      endRow = occurrences[occurrences.length - 1].getRowIndex();
      if (endRow >= firstDataRow) {
        table = ws.getRange(firstDataRow, 1, endRow - firstDataRow + 1, 9).getValues();
      }
    }

    // 2. Indexar dados existentes
    const keyFor = (houseVal) => _normHouse(houseVal);
    const index = new Map();
    const duplicatesToDelete = [];

    table.forEach((row, i) => {
      if (row.every(cell => _norm(cell) === '')) return;
      
      const key = keyFor(row[COLS.HOUSE]);
      const absRow = firstDataRow + i;

      if (!index.has(key)) {
        index.set(key, { rowIndex: absRow, row: row.slice() });
      } else {
        // Merge duplicatas encontradas na leitura
        const keeper = index.get(key);
        for (let c = COLS.REF; c <= COLS.OBS; c++) {
          keeper.row[c] = mergeCell(keeper.row[c], row[c], isEditMode);
        }
        duplicatesToDelete.push(absRow);
      }
    });

    // 3. Preparar dados novos (Input Payload)
    const housesClean = Array.from(new Set(houses.map(_normHouse)));
    const houseDataMap = new Map();
    
    housesClean.forEach(h => houseDataMap.set(h, { 
      mawb: mawbNorm, house: h, 
      refs: [], consignees: [], entregas: [], dtas: [], 
      previsoes: [], responsaveis: [], observacoes: [] 
    }));

    const appendToMap = (arr, field) => {
      arr.forEach(({ house, value }) => {
        const h = _normHouse(house);
        if (houseDataMap.has(h)) houseDataMap.get(h)[field].push(_norm(value));
      });
    };

    appendToMap(refs, 'refs');
    appendToMap(consignees, 'consignees');
    appendToMap(entregas, 'entregas');
    appendToMap(dtas, 'dtas');
    appendToMap(previsoes, 'previsoes');
    appendToMap(responsaveis, 'responsaveis');
    appendToMap(observacoes, 'observacoes');

    const calculatedRows = [...houseDataMap.values()].map(d => ([
      d.mawb, d.house, 
      joinVals(d.refs), joinVals(d.consignees), joinVals(d.entregas), 
      joinVals(d.dtas), joinVals(d.previsoes), joinVals(d.responsaveis), joinVals(d.observacoes)
    ]));

    // 4. Separar Updates de Inserts
    const toInsert = [];
    const pendingUpdates = [];

    calculatedRows.forEach(r => {
      const key = keyFor(r[COLS.HOUSE]);
      if (index.has(key)) {
        const kept = index.get(key);
        const updated = kept.row.slice();
        for (let c = COLS.REF; c <= COLS.OBS; c++) {
          updated[c] = mergeCell(updated[c], r[c], isEditMode);
        }
        pendingUpdates.push({ rowIndex: kept.rowIndex, row: updated });
        index.set(key, { rowIndex: kept.rowIndex, row: updated }); // Atualiza index para caso de uso futuro
      } else {
        toInsert.push(r);
      }
    });

    // 5. Aplicar alterações na planilha
    
    // Updates
    pendingUpdates.forEach(update => {
      const range = ws.getRange(update.rowIndex, 1, 1, 9);
      range.setNumberFormat('@').setValues([update.row]);
    });

    // Delete duplicatas antigas
    if (duplicatesToDelete.length) {
      duplicatesToDelete.sort((a, b) => b - a).forEach(r => ws.deleteRow(r));
    }

    // Inserts
    if (toInsert.length) {
      let insertRow;
      if (occurrences.length > 0) {
        insertRow = endRow + 1;
        ws.insertRowsAfter(endRow, toInsert.length);
      } else {
        const lastRow = ws.getLastRow();
        insertRow = lastRow === 0 ? 1 : lastRow + 2;
      }
      const range = ws.getRange(insertRow, 1, toInsert.length, 9);
      range.setNumberFormat('@').setValues(toInsert);
    }

    _updateMawbBlockColoring(ws, mawbNorm);

    return _asOk({
      inserted: toInsert.length,
      updated: pendingUpdates.length,
      removed_duplicates: duplicatesToDelete.length
    });

  } catch (err) {
    return _asError('Falha ao salvar os dados.', err && err.message ? err.message : String(err));
  }
}
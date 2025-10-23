/** =========================
 *  BACKEND (Apps Script)
 *  =========================
 *  - Garante 1 linha por (MAWB, HOUSE)
 *  - UPSERT: atualiza se já existir; insere se não existir
 *  - Consolida duplicatas antigas do mesmo (MAWB, HOUSE)
 *  - NOVA FEATURE: Mantém sempre uma linha em branco como separador entre grupos de dados
 *    * Se as últimas 2 linhas estão vazias: insere na penúltima (mantém separador no final)
 *    * Se apenas a última está vazia: pula uma linha e insere (cria novo separador)
 *    * Se não há linhas vazias: adiciona separador + dados
 */

/** Liste aqui as planilhas disponíveis para o usuário escolher. */
const SPREADSHEETS = [
  // Exemplo:
  { id: '1Qrzq3NatjRtLE8CiQbMiWRHvwFgUKA5ymimoR6JAsV0', name: 'CCT Teste' },
  { id: '1qgP2RIXiA5cjO-EdSUjti11r6jBXVJR0PeMotHoBAA4', name: 'CCT Teste 2' }
  // Adicione mais planilhas aqui conforme necessário
];

/** Separador para múltiplos valores na mesma célula */
const MULTI_JOIN = '; ';

/** Utilitário: erro */
function _asError(message, detail) {
  return { ok: false, message, detail: detail || null };
}

/** Utilitário: ok */
function _asOk(payload) {
  return { ok: true, payload };
}

/** Normaliza string */
function _norm(v) {
  return v == null ? '' : String(v).trim();
}

/** Normaliza HOUSE (trim + maiúsculas) */
function _normHouse(v) {
  return _norm(v).toUpperCase();
}

/** Valida MAWB - deve ser numérico com exatamente 11 dígitos */
function validateMAWB(mawb) {
  let normalized = _norm(mawb);

  if (!normalized) {
    return { valid: false, message: 'MAWB é obrigatório' };
  }

  // Remove o hífen para validação
  normalized = normalized.replace(/-/g, '');

  if (!/^\d+$/.test(normalized)) {
    return { valid: false, message: 'MAWB deve conter apenas números' };
  }

  if (normalized.length !== 11) {
    return { valid: false, message: 'MAWB deve ter exatamente 11 dígitos' };
  }

  return { valid: true };
}

/** Valida HOUSE - máximo 11 caracteres */
function validateHOUSE(house) {
  const normalized = _norm(house);

  if (!normalized) {
    return { valid: false, message: 'HOUSE é obrigatório' };
  }

  if (normalized.length > 11) {
    return { valid: false, message: 'HOUSE contém mais de 11 caracteres' };
  }

  return { valid: true };
}

/** Serve a interface HTML */
function doGet() {
  return HtmlService.createTemplateFromFile('index')
    .evaluate()
    .setTitle('Sistema de Gestão MAWB/HOUSE')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/** Função para incluir arquivos CSS/JS externos se necessário */
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

/** Retorna a lista de planilhas disponíveis */
function getSpreadsheets() {
  try {
    return _asOk(SPREADSHEETS);
  } catch (err) {
    return _asError('Erro ao carregar planilhas', err.message);
  }
}

/**
 * Trigger que é executado quando uma célula é editada manualmente.
 * @param {Event} e O objeto de evento.
 */
function onEdit(e) {
  const range = e.range;
  const sheet = e.source.getActiveSheet();
  const spreadsheet = e.source; // A planilha inteira
  const editedCol = range.getColumn();
  const editedRow = range.getRow();

  // Colunas que disparam a formatação
  const TRIGGER_COLS = [5, 6, 9]; // ENT, DTA, OBS

  if (TRIGGER_COLS.includes(editedCol) && range.getNumRows() === 1) {
    // Ignora o cabeçalho
    if (editedRow === 1) {
      const firstCell = _norm(sheet.getRange(1, 1).getValue()).toUpperCase();
      if (firstCell.includes('MAWB')) return;
    }

    const rowData = sheet.getRange(editedRow, 1, 1, 9).getValues()[0];
    const colors = getRowColors(rowData, spreadsheet.getName()); // Usa o nome do arquivo da planilha
    sheet.getRange(editedRow, 1, 1, colors.length).setBackgrounds([colors]);
  }
}


/**
 * Calcula a formatação de cores para uma linha com base nos valores das colunas.
 * @param {Array<String>} rowData Os dados da linha.
 * @param {String} sheetName O nome da planilha.
 * @returns {Array<String>} Uma matriz de códigos de cores para a linha.
 */
function getRowColors(rowData, sheetName) {
  const COLS = { MAWB: 0, HOUSE: 1, REF: 2, CONS: 3, ENT: 4, DTA: 5, PREV: 6, RESP: 7, OBS: 8 };
  const COLORS = {
    LIGHT_BLUE: '#cfe2f3',
    REGULAR_BLUE: '#9fc5e8',
    GREEN: '#d9ead3',
    ORANGE: '#f6b26b',
    NO_COLOR: null
  };
  const FRIDGE_CODES = ["FRI", "FRO", "COL", "ERT", "CRT"];

  const colors = new Array(COLS.OBS + 1).fill(COLORS.NO_COLOR);

  const entregaValue = _norm(rowData[COLS.ENT]).toUpperCase();
  const dtaValue = _norm(rowData[COLS.DTA]).toUpperCase();
  const obsValue = _norm(rowData[COLS.OBS]).toUpperCase();

  // Regra 1: Exportação (prioridade máxima)
  if (entregaValue === 'EXPORTAÇÃO') {
    return colors.fill(COLORS.LIGHT_BLUE);
  }

  // Regra 2: DTA
  if (dtaValue && !dtaValue.startsWith('GRU') && !dtaValue.startsWith('VCP')) {
    for (let i = COLS.MAWB; i <= COLS.RESP; i++) {
      colors[i] = COLORS.GREEN;
    }
    // Regra específica para a planilha "CCT Teste"
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


/**
 * saveEntries(payload)
 * payload:
 * {
 *  spreadsheetId: string,
 *  mawb: string,
 *  houses: string[],                       // ["H1","H2",...]
 *  refs?:        {house,value}[],
 *  consignees?:  {house,value}[],
 *  entregas?:    {house,value}[],
 *  dtas?:        {house,value}[],
 *  previsoes?:   {house,value}[],
 *  responsaveis?:{house,value}[],
 *  observacoes?: {house,value}[]
 * }
 */
function saveEntries(payload) {
  try {
    if (!payload || typeof payload !== 'object') return _asError('Payload inválido.');

    const {
      spreadsheetId,
      mawb: rawMawb, // Renomeia para indicar que pode ter o hífen
      houses,
      refs = [],
      consignees = [],
      entregas = [],
      dtas = [],
      previsoes = [],
      responsaveis = [],
      observacoes = []
    } = payload;

    // Remove o hífen do MAWB para consistência
    const mawb = _norm(rawMawb).replace(/-/g, '');

    if (!spreadsheetId) return _asError('Selecione uma planilha.');
    if (!Array.isArray(houses) || houses.length === 0) return _asError('Adicione ao menos um HOUSE.');

    // Validações de entrada
    const mawbValidation = validateMAWB(mawb);
    if (!mawbValidation.valid) {
      return _asError(mawbValidation.message);
    }

    // Valida todos os HOUSE codes
    for (let i = 0; i < houses.length; i++) {
      const houseValidation = validateHOUSE(houses[i]);
      if (!houseValidation.valid) {
        return _asError(`HOUSE "${houses[i]}": ${houseValidation.message}`);
      }
    }

    const ss = SpreadsheetApp.openById(spreadsheetId);
    const sheets = ss.getSheets();
    if (sheets.length < 3) return _asError('A planilha selecionada não possui 3 abas.');
    const ws = sheets[2]; // terceira aba

    // ----------------- Helpers -----------------
    const COLS = { MAWB:1, HOUSE:2, REF:3, CONS:4, ENT:5, DTA:6, PREV:7, RESP:8, OBS:9 };

    const collectByHouse = (arr, house) => {
      const values = arr
        .filter(x => _normHouse(x && x.house) === house)
        .map(x => _norm(x && x.value))
        .filter(Boolean);
      return Array.from(new Set(values));
    };

    const joinVals = (vals) => (vals.length ? vals.join(MULTI_JOIN) : '');

    const mergeCell = (oldVal, newVal) => {
      const a = _norm(oldVal);
      const b = _norm(newVal);
      if (!a && !b) return '';
      if (!a) return b;
      if (!b) return a;
      const parts = a.split(MULTI_JOIN).map(_norm);
      return parts.includes(b) ? a : `${a}${MULTI_JOIN}${b}`;
    };

    const mawbNorm = _norm(mawb);

    // --- Performance Improvement: Use TextFinder to locate MAWB block ---
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

    const keyFor = (houseVal) => _normHouse(houseVal);
    const index = new Map();
    const duplicatesToDelete = [];

    table.forEach((row, i) => {
      if (row.every(cell => _norm(cell) === '')) return;

      const key = keyFor(row[COLS.HOUSE - 1]);
      const absRow = firstDataRow + i;

      if (!index.has(key)) {
        index.set(key, { rowIndex: absRow, row: row.slice() });
      } else {
        const keeper = index.get(key);
        for (let c = COLS.REF; c <= COLS.OBS; c++) {
          keeper.row[c - 1] = mergeCell(keeper.row[c - 1], row[c - 1]);
        }
        duplicatesToDelete.push(absRow);
      }
    });

    const housesClean = Array.from(new Set(houses.map(_normHouse)));

    const houseDataMap = new Map();
    housesClean.forEach(h => {
      houseDataMap.set(h, {
        mawb: mawbNorm, house: h, refs: [], consignees: [], entregas: [],
        dtas: [], previsoes: [], responsaveis: [], observacoes: []
      });
    });

    const appendToMap = (map, arr, field) => {
      arr.forEach(({ house, value }) => {
        const h = _normHouse(house);
        if (map.has(h)) map.get(h)[field].push(_norm(value));
      });
    };

    appendToMap(houseDataMap, refs, 'refs');
    appendToMap(houseDataMap, consignees, 'consignees');
    appendToMap(houseDataMap, entregas, 'entregas');
    appendToMap(houseDataMap, dtas, 'dtas');
    appendToMap(houseDataMap, previsoes, 'previsoes');
    appendToMap(houseDataMap, responsaveis, 'responsaveis');
    appendToMap(houseDataMap, observacoes, 'observacoes');

    const calculatedRows = [...houseDataMap.values()].map(data => ([
      data.mawb, data.house, joinVals(data.refs), joinVals(data.consignees),
      joinVals(data.entregas), joinVals(data.dtas), joinVals(data.previsoes),
      joinVals(data.responsaveis), joinVals(data.observacoes)
    ]));

    const toInsert = [];
    const pendingUpdates = [];

    calculatedRows.forEach(r => {
      const key = keyFor(r[COLS.HOUSE - 1]);
      if (index.has(key)) {
        const kept = index.get(key);
        const updated = kept.row.slice();
        for (let c = COLS.REF; c <= COLS.OBS; c++) {
          updated[c - 1] = mergeCell(updated[c - 1], r[c - 1]);
        }
        pendingUpdates.push({ rowIndex: kept.rowIndex, row: updated });
        index.set(key, { rowIndex: kept.rowIndex, row: updated });
      } else {
        toInsert.push(r);
      }
    });

    pendingUpdates.forEach(update => {
      const range = ws.getRange(update.rowIndex, 1, 1, 9);
      range.setNumberFormat('@');
      range.setValues([update.row]);
      range.setBackgrounds([getRowColors(update.row, ss.getName())]);
    });

    if (duplicatesToDelete.length) {
      duplicatesToDelete.sort((a, b) => b - a).forEach(r => ws.deleteRow(r));
    }

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
      range.setNumberFormat('@');
      range.setValues(toInsert);
      range.setBackgrounds(toInsert.map(row => getRowColors(row, ss.getName())));
    }

    return _asOk({
      inserted: toInsert.length,
      updated: pendingUpdates.length,
      removed_duplicates: duplicatesToDelete.length
    });

  } catch (err) {
    return _asError('Falha ao salvar os dados.', err && err.message ? err.message : String(err));
  }
}
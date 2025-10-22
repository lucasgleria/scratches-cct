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
    const colors = getRowColors(rowData, sheet.getName());
    sheet.getRange(editedRow, 1, 1, colors.length).setBackgrounds([colors]);
  }
}

function findMawbGroupBoundaries(ws, mawb) {
  const lastRow = ws.getLastRow();
  if (lastRow === 0) {
    return null;
  }
  const mawbColumn = ws.getRange(1, 1, lastRow, 1).getValues();
  const normalizedMawb = _norm(mawb);
  let startRow = -1;
  let endRow = -1;

  for (let i = 0; i < lastRow; i++) {
    if (_norm(mawbColumn[i][0]) === normalizedMawb) {
      if (startRow === -1) {
        startRow = i + 1;
      }
      endRow = i + 1;
    } else {
      if (startRow !== -1) {
        // We've found the end of the group
        break;
      }
    }
  }

  if (startRow === -1) {
    return null;
  }

  return { startRow, endRow };
}

function findLastDataRow(ws) {
  const lastRow = ws.getLastRow();
  if (lastRow === 0) {
    return 0;
  }
  const range = ws.getRange(1, 1, lastRow, ws.getLastColumn()).getValues();
  for (let i = lastRow - 1; i >= 0; i--) {
    if (!range[i].every(c => _norm(c) === '')) {
      return i + 1;
    }
  }
  return 0;
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
    if (sheetName === 'CCT Teste') {
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
      return parts.includes(b) ? a : a + MULTI_JOIN + b;
    };

    // Detecta cabeçalho (opcional)
    const hasRows = ws.getLastRow() > 0;
    let headerRow = 0;
    if (hasRows) {
      const first = ws.getRange(1, 1, 1, 2).getValues()[0];
      if (/mawb/i.test(String(first[0])) && /house/i.test(String(first[1]))) headerRow = 1;
    }
    const firstDataRow = headerRow + 1;
    const lastRow = ws.getLastRow();

    // Lê tabela atual
    let table = [];
    if (lastRow >= firstDataRow) {
      table = ws.getRange(firstDataRow, 1, lastRow - firstDataRow + 1, 9).getValues();
    }

    // Índice por chave (case-insensitive para HOUSE)
    // Se existirem duplicatas já salvas, manteremos a primeira e consolidaremos as demais.
    const keyFor = (mawbVal, houseVal) => `${_norm(mawbVal)}|||${_normHouse(houseVal)}`;
    const index = new Map();           // key -> {rowIndex, row}
    const duplicatesToDelete = [];     // linha absoluta para deletar (será apagada no fim)

    table.forEach((row, i) => {
      // Pula linhas em branco para não as marcar como duplicatas
      if (row.every(cell => _norm(cell) === '')) {
        return;
      }
      const key = keyFor(row[COLS.MAWB - 1], row[COLS.HOUSE - 1]);
      const absRow = firstDataRow + i;
      if (!index.has(key)) {
        index.set(key, { rowIndex: absRow, row: row.slice() });
      } else {
        // Duplicata antiga -> consolidar e marcar para excluir
        const keeper = index.get(key);
        for (let c = COLS.REF; c <= COLS.OBS; c++) {
          keeper.row[c - 1] = mergeCell(keeper.row[c - 1], row[c - 1]);
        }
        index.set(key, keeper);
        duplicatesToDelete.push(absRow);
      }
    });

    // Normaliza e deduplica lista de houses do payload
    const housesClean = Array.from(new Set(houses.map(_normHouse)));

    // Valida vínculos (cada item deve apontar pra um HOUSE informado)
    const valid = new Set(housesClean);
    const checkLinked = (arr, label) => {
      for (let i = 0; i < arr.length; i++) {
        const h = _normHouse(arr[i] && arr[i].house);
        if (!valid.has(h)) throw new Error(`Item de ${label} vinculado a HOUSE inexistente: "${arr[i] && arr[i].house}".`);
      }
    };
    checkLinked(refs, 'REF');
    checkLinked(consignees, 'CONSIGNEE');
    checkLinked(entregas, 'ENTREGA');
    checkLinked(dtas, 'DTA');
    checkLinked(previsoes, 'PREVISÃO');
    checkLinked(responsaveis, 'RESPONSÁVEL');
    checkLinked(observacoes, 'OBSERVAÇÃO');

    // Calcula linhas-alvo (UMA por HOUSE)
    // Cria um mapa consolidado por HOUSE
    const houseDataMap = new Map();

    const mawbNorm = _norm(mawb);

    housesClean.forEach(h => {
      houseDataMap.set(h, {
        mawb: mawbNorm,
        house: h,
        refs: [],
        consignees: [],
        entregas: [],
        dtas: [],
        previsoes: [],
        responsaveis: [],
        observacoes: []
      });
    });

    const appendToMap = (map, arr, field) => {
      arr.forEach(({ house, value }) => {
        const h = _normHouse(house);
        if (map.has(h)) {
          const item = map.get(h);
          item[field].push(_norm(value));
        }
      });
    };

    // Preenche o mapa com todos os dados
    appendToMap(houseDataMap, refs, 'refs');
    appendToMap(houseDataMap, consignees, 'consignees');
    appendToMap(houseDataMap, entregas, 'entregas');
    appendToMap(houseDataMap, dtas, 'dtas');
    appendToMap(houseDataMap, previsoes, 'previsoes');
    appendToMap(houseDataMap, responsaveis, 'responsaveis');
    appendToMap(houseDataMap, observacoes, 'observacoes');

    // Gera linhas finais consolidadas (1 por HOUSE)
    const calculatedRows = [...houseDataMap.values()].map(data => ([
      data.mawb,
      data.house,
      joinVals(data.refs),
      joinVals(data.consignees),
      joinVals(data.entregas),
      joinVals(data.dtas),
      joinVals(data.previsoes),
      joinVals(data.responsaveis),
      joinVals(data.observacoes)
    ]));


    // UPSERT
    const toInsert = [];
    const toUpdateBlocks = []; // blocos contíguos para reduzir chamadas

    // Aplicar merge/insert
    const pendingUpdates = [];
    calculatedRows.forEach(r => {
      const key = keyFor(r[COLS.MAWB - 1], r[COLS.HOUSE - 1]);
      if (index.has(key)) {
        const kept = index.get(key);
        const updated = kept.row.slice();
        for (let c = COLS.REF; c <= COLS.OBS; c++) {
          updated[c - 1] = mergeCell(updated[c - 1], r[c - 1]);
        }
        pendingUpdates.push({ rowIndex: kept.rowIndex, row: updated });
        // Atualiza no índice para futuras mesclas dentro da mesma execução
        index.set(key, { rowIndex: kept.rowIndex, row: updated });
      } else {
        toInsert.push(r);
      }
    });

    // Ordena updates e faz blocos contíguos
    if (pendingUpdates.length) {
      pendingUpdates.sort((a, b) => a.rowIndex - b.rowIndex);
      let s = 0;
      while (s < pendingUpdates.length) {
        let e = s + 1;
        let startRow = pendingUpdates[s].rowIndex;
        while (e < pendingUpdates.length &&
               pendingUpdates[e].rowIndex === pendingUpdates[e - 1].rowIndex + 1) e++;
        const block = pendingUpdates.slice(s, e);
        toUpdateBlocks.push({ startRow, values: block.map(b => b.row) });
        s = e;
      }
    }

    // Escritas: updates, deletes (duplicatas antigas), inserts
    // Updates
    toUpdateBlocks.forEach(b => {
      const range = ws.getRange(b.startRow, 1, b.values.length, 9);
      // Força o formato de texto para as colunas MAWB e HOUSE antes de inserir os dados
      ws.getRange(b.startRow, 1, b.values.length, 2).setNumberFormat('@');
      range.setValues(b.values);
      const colorMap = b.values.map(row => getRowColors(row, ws.getName()));
      range.setBackgrounds(colorMap);
    });

    // Remove duplicatas antigas do mesmo MAWB/HOUSE (de baixo pra cima)
    if (duplicatesToDelete.length) {
      duplicatesToDelete.sort((a, b) => b - a).forEach(r => ws.deleteRow(r));
    }

    // Inserts com lógica de agrupamento por MAWB
    if (toInsert.length) {
      const boundaries = findMawbGroupBoundaries(ws, mawb);
      let insertRow;

      if (boundaries) {
        // MAWB existente, insere logo abaixo
        insertRow = boundaries.endRow + 1;
        ws.insertRowsAfter(boundaries.endRow, toInsert.length);
      } else {
        // Novo MAWB, encontra a última linha com dados e adiciona separador
        const lastDataRow = findLastDataRow(ws);
        insertRow = lastDataRow === 0 ? 1 : lastDataRow + 2;
      }
      const range = ws.getRange(insertRow, 1, toInsert.length, 9);
      // Força o formato de texto para as colunas MAWB e HOUSE antes de inserir os dados
      ws.getRange(insertRow, 1, toInsert.length, 2).setNumberFormat('@');
      range.setValues(toInsert);
      const colorMap = toInsert.map(row => getRowColors(row, ws.getName()));
      range.setBackgrounds(colorMap);
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
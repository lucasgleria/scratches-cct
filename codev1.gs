/** =========================
 *  BACKEND (Apps Script)
 *  =========================
 *  Lembrete: preencha a constante SPREADSHEETS com suas planilhas.
 */

/** Liste aqui as planilhas disponíveis para o usuário escolher. */
const SPREADSHEETS = [
  // Exemplo:
  { id: '1Qrzq3NatjRtLE8CiQbMiWRHvwFgUKA5ymimoR6JAsV0', name: 'CCT Teste' }
];

/** Utilitário: retorna erro padronizado */
function _asError(message, detail) {
  return { ok: false, message, detail: detail || null };
}

/** Utilitário: resposta de sucesso */
function _asOk(payload) {
  return { ok: true, payload };
}

/** Exige que exista ao menos 1 planilha configurada */
function _ensureSheetsConfigured() {
  if (!Array.isArray(SPREADSHEETS) || SPREADSHEETS.length === 0) {
    throw new Error('Nenhuma planilha configurada em SPREADSHEETS. Edite o Code.gs e adicione suas planilhas.');
  }
}

/** Publica a interface */
function doGet() {
  _ensureSheetsConfigured();
  const tpl = HtmlService.createTemplateFromFile('Index');
  tpl.appName = 'Cadastro (3ª aba)';
  tpl.spreadsheets = SPREADSHEETS; // Passa dados para o HTML
  const html = tpl.evaluate()
    .setTitle('Cadastro na 3ª aba')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  return html;
}

/** Exposto ao front: lista as planilhas configuradas (fallback se precisar) */
function getSpreadsheets() {
  try {
    _ensureSheetsConfigured();
    return _asOk(SPREADSHEETS);
  } catch (err) {
    return _asError('Falha ao carregar lista de planilhas.', err.message);
  }
}

/**
 * Salva os dados na 3ª aba da planilha selecionada.
 * Espera um objeto com a estrutura:
 * {
 *   spreadsheetId: string,
 *   mawb: string,                  // único
 *   houses: string[],              // 1..n
 *   refs: Array<{house:string, value:string}>,
 *   consignees: Array<{house:string, value:string}>,
 *   entregas: Array<{house:string, value:string}>, // value pode ser enum ou texto livre
 *   dtas: Array<{house:string, value:string}>,
 *   previsoes: Array<{house:string, value:string}>,
 *   responsaveis: Array<{house:string, value:string}>,
 *   observacoes: Array<{house:string, value:string}>
 * }
 */
function saveEntries(payload) {
  try {
    // --------- Validação básica ----------
    if (!payload || typeof payload !== 'object') {
      return _asError('Payload inválido.');
    }
    const {
      spreadsheetId,
      mawb,
      houses,
      refs = [],
      consignees = [],
      entregas = [],
      dtas = [],
      previsoes = [],
      responsaveis = [],
      observacoes = []
    } = payload;

    if (!spreadsheetId) return _asError('Selecione uma planilha.');
    if (!mawb && mawb !== '') return _asError('Campo MAWB ausente.');

    if (!Array.isArray(houses) || houses.length === 0) {
      return _asError('Adicione ao menos um HOUSE.');
    }

    // Garante strings (ou vazio) e normaliza espaços.
    const norm = v => (v == null ? '' : String(v).trim());

    const housesClean = houses.map(norm).filter(h => true); // permitimos vazio, se usuário quiser

    // Verifica se todos os itens vinculados têm HOUSE válido (um dos adicionados)
    const validHouse = new Set(housesClean);
    const checkLinked = (arr, label) => {
      for (let i = 0; i < arr.length; i++) {
        const { house } = arr[i] || {};
        if (!validHouse.has(norm(house))) {
          throw new Error(`Item de ${label} vinculado a HOUSE inexistente: "${house}" (verifique a lista de HOUSEs).`);
        }
      }
    };
    checkLinked(refs, 'REF');
    checkLinked(consignees, 'CONSIGNEE');
    checkLinked(entregas, 'ENTREGA');
    checkLinked(dtas, 'DTA');
    checkLinked(previsoes, 'PREVISÃO');
    checkLinked(responsaveis, 'RESPONSÁVEL');
    checkLinked(observacoes, 'OBSERVAÇÃO');

    // --------- Abre planilha e 3ª aba ----------
    const ss = SpreadsheetApp.openById(spreadsheetId);
    const sheets = ss.getSheets();
    if (sheets.length < 3) {
      return _asError('A planilha selecionada não possui 3 abas.');
    }
    const target = sheets[2]; // 3ª aba (index 2)

    // Cabeçalho esperado (opcionalmente você pode reforçar aqui):
    // MAWB, HOUSE, REF, CONSIGNEE, ENTREGA, DTA, PREVISÃO, RESPONSÁVEL, OBSERVAÇÃO

    // --------- Expansão em linhas -----------
    // Estratégia:
    // - Sempre que houver um item de REF/CONSIGNEE/ENTREGA/DTA/PREVISÃO/RESPONSÁVEL/OBSERVAÇÃO,
    //   geramos UMA linha para ele, com MAWB e HOUSE preenchidos e as demais colunas vazias.
    // - Se um HOUSE não tiver nenhum item associado, inserimos UMA linha apenas com MAWB e HOUSE.
    //
    // Isso evita combinações cartesianas e mantém cada “inclusão” explícita em sua linha.

    const rows = [];
    const pushRow = (houseVal, ref, consignee, entrega, dta, prev, resp, obs) => {
      rows.push([
        norm(mawb),       // MAWB
        norm(houseVal),   // HOUSE
        norm(ref),        // REF
        norm(consignee),  // CONSIGNEE
        norm(entrega),    // ENTREGA
        norm(dta),        // DTA
        norm(prev),       // PREVISÃO
        norm(resp),       // RESPONSÁVEL
        norm(obs)         // OBSERVAÇÃO
      ]);
    };

    // Índice rápido por HOUSE para saber se recebeu itens
    const hasAnyByHouse = new Map(housesClean.map(h => [h, false]));
    const mark = h => hasAnyByHouse.set(h, true);

    // Helper para iterar arrays vinculadas a HOUSE
    const expand = (arr, colName) => {
      arr.forEach(({ house, value }) => {
        const h = norm(house);
        const v = norm(value); // pode ser vazio (deixar em branco)
        mark(h);
        switch (colName) {
          case 'REF':          pushRow(h, v, '', '', '', '', '', ''); break;
          case 'CONSIGNEE':    pushRow(h, '', v, '', '', '', '', ''); break;
          case 'ENTREGA':      pushRow(h, '', '', v, '', '', '', ''); break;
          case 'DTA':          pushRow(h, '', '', '', v, '', '', ''); break;
          case 'PREV':         pushRow(h, '', '', '', '', v, '', ''); break;
          case 'RESP':         pushRow(h, '', '', '', '', '', v, ''); break;
          case 'OBS':          pushRow(h, '', '', '', '', '', '', v); break;
        }
      });
    };

    expand(refs, 'REF');
    expand(consignees, 'CONSIGNEE');
    expand(entregas, 'ENTREGA');
    expand(dtas, 'DTA');
    expand(previsoes, 'PREV');
    expand(responsaveis, 'RESP');
    expand(observacoes, 'OBS');

    // Para cada HOUSE sem itens vinculados, insere uma linha mínima (MAWB + HOUSE)
    housesClean.forEach(h => {
      if (!hasAnyByHouse.get(h)) {
        pushRow(h, '', '', '', '', '', '', '');
      }
    });

    if (rows.length === 0) {
      // Isso só ocorreria se não houvesse HOUSE; mas já validamos antes.
      return _asError('Nada para salvar.');
    }

    // --------- Append ----------
    target.getRange(target.getLastRow() + 1, 1, rows.length, 9).setValues(rows);

    return _asOk({ inserted: rows.length });

  } catch (err) {
    return _asError('Falha ao salvar os dados.', err.message);
  }
}

/** Utilitário para incluir HTML parcial se quiser (não usado, mas útil) */
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

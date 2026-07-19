// premiar-sdr.js
// Gera a planilha de premiação: Reuniões (Equipe Gabriel + Equipe Paola) + Negócios Fechados
// Uso: node --env-file=.env automacoes/premiar-sdr.js
//
// Valores dependem de senioridade (+1 ano de casa vs <1 ano):
//
// JÚNIOR (<1 ano):
//   Reunião NOVA  VC        = R$  7,50 | NOVA  Presencial = R$ 15,00
//   Reunião FUP   VC        = R$  5,00 | FUP   Presencial = R$  7,50
//   Fechamento               = R$100,00 (flat)
//
// SÊNIOR (+1 ano):
//   Reunião NOVA  VC        = R$ 10,00 | NOVA  Presencial = R$ 25,00
//   Reunião FUP   VC        = R$  5,00 | FUP   Presencial = R$ 10,00
//   Fechamento               = R$120 (1–5) / R$150 (6–10) / R$180 (11+)
//
// SDRs sênior (+1 ano): Gabriel, Ana Beatriz, Leticia

const ExcelJS = require('exceljs');
const path    = require('path');
const fs      = require('fs');

const TOKEN = process.env.AGENDOR_API_TOKEN;
if (!TOKEN) {
  console.error('AGENDOR_API_TOKEN não encontrado. Rode com: node --env-file=.env automacoes/premiar-sdr.js');
  process.exit(1);
}

const BASE         = 'https://api.agendor.com.br/v3';
const AUTH_HEADERS = { Authorization: `Token ${TOKEN}` };

// ─── SDR Map — Reuniões ───────────────────────────────────────────────────────
const SDR_MAP = {
  'Gabriel Santana Barra - Matriz SP':         'Gabriel',
  'Michelle Valente Araujo - Matriz SP':        'Michelle',
  'Kailany Santos da Silva - Matriz SP':        'Kailany',
  'Ana Karolayne - Matriz SP':                  'Ana Karolayne',
  'Igor Souza Ferreira - Matriz SP':            'Igor',
  'Larissa Santos - Filial Porto Alegre':       'Larissa',
  'Ana Belle - Filial PR':                      'Ana Belle',
  'Leticia Pereira Mascarin - Filial Campinas': 'Leticia',
  'Ana Beatriz Soares Silva - Filial Alpha':    'Ana Beatriz',
  'Lawanny Souza Brasil Luz - Filial Alpha':    'Lawanny',
  'Railanne Rosario do Carmo - Filial Alpha':   'Railanne',
};

// ─── SDR Map — Negócios Fechados ─────────────────────────────────────────────
const NEGOCIOS_SDR_MAP = {
  'Gabriel Santana Barra - Matriz SP':         'Gabriel',
  'Michelle Valente Araujo - Matriz SP':        'Michelle',
  'Kailany Santos da Silva - Matriz SP':        'Kailany',
  'Ana Karolayne - Matriz SP':                  'Ana Karolayne',
  'Igor Souza Ferreira - Matriz SP':            'Igor',
  'Larissa Santos - Filial Porto Alegre':       'Larissa',
  'Ana Belle - Filial PR':                      'Ana Belle',
  'Leticia Pereira Mascarin - Filial Campinas': 'Leticia',
  'Ana Beatriz Soares Silva - Filial Alpha':    'Ana Beatriz',
  'Lawanny Souza Brasil Luz - Filial Alpha':    'Lawanny',
  'Railanne Rosario do Carmo - Filial Alpha':   'Railanne',
  'Igor Vasconcelos - Matriz SP':               'Igor Vasconcelos',
};

const EQUIPE_GABRIEL   = ['Gabriel', 'Michelle', 'Kailany', 'Ana Karolayne', 'Igor'];
const EQUIPE_PAOLA     = ['Larissa', 'Ana Belle', 'Leticia', 'Ana Beatriz', 'Lawanny', 'Railanne'];
const SDR_ORDER        = [...EQUIPE_GABRIEL, ...EQUIPE_PAOLA];
const NEGOCIOS_ORDER   = [...SDR_ORDER, 'Igor Vasconcelos'];

// ─── Período ──────────────────────────────────────────────────────────────────
const INICIO = new Date('2026-06-22T00:00:00.000Z');
const FIM    = new Date('2026-07-18T02:59:59.000Z');

// ─── Senioridade ──────────────────────────────────────────────────────────────
// SDRs com +1 ano de casa (tabela de valores diferenciada)
const SENIOR_SDRS = new Set(['Gabriel', 'Ana Beatriz', 'Leticia']);

// ─── Tabelas de valores por reunião ──────────────────────────────────────────
const RATES = {
  senior: {
    NOVA: { 'VIDEO CONFERÊNCIA': 10.00, PRESENCIAL: 25.00 },
    FUP:  { 'VIDEO CONFERÊNCIA':  5.00, PRESENCIAL: 10.00 },
  },
  junior: {
    NOVA: { 'VIDEO CONFERÊNCIA':  7.50, PRESENCIAL: 15.00 },
    FUP:  { 'VIDEO CONFERÊNCIA':  5.00, PRESENCIAL:  7.50 },
  },
};

function calcularValorReuniao(tipo, modalidade, isSenior) {
  const tabela = isSenior ? RATES.senior : RATES.junior;
  return (tabela[tipo] || {})[modalidade] || 0;
}

// Fórmula Excel para a coluna VALOR — referencia TIPO (col D) e MODALIDADE (col E)
function formulaValorReuniao(tipoLetra, modalLetra, linha, isSenior) {
  const T = `${tipoLetra}${linha}`;
  const M = `${modalLetra}${linha}`;
  if (isSenior) {
    return `=IF(AND(${T}="NOVA",${M}="PRESENCIAL"),25,IF(AND(${T}="FUP",${M}="PRESENCIAL"),10,IF(${T}="NOVA",10,IF(${T}="FUP",5,0))))`;
  }
  return `=IF(AND(${T}="NOVA",${M}="PRESENCIAL"),15,IF(AND(${T}="FUP",${M}="PRESENCIAL"),7.5,IF(${T}="NOVA",7.5,IF(${T}="FUP",5,0))))`;
}

// ─── Comissão de fechamentos ──────────────────────────────────────────────────
// Júnior: R$100 flat  |  Sênior: blocos progressivos 120 / 150 / 180
function valorUnitPorPosicao(posicao, isSenior) {
  if (!isSenior) return 100;
  if (posicao <= 5)  return 120;
  if (posicao <= 10) return 150;
  return 180;
}

function comissaoFechamento(total, isSenior) {
  if (!isSenior) return total * 100;
  const t1 = Math.min(total, 5)                * 120;
  const t2 = Math.max(0, Math.min(total-5, 5)) * 150;
  const t3 = Math.max(0, total - 10)           * 180;
  return t1 + t2 + t3;
}

// Cor de linha por faixa de fechamento
function corFaixaFechamento(posicao, isSenior) {
  if (!isSenior) return 'FFD6E4F7'; // azul claro — júnior R$100 flat
  if (posicao <= 5)  return 'FFD6EBCD'; // verde claro   — R$120
  if (posicao <= 10) return 'FFFFF0CC'; // amarelo claro — R$150
  return 'FFFFD9B3';                    // laranja claro — R$180
}

// Fórmula Excel para VALOR do fechamento — usa posição relativa ao bloco do SDR
function formulaValorFechamento(primeiraLinha, isSenior) {
  if (!isSenior) return { formula: '=100', result: 100 };
  const pos = `ROW()-${primeiraLinha}+1`;
  return { formula: `=IF(${pos}<=5,120,IF(${pos}<=10,150,180))`, result: null };
}

// ─── Cores ────────────────────────────────────────────────────────────────────
const COR_HEADER   = 'FF1B2A4A';
const COR_EQUIPE_G = 'FF1B2A4A';
const COR_EQUIPE_P = 'FF2E4A7A';
const COR_DEALS    = 'FF1A4A2E'; // verde escuro — negócios fechados
const COR_NOVA     = 'FFD6E4F7';
const COR_FUP      = 'FFFFF3CD';
// Cores por faixa de fechamento (definidas em corFaixaFechamento())
const BRANCO       = 'FFFFFFFF';

// ─── Helpers API ──────────────────────────────────────────────────────────────
async function agendorGet(url) {
  const res = await fetch(url, { headers: AUTH_HEADERS });
  if (!res.ok) throw new Error(`HTTP ${res.status}: ${await res.text()}`);
  return res.json();
}

async function coletarOrgsComVisita(dueDateGt, dueDateLt) {
  const orgs = new Set();
  for (let page = 1; page <= 20; page++) {
    const url  = `${BASE}/tasks?per_page=100&page=${page}&dueDateGt=${dueDateGt}&dueDateLt=${dueDateLt}`;
    const json = await agendorGet(url);
    const dados = json.data || [];
    for (const t of dados) {
      if (t.type === 'Visita' && t.finishedAt && t.organization) orgs.add(t.organization.id);
    }
    if (dados.length < 100) break;
    await sleep(100);
  }
  return orgs;
}

// ─── Helpers gerais ───────────────────────────────────────────────────────────
function detectModalidade(desc) {
  const u = (desc || '').toString().toUpperCase();
  if (u.includes('PRESENCIAL')) return 'PRESENCIAL';
  return 'VIDEO CONFERÊNCIA';
}

function formatDate(d) {
  if (!d) return '';
  const dt = new Date(d);
  return `${String(dt.getUTCDate()).padStart(2,'0')}/${String(dt.getUTCMonth()+1).padStart(2,'0')}/${dt.getUTCFullYear()}`;
}

function extrairArea(negocio) {
  if (!negocio) return '';
  const match = negocio.toString().match(/\/\s*(.+)$/);
  return match ? match[1].trim() : '';
}

function sleep(ms) { return new Promise(r => setTimeout(r, ms)); }

// ─── Helpers Excel ────────────────────────────────────────────────────────────
function aplicarEstiloHeader(row, corFundo = COR_HEADER) {
  row.eachCell(cell => {
    cell.font      = { bold: true, color: { argb: BRANCO }, name: 'Arial Narrow', size: 10 };
    cell.fill      = { type: 'pattern', pattern: 'solid', fgColor: { argb: corFundo } };
    cell.alignment = { horizontal: 'center', vertical: 'middle' };
  });
  row.height = 22;
}

const COL = { DATA: 1, EMPRESA: 2, AREA: 3, TIPO: 4, MODAL: 5, VALOR: 6, LINK: 7 };
const LETRA = { 1:'A', 2:'B', 3:'C', 4:'D', 5:'E', 6:'F', 7:'G' };

// ─── Ler negócios fechados ────────────────────────────────────────────────────
async function lerNegociosFechados() {
  const wb = new ExcelJS.Workbook();
  await wb.xlsx.readFile('dados/Negocios ganho junho - julho.xlsx');
  const ws = wb.worksheets[0];

  const hdrMap = {};
  ws.getRow(1).eachCell((cell, col) => { hdrMap[cell.value] = col; });
  const C = {
    usuario:      hdrMap['Usuário responsável'],
    empresa:      hdrMap['Empresa relacionada'],
    titulo:       hdrMap['Título do negócio'],
    produto:      hdrMap['Produto'],
    status:       hdrMap['Status'],
    dataConcl:    hdrMap['Data de conclusão'],
    codEmpresa:   hdrMap['Código da Empresa'],
    valor:        hdrMap['Valor'],
  };

  const negocios = [];
  ws.eachRow((row, rowNum) => {
    if (rowNum === 1) return;
    const usuario = row.getCell(C.usuario).value;
    const sdr     = NEGOCIOS_SDR_MAP[usuario];
    if (!sdr) return; // ignorar usuários fora do mapa (ex: Victor Bernardes)
    if (row.getCell(C.status).value !== 'Ganho') return;

    const dataRaw = row.getCell(C.dataConcl).value;
    negocios.push({
      sdr,
      empresa:    row.getCell(C.empresa).value || '',
      titulo:     row.getCell(C.titulo).value  || '',
      produto:    row.getCell(C.produto).value  || '',
      data:       dataRaw ? formatDate(dataRaw) : '',
      orgId:      row.getCell(C.codEmpresa).value ? Number(row.getCell(C.codEmpresa).value) : null,
    });
  });

  // Agrupar por SDR
  const porSdr = {};
  for (const n of negocios) {
    if (!porSdr[n.sdr]) porSdr[n.sdr] = [];
    porSdr[n.sdr].push(n);
  }
  return porSdr;
}

// ─── Gerar aba de Negócios Fechados ─────────────────────────────────────────
// Colunas: A=DATA  B=EMPRESA  C=PRODUTO/SERVIÇO  D=VALOR  E=LINK AGENDOR
//
// VALOR usa fórmula por posição dentro do bloco de cada SDR:
//   =IF(posição<=5, 120, IF(posição<=10, 150, 180))
// Cor de fundo por faixa: verde=R$120 | amarelo=R$150 | laranja=R$180
function gerarAbaNegociosFechados(outWb, porSdr, sdrsParaMostrar, tituloAba, corCabecalho) {
  const sheet = outWb.addWorksheet(tituloAba);

  // Título
  sheet.mergeCells('A1:E1');
  const tc = sheet.getCell('A1');
  tc.value     = 'NEGÓCIOS FECHADOS — GRUPO VIGNA — MAIO/JUNHO 2026';
  tc.font      = { bold: true, size: 12, color: { argb: BRANCO }, name: 'Arial Narrow' };
  tc.fill      = { type: 'pattern', pattern: 'solid', fgColor: { argb: corCabecalho } };
  tc.alignment = { horizontal: 'center', vertical: 'middle' };
  sheet.getRow(1).height = 28;

  // Legenda de faixas
  sheet.addRow([]);
  const legRow = sheet.addRow([
    '🟦 Júnior (<1 ano): R$100/cada (flat)',
    '🟩 Sênior (+1 ano): R$120 (1–5)',
    '🟨 Sênior (+1 ano): R$150 (6–10)',
    '🟧 Sênior (+1 ano): R$180 (11+)',
  ]);
  legRow.getCell(1).font = { bold: true, size: 9, color: { argb: 'FF1B2A4A' } };
  legRow.getCell(2).font = { bold: true, size: 9, color: { argb: 'FF2E6B30' } };
  legRow.getCell(3).font = { bold: true, size: 9, color: { argb: 'FF7A6000' } };
  legRow.getCell(4).font = { bold: true, size: 9, color: { argb: 'FF8B3A00' } };
  legRow.height = 14;
  sheet.addRow([]);

  sheet.getColumn(1).width = 14;
  sheet.getColumn(2).width = 40;
  sheet.getColumn(3).width = 28;
  sheet.getColumn(4).width = 14;
  sheet.getColumn(5).width = 50;

  const comissoesPorSdr = {};

  for (const sdrName of sdrsParaMostrar) {
    const isSeniorSdr   = SENIOR_SDRS.has(sdrName);
    const deals         = porSdr[sdrName] || [];
    const total         = deals.length;
    const totalComissao = comissaoFechamento(total, isSeniorSdr);
    comissoesPorSdr[sdrName] = { total, totalComissao };

    // Subheader do SDR (com nível)
    const nivelLabel = isSeniorSdr ? '+1 ano' : '<1 ano';
    const sdrRow = sheet.addRow([`${sdrName.toUpperCase()}  (${nivelLabel})`]);
    sdrRow.getCell(1).font  = { bold: true, color: { argb: BRANCO }, name: 'Arial Narrow', size: 10 };
    sdrRow.getCell(1).fill  = { type: 'pattern', pattern: 'solid', fgColor: { argb: corCabecalho } };
    sdrRow.height = 18;
    sheet.mergeCells(`A${sdrRow.number}:E${sdrRow.number}`);

    // Nota de valores do SDR
    const notaSdr = isSeniorSdr
      ? sheet.addRow(['', '  +1 ano: R$120 (1–5) | R$150 (6–10) | R$180 (11+)'])
      : sheet.addRow(['', '  <1 ano: R$100 por fechamento (flat)']);
    notaSdr.getCell(2).font = { italic: true, size: 9, color: { argb: 'FF555555' } };
    notaSdr.height = 13;

    // Header das colunas
    const hRow = sheet.addRow(['DATA FECHAMENTO', 'EMPRESA', 'PRODUTO / SERVIÇO', 'VALOR', 'LINK AGENDOR']);
    aplicarEstiloHeader(hRow, COR_DEALS);

    if (deals.length === 0) {
      const emRow = sheet.addRow(['', '(sem negócios no período)']);
      emRow.getCell(2).font = { italic: true, color: { argb: 'FF888888' } };
      sheet.addRow([]);
      continue;
    }

    const primeiraLinha = sheet.rowCount + 1;

    deals.forEach((d, idx) => {
      const posicao   = idx + 1;
      const valorUnit = valorUnitPorPosicao(posicao, isSeniorSdr);
      const link      = d.orgId ? `https://beta.agendor.com.br/tasks?organizationId=${d.orgId}` : '';
      const fmla      = formulaValorFechamento(primeiraLinha, isSeniorSdr);

      const row = sheet.addRow([d.data, d.empresa, d.produto, null, link]);
      row.getCell(4).value = { formula: fmla.formula, result: valorUnit };

      const bg = corFaixaFechamento(posicao, isSeniorSdr);
      row.eachCell(cell => {
        cell.fill      = { type: 'pattern', pattern: 'solid', fgColor: { argb: bg } };
        cell.alignment = { vertical: 'middle' };
        cell.border    = { bottom: { style: 'thin', color: { argb: 'FFD0D0D0' } } };
      });
      row.getCell(1).alignment = { horizontal: 'center', vertical: 'middle' };
      row.getCell(4).numFmt    = 'R$ #,##0.00';
      row.getCell(4).alignment = { horizontal: 'right', vertical: 'middle' };
      row.height = 18;
    });

    const ultimaLinha = sheet.rowCount;

    // Total com SUM
    const totRow = sheet.addRow(['', `TOTAL: ${total} fechamento(s)`, '', null, '']);
    totRow.getCell(4).value = {
      formula: `=SUM(D${primeiraLinha}:D${ultimaLinha})`,
      result:  totalComissao,
    };
    totRow.getCell(2).font      = { bold: true, name: 'Arial Narrow' };
    totRow.getCell(4).numFmt    = 'R$ #,##0.00';
    totRow.getCell(4).font      = { bold: true, name: 'Arial Narrow' };
    totRow.getCell(4).alignment = { horizontal: 'right' };
    totRow.height = 18;

    sheet.addRow([]); // espaçamento entre SDRs
  }

  // Total geral
  const gtTotal = Object.values(comissoesPorSdr).reduce((s, d) => s + d.totalComissao, 0);
  sheet.addRow([]);
  const gtRow = sheet.addRow(['', 'TOTAL GERAL — NEGÓCIOS FECHADOS', '', gtTotal, '']);
  gtRow.getCell(2).font      = { bold: true, size: 11, name: 'Arial Narrow' };
  gtRow.getCell(4).numFmt    = 'R$ #,##0.00';
  gtRow.getCell(4).font      = { bold: true, size: 11, name: 'Arial Narrow' };
  gtRow.getCell(4).alignment = { horizontal: 'right' };
  gtRow.height = 22;

  return comissoesPorSdr;
}

// ─── MAIN ────────────────────────────────────────────────────────────────────
async function main() {
  console.log('=== PREMIAÇÃO SDR — GRUPO VIGNA — 22/06 → 17/07/2026 ===\n');

  // ── 1. Reuniões: ler planilha ──────────────────────────────────────────────
  const wb = new ExcelJS.Workbook();
  await wb.xlsx.readFile('dados/Reunião realizadas junho - julho.xlsx');
  const wsSource = wb.worksheets[0];

  const hdrMap = {};
  wsSource.getRow(1).eachCell((cell, col) => { hdrMap[cell.value] = col; });
  const C = {
    usuario:       hdrMap['Usuário que realizou a tarefa'],
    empresa:       hdrMap['Empresa relacionada'],
    codigoEmpresa: hdrMap['Código da Empresa'],
    negocio:       hdrMap['Negócio relacionado'],
    dataAgend:     hdrMap['Data de agendamento'],
    descricao:     hdrMap['Descrição'],
  };

  const reunioes  = [];
  const excluidos = [];

  wsSource.eachRow((row, rowNum) => {
    if (rowNum === 1) return;
    const usuario = row.getCell(C.usuario).value;
    const sdr     = SDR_MAP[usuario];
    if (!sdr) return;

    const dataRaw = row.getCell(C.dataAgend).value;
    if (!dataRaw) return;
    const dt = new Date(dataRaw);

    if (dt < INICIO || dt > FIM) {
      excluidos.push({ rowNum, sdr, data: formatDate(dt), empresa: row.getCell(C.empresa).value || '(sem empresa)' });
      return;
    }

    const orgId = row.getCell(C.codigoEmpresa).value;
    reunioes.push({
      rowNum, sdr,
      orgId: orgId ? Number(orgId) : null,
      empresa:       row.getCell(C.empresa).value || '',
      area:          extrairArea(row.getCell(C.negocio).value),
      dataAgendamento: dt,
      dataFormatada:   formatDate(dt),
      modalidade:      detectModalidade(row.getCell(C.descricao).value),
      tipo: null, valor: null,
    });
  });

  console.log(`Reuniões no período: ${reunioes.length} (${SDR_ORDER.length} SDRs)`);

  // ── 2. Histórico Agendor (NOVA vs FUP) ────────────────────────────────────
  console.log('\nBuscando histórico de visitas (fev–mai/2026)...');
  const lotes = [['2026-03-22','2026-04-22'],['2026-04-22','2026-05-22'],['2026-05-22','2026-06-22']];
  const orgsComHistorico = new Set();
  for (const [gt, lt] of lotes) {
    process.stdout.write(`  ${gt} → ${lt}...`);
    const orgs = await coletarOrgsComVisita(gt, lt);
    orgs.forEach(id => orgsComHistorico.add(id));
    console.log(` ${orgs.size} orgs`);
    await sleep(300);
  }
  const nossoOrgIds   = new Set(reunioes.filter(r => r.orgId).map(r => r.orgId));
  const orgsRelevantes = new Set([...orgsComHistorico].filter(id => nossoOrgIds.has(id)));
  console.log(`${orgsRelevantes.size} empresa(s) com visita anterior → FUP automático.`);

  // ── 3. Classificar NOVA / FUP ─────────────────────────────────────────────
  const sortedPorData = [...reunioes].sort((a, b) => a.dataAgendamento - b.dataAgendamento);
  const primeiraOcorrencia = new Map();
  for (const r of sortedPorData) {
    if (!r.orgId) continue;
    const key = `${r.sdr}:${r.orgId}`;
    if (!primeiraOcorrencia.has(key)) primeiraOcorrencia.set(key, r.dataAgendamento.getTime());
  }
  for (const r of reunioes) {
    if (!r.orgId || orgsRelevantes.has(r.orgId)) {
      r.tipo = 'FUP';
    } else {
      const key = `${r.sdr}:${r.orgId}`;
      r.tipo = (r.dataAgendamento.getTime() === primeiraOcorrencia.get(key)) ? 'NOVA' : 'FUP';
    }
    r.valor = calcularValorReuniao(r.tipo, r.modalidade, SENIOR_SDRS.has(r.sdr));
  }

  // ── 4. Ler negócios fechados ───────────────────────────────────────────────
  console.log('\nLendo negócios fechados...');
  const negociosPorSdr = await lerNegociosFechados();
  for (const [sdr, deals] of Object.entries(negociosPorSdr)) {
    const isSr = SENIOR_SDRS.has(sdr);
  console.log(`  ${sdr} (${isSr?'sênior':'júnior'}): ${deals.length} fechamento(s) → R$ ${comissaoFechamento(deals.length, isSr).toFixed(2).replace('.',',')}`);
  }

  // ── 5. Gerar Excel ────────────────────────────────────────────────────────
  const outWb = new ExcelJS.Workbook();
  const summaryData   = {};
  const summaryDeals  = {};

  // Criar Resumo primeiro (fica na 1ª posição)
  const capa = outWb.addWorksheet('Resumo');

  // ── Aba Negócios Fechados (geral — todas as equipes) ──────────────────────
  const comissoesFechamentos = gerarAbaNegociosFechados(
    outWb, negociosPorSdr, NEGOCIOS_ORDER, 'Negócios Fechados', COR_DEALS
  );
  for (const [sdr, d] of Object.entries(comissoesFechamentos)) summaryDeals[sdr] = d;

  // ── Aba Igor Vasconcelos (dedicada) ──────────────────────────────────────
  const ivDeals    = negociosPorSdr['Igor Vasconcelos'] || [];
  const ivTot      = ivDeals.length;
  const ivSenior   = SENIOR_SDRS.has('Igor Vasconcelos'); // false — júnior
  const ivComiss   = comissaoFechamento(ivTot, ivSenior);
  const ivSheet  = outWb.addWorksheet('Igor Vasconcelos');

  ivSheet.mergeCells('A1:E1');
  const ivTitle = ivSheet.getCell('A1');
  ivTitle.value     = 'IGOR VASCONCELOS — NEGÓCIOS FECHADOS — MAIO/JUNHO 2026';
  ivTitle.font      = { bold: true, size: 12, color: { argb: BRANCO }, name: 'Arial Narrow' };
  ivTitle.fill      = { type: 'pattern', pattern: 'solid', fgColor: { argb: COR_DEALS } };
  ivTitle.alignment = { horizontal: 'center', vertical: 'middle' };
  ivSheet.getRow(1).height = 28;
  ivSheet.addRow([]);
  const ivLeg = ivSheet.addRow(['⬛ 1–5: R$120/cada', '⬛ 6–10: R$150/cada', '⬛ 11+: R$180/cada']);
  ivLeg.getCell(1).font = { bold: true, size: 9, color: { argb: 'FF2E6B30' } };
  ivLeg.getCell(2).font = { bold: true, size: 9, color: { argb: 'FF7A6000' } };
  ivLeg.getCell(3).font = { bold: true, size: 9, color: { argb: 'FF8B3A00' } };
  ivLeg.height = 14;
  ivSheet.addRow([]);
  ivSheet.getColumn(1).width = 14; ivSheet.getColumn(2).width = 40;
  ivSheet.getColumn(3).width = 28; ivSheet.getColumn(4).width = 14;
  ivSheet.getColumn(5).width = 50;

  const ivHdr = ivSheet.addRow(['DATA FECHAMENTO', 'EMPRESA', 'PRODUTO / SERVIÇO', 'VALOR', 'LINK AGENDOR']);
  aplicarEstiloHeader(ivHdr, COR_DEALS);

  const ivPrimeiraLinha = ivSheet.rowCount + 1;
  ivDeals.forEach((d, idx) => {
    const posicao   = idx + 1;
    const valorUnit = valorUnitPorPosicao(posicao, ivSenior);
    const link      = d.orgId ? `https://beta.agendor.com.br/tasks?organizationId=${d.orgId}` : '';
    const fmla      = formulaValorFechamento(ivPrimeiraLinha, ivSenior);
    const row = ivSheet.addRow([d.data, d.empresa, d.produto, null, link]);
    row.getCell(4).value = { formula: fmla.formula, result: valorUnit };
    const bg = corFaixaFechamento(posicao, ivSenior);
    row.eachCell(cell => {
      cell.fill      = { type: 'pattern', pattern: 'solid', fgColor: { argb: bg } };
      cell.alignment = { vertical: 'middle' };
      cell.border    = { bottom: { style: 'thin', color: { argb: 'FFD0D0D0' } } };
    });
    row.getCell(1).alignment = { horizontal: 'center', vertical: 'middle' };
    row.getCell(4).numFmt    = 'R$ #,##0.00';
    row.getCell(4).alignment = { horizontal: 'right', vertical: 'middle' };
    row.height = 18;
  });
  const ivUltimaLinha = ivSheet.rowCount;

  ivSheet.addRow([]);
  const ivResRow = ivSheet.addRow(['', `TOTAL: ${ivTot} fechamento(s)`, '', null, '']);
  ivResRow.getCell(4).value = {
    formula: `=SUM(D${ivPrimeiraLinha}:D${ivUltimaLinha})`,
    result:  ivComiss,
  };
  ivResRow.getCell(2).font      = { bold: true, name: 'Arial Narrow' };
  ivResRow.getCell(4).numFmt    = 'R$ #,##0.00';
  ivResRow.getCell(4).font      = { bold: true, name: 'Arial Narrow' };
  ivResRow.getCell(4).alignment = { horizontal: 'right' };

  // ── Abas individuais de reuniões por SDR ─────────────────────────────────
  for (const sdrName of SDR_ORDER) {
    const sdrReunioes = reunioes
      .filter(r => r.sdr === sdrName)
      .sort((a, b) => a.dataAgendamento - b.dataAgendamento);

    const isSenior = SENIOR_SDRS.has(sdrName);
    const sheet    = outWb.addWorksheet(sdrName);
    const equipe   = EQUIPE_GABRIEL.includes(sdrName) ? 'Equipe Gabriel' : 'Equipe Paola';
    const nivel    = isSenior ? '+1 ano' : '<1 ano';

    sheet.mergeCells('A1:G1');
    const titleCell = sheet.getCell('A1');
    titleCell.value     = `PREMIAÇÃO — ${sdrName.toUpperCase()} (${equipe} | ${nivel}) — 22/06 → 17/07/2026`;
    titleCell.font      = { bold: true, size: 11, color: { argb: BRANCO }, name: 'Arial Narrow' };
    titleCell.fill      = { type: 'pattern', pattern: 'solid', fgColor: { argb: COR_HEADER } };
    titleCell.alignment = { horizontal: 'center', vertical: 'middle' };
    sheet.getRow(1).height = 28;

    // Linha de tabela de valores do SDR
    sheet.addRow([]);
    const tabelaRow = isSenior
      ? sheet.addRow(['', 'NOVA VC: R$10 | NOVA Presencial: R$25 | FUP VC: R$5 | FUP Presencial: R$10'])
      : sheet.addRow(['', 'NOVA VC: R$7,50 | NOVA Presencial: R$15 | FUP VC: R$5 | FUP Presencial: R$7,50']);
    tabelaRow.getCell(2).font = { italic: true, size: 9, color: { argb: 'FF1B2A4A' }, bold: true };
    tabelaRow.height = 13;

    const headerRow = sheet.addRow(['DATA', 'EMPRESA', 'ÁREA', 'TIPO', 'MODALIDADE', 'VALOR', 'LINK AGENDOR']);
    aplicarEstiloHeader(headerRow);

    sheet.getColumn(1).width = 12; sheet.getColumn(2).width = 38;
    sheet.getColumn(3).width = 22; sheet.getColumn(4).width = 8;
    sheet.getColumn(5).width = 20; sheet.getColumn(6).width = 14;
    sheet.getColumn(7).width = 50;

    const notaRow = sheet.addRow(['', '⚡ Altere TIPO e MODALIDADE → VALOR recalcula automaticamente']);
    notaRow.getCell(2).font = { italic: true, size: 9, color: { argb: 'FF555555' } };
    notaRow.height = 13;

    const PRIMEIRA_LINHA = 6; // título(1) + branco(2) + tabela(3) + header(4) + nota(5) → dados a partir de 6
    let totalNovas = 0, totalFups = 0, totalValor = 0, linhaAtual = PRIMEIRA_LINHA;

    for (const r of sdrReunioes) {
      const link = r.orgId ? `https://beta.agendor.com.br/tasks?organizationId=${r.orgId}` : '';
      const row  = sheet.addRow([r.dataFormatada, r.empresa, r.area, r.tipo, r.modalidade, null, link]);

      // Fórmula VALOR: considera TIPO (col D) + MODALIDADE (col E) + nível do SDR
      row.getCell(COL.VALOR).value = {
        formula: formulaValorReuniao(LETRA[COL.TIPO], LETRA[COL.MODAL], linhaAtual, isSenior),
        result:  r.valor,
      };

      // Dropdown TIPO
      row.getCell(COL.TIPO).dataValidation = {
        type: 'list', allowBlank: false, showDropDown: false, formulae: ['"NOVA,FUP"'],
      };
      // Dropdown MODALIDADE
      row.getCell(COL.MODAL).dataValidation = {
        type: 'list', allowBlank: false, showDropDown: false,
        formulae: ['"VIDEO CONFERÊNCIA,PRESENCIAL"'],
      };

      const bg = r.tipo === 'NOVA' ? COR_NOVA : COR_FUP;
      row.eachCell(cell => {
        cell.fill      = { type: 'pattern', pattern: 'solid', fgColor: { argb: bg } };
        cell.alignment = { vertical: 'middle' };
        cell.border    = { bottom: { style: 'thin', color: { argb: 'FFD0D0D0' } } };
      });
      row.getCell(COL.TIPO).alignment  = { horizontal: 'center', vertical: 'middle' };
      row.getCell(COL.MODAL).alignment = { horizontal: 'center', vertical: 'middle' };
      row.getCell(COL.VALOR).numFmt    = 'R$ #,##0.00';
      row.getCell(COL.VALOR).alignment = { horizontal: 'right', vertical: 'middle' };
      row.height = 18;

      if (r.tipo === 'NOVA') totalNovas++; else totalFups++;
      totalValor += r.valor;
      linhaAtual++;
    }

    sheet.addRow([]);
    linhaAtual++;
    const totalRow = sheet.addRow(['', 'TOTAL', '', `${totalNovas} N / ${totalFups} F`, '', null]);
    totalRow.getCell(COL.VALOR).value = {
      formula: `=SUM(F${PRIMEIRA_LINHA}:F${linhaAtual - 1})`, result: totalValor,
    };
    totalRow.getCell(COL.EMPRESA).font   = { bold: true, name: 'Arial Narrow' };
    totalRow.getCell(COL.TIPO).font      = { bold: true, name: 'Arial Narrow' };
    totalRow.getCell(COL.TIPO).alignment = { horizontal: 'center' };
    totalRow.getCell(COL.VALOR).numFmt   = 'R$ #,##0.00';
    totalRow.getCell(COL.VALOR).font     = { bold: true };
    totalRow.getCell(COL.VALOR).alignment = { horizontal: 'right' };

    summaryData[sdrName] = { novas: totalNovas, fups: totalFups, total: sdrReunioes.length, valor: totalValor };
  }

  // ── Resumo (Capa) ─────────────────────────────────────────────────────────
  capa.mergeCells('B2:J2');
  const capaTitle = capa.getCell('B2');
  capaTitle.value     = 'PREMIAÇÃO SDR — GRUPO VIGNA — 22/06 → 17/07/2026';
  capaTitle.font      = { bold: true, size: 14, color: { argb: BRANCO }, name: 'Arial Narrow' };
  capaTitle.fill      = { type: 'pattern', pattern: 'solid', fgColor: { argb: COR_HEADER } };
  capaTitle.alignment = { horizontal: 'center', vertical: 'middle' };
  capa.getRow(2).height = 36;
  capa.addRow([]);

  // Colunas do resumo
  const COLS_RES = ['', 'SDR', 'NOVAS', 'FUPs', 'TOTAL REUNIÕES',
                    'COMISSÃO REUNIÕES', 'FECHAMENTOS', 'VALOR UNIT.', 'COMISSÃO FECHAMENTOS', 'TOTAL GERAL'];

  function bloco(sdrs, label, cor) {
    const hdr = capa.addRow(['', `— ${label} —`]);
    hdr.getCell(2).font = { bold: true, color: { argb: BRANCO }, name: 'Arial Narrow', size: 10 };
    hdr.getCell(2).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: cor } };
    capa.mergeCells(`B${hdr.number}:J${hdr.number}`);
    hdr.height = 18;

    const sh = capa.addRow(COLS_RES);
    aplicarEstiloHeader(sh, cor);

    let totR = 0, totF = 0, totT = 0, totCV = 0, totFV = 0;
    for (const sdrName of sdrs) {
      const m  = summaryData[sdrName]  || { novas:0, fups:0, total:0, valor:0 };
      const fd = summaryDeals[sdrName] || { total:0, totalComissao:0 };
      const totalGeral = m.valor + fd.totalComissao;
      const row = capa.addRow(['', sdrName, m.novas, m.fups, m.total,
                                m.valor, fd.total, '(blocos)', fd.totalComissao, totalGeral]);
      row.getCell(2).font = { name: 'Arial Narrow' };
      [6, 8, 9, 10].forEach(i => { row.getCell(i).numFmt = 'R$ #,##0.00'; row.getCell(i).alignment = { horizontal: 'right' }; });
      row.getCell(7).alignment = { horizontal: 'center' };
      row.getCell(10).font = { bold: true };
      row.height = 18;
      totR += m.valor; totF += fd.totalComissao; totCV += m.total; totFV += fd.total; totT += totalGeral;
    }

    const stRow = capa.addRow(['', `SUBTOTAL ${label.split(' ')[1]}`, '', '', '', totR, totFV, '', totF, totT]);
    stRow.getCell(2).font = { bold: true, name: 'Arial Narrow' };
    [6, 9, 10].forEach(i => { stRow.getCell(i).numFmt = 'R$ #,##0.00'; stRow.getCell(i).alignment = { horizontal: 'right' }; stRow.getCell(i).font = { bold: true }; });
    stRow.height = 18;
    capa.addRow([]);
    return { totR, totF, totT };
  }

  const bg = bloco(EQUIPE_GABRIEL, 'EQUIPE GABRIEL', COR_EQUIPE_G);
  const bp = bloco(EQUIPE_PAOLA,   'EQUIPE PAOLA',   COR_EQUIPE_P);

  // Igor Vasconcelos (linha separada — só fechamentos)
  const ivComissaoR = comissaoFechamento(ivTot);
  const ivRow = capa.addRow(['', 'Igor Vasconcelos', '', '', '', 0, ivTot, '(blocos)', ivComissaoR, ivComissaoR]);
  ivRow.getCell(2).font = { italic: true, name: 'Arial Narrow' };
  [6, 8, 9, 10].forEach(i => { ivRow.getCell(i).numFmt = 'R$ #,##0.00'; ivRow.getCell(i).alignment = { horizontal: 'right' }; });
  ivRow.getCell(7).alignment = { horizontal: 'center' };
  capa.addRow([]);

  // Total Geral
  const gtTotal = bg.totT + bp.totT + ivComissaoR;
  const gtRow   = capa.addRow(['', 'TOTAL GERAL', '', '', '', '', '', '', '', gtTotal]);
  gtRow.getCell(2).font = { bold: true, size: 12, name: 'Arial Narrow' };
  gtRow.getCell(10).numFmt = 'R$ #,##0.00';
  gtRow.getCell(10).font   = { bold: true, size: 12 };
  gtRow.getCell(10).alignment = { horizontal: 'right' };
  gtRow.height = 24;

  // Legenda de valores
  capa.addRow([]); capa.addRow([]);
  const legHdr = capa.addRow(['', 'TABELA DE VALORES']);
  legHdr.getCell(2).font = { bold: true, name: 'Arial Narrow' };
  capa.addRow(['', '── JÚNIOR (<1 ano) ──────────────────────────']);
  capa.addRow(['', 'REUNIÃO NOVA  VC', '', '', '', 'R$ 7,50']);
  capa.addRow(['', 'REUNIÃO NOVA  Presencial', '', '', '', 'R$ 15,00']);
  capa.addRow(['', 'REUNIÃO FUP   VC', '', '', '', 'R$ 5,00']);
  capa.addRow(['', 'REUNIÃO FUP   Presencial', '', '', '', 'R$ 7,50']);
  capa.addRow(['', 'FECHAMENTO (flat)', '', '', '', 'R$ 100,00']);
  capa.addRow(['', '── SÊNIOR (+1 ano) ─────────────────────────']);
  capa.addRow(['', 'REUNIÃO NOVA  VC', '', '', '', 'R$ 10,00']);
  capa.addRow(['', 'REUNIÃO NOVA  Presencial', '', '', '', 'R$ 25,00']);
  capa.addRow(['', 'REUNIÃO FUP   VC', '', '', '', 'R$ 5,00']);
  capa.addRow(['', 'REUNIÃO FUP   Presencial', '', '', '', 'R$ 10,00']);
  capa.addRow(['', 'FECHAMENTO 1–5', '', '', '', 'R$ 120,00/cada']);
  capa.addRow(['', 'FECHAMENTO 6–10', '', '', '', 'R$ 150,00/cada']);
  capa.addRow(['', 'FECHAMENTO 11+', '', '', '', 'R$ 180,00/cada']);
  capa.addRow(['', 'SDRs sênior (+1 ano): Gabriel, Ana Beatriz, Leticia']);

  // Larguras da capa
  capa.getColumn(1).width  = 3;
  capa.getColumn(2).width  = 25;
  capa.getColumn(3).width  = 8;
  capa.getColumn(4).width  = 8;
  capa.getColumn(5).width  = 14;
  capa.getColumn(6).width  = 18;
  capa.getColumn(7).width  = 13;
  capa.getColumn(8).width  = 13;
  capa.getColumn(9).width  = 20;
  capa.getColumn(10).width = 16;

  // ── Salvar ────────────────────────────────────────────────────────────────
  const outDir  = path.join('relatorios', 'premiacao');
  fs.mkdirSync(outDir, { recursive: true });
  const outPath = path.join(outDir, 'Premiacao_SDR_JUNHO_JULHO_2026.xlsx');
  await outWb.xlsx.writeFile(outPath);

  console.log(`\n✓ Arquivo gerado: ${outPath}`);
  console.log('\n=== RESULTADO FINAL ===');
  console.log('\nEQUIPE GABRIEL:');
  for (const sdr of EQUIPE_GABRIEL) {
    const m  = summaryData[sdr]  || { novas:0, fups:0, total:0, valor:0 };
    const fd = summaryDeals[sdr] || { total:0, totalComissao:0 };
    console.log(`  ${sdr}: reuniões R$ ${m.valor.toFixed(2).replace('.',',')} + fechamentos R$ ${fd.totalComissao.toFixed(2).replace('.',',')} = R$ ${(m.valor+fd.totalComissao).toFixed(2).replace('.',',')}`);
  }
  console.log('\nEQUIPE PAOLA:');
  for (const sdr of EQUIPE_PAOLA) {
    const m  = summaryData[sdr]  || { novas:0, fups:0, total:0, valor:0 };
    const fd = summaryDeals[sdr] || { total:0, totalComissao:0 };
    console.log(`  ${sdr}: reuniões R$ ${m.valor.toFixed(2).replace('.',',')} + fechamentos R$ ${fd.totalComissao.toFixed(2).replace('.',',')} = R$ ${(m.valor+fd.totalComissao).toFixed(2).replace('.',',')}`);
  }
  console.log(`  Igor Vasconcelos: fechamentos R$ ${ivComissaoR.toFixed(2).replace('.',',')} = R$ ${ivComissaoR.toFixed(2).replace('.',',')}`);
  console.log(`\nTOTAL GERAL: R$ ${gtTotal.toFixed(2).replace('.',',')}`);
  if (excluidos.length > 0) console.log(`\nObs.: ${excluidos.length} registro(s) excluído(s) por fora do período.`);
}

main().catch(e => { console.error('ERRO:', e.message); process.exit(1); });

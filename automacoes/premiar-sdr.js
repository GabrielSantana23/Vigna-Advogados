// premiar-sdr.js
// Gera a planilha de premiação de todos os SDRs (Equipe Gabriel + Equipe Paola)
// Uso: node --env-file=.env automacoes/premiar-sdr.js
//
// Lógica NOVA vs FUP:
//   - Busca globalmente Visitas finalizadas nos 3 meses antes do período (lotes de 31 dias)
//   - Se a empresa já teve visita antes de 21/05 → FUP
//   - Se não → primeira ocorrência no período = NOVA, demais = FUP
//
// Valores:
//   - NOVA = R$ 7,50  |  FUP = R$ 5,00  (flat, coluna VALOR é fórmula Excel auto-atualizável)

const ExcelJS = require('exceljs');
const path = require('path');
const fs = require('fs');

const TOKEN = process.env.AGENDOR_API_TOKEN;
if (!TOKEN) {
  console.error('AGENDOR_API_TOKEN não encontrado. Rode com: node --env-file=.env automacoes/premiar-sdr.js');
  process.exit(1);
}

const BASE = 'https://api.agendor.com.br/v3';
const AUTH_HEADERS = { Authorization: `Token ${TOKEN}` };

// ─── Mapeamento de usuários do Agendor → nome curto ───────────────────────────
const SDR_MAP = {
  // Equipe Gabriel
  'Gabriel Santana Barra - Matriz SP':        'Gabriel',
  'Michelle Valente Araujo - Matriz SP':       'Michelle',
  'Kailany Santos da Silva - Matriz SP':       'Kailany',
  'Ana Karolayne - Matriz SP':                 'Ana Karolayne',
  'Igor Souza Ferreira - Matriz SP':           'Igor',
  // Equipe Paola
  'Larissa Santos - Filial Porto Alegre':      'Larissa',
  'Ana Belle - Filial PR':                     'Ana Belle',
  'Leticia Pereira Mascarin - Filial Campinas':'Leticia',
  'Ana Beatriz Soares Silva - Filial Alpha':   'Ana Beatriz',
  'Lawanny Souza Brasil Luz - Filial Alpha':   'Lawanny',
  'Railanne Rosario do Carmo - Filial Alpha':  'Railanne',
};

const EQUIPE_GABRIEL = ['Gabriel', 'Michelle', 'Kailany', 'Ana Karolayne', 'Igor'];
const EQUIPE_PAOLA   = ['Larissa', 'Ana Belle', 'Leticia', 'Ana Beatriz', 'Lawanny', 'Railanne'];
const SDR_ORDER      = [...EQUIPE_GABRIEL, ...EQUIPE_PAOLA];

// ─── Período e valores ─────────────────────────────────────────────────────────
const INICIO = new Date('2026-05-21T00:00:00.000Z');
const FIM    = new Date('2026-06-20T02:59:59.000Z'); // 19/06 23:59 BRT

const VALOR_NOVA = 7.50;
const VALOR_FUP  = 5.00;

// ─── Cores ────────────────────────────────────────────────────────────────────
const COR_HEADER       = 'FF1B2A4A'; // azul escuro Vigna
const COR_EQUIPE_G     = 'FF1B2A4A'; // azul escuro — equipe Gabriel
const COR_EQUIPE_P     = 'FF2E4A7A'; // azul médio — equipe Paola
const COR_NOVA         = 'FFD6E4F7'; // azul claro
const COR_FUP          = 'FFFFF3CD'; // amarelo claro
const BRANCO           = 'FFFFFFFF';

// ─── Helpers API ──────────────────────────────────────────────────────────────
async function agendorGet(url) {
  const res = await fetch(url, { headers: AUTH_HEADERS });
  if (!res.ok) throw new Error(`HTTP ${res.status}: ${await res.text()}`);
  return res.json();
}

async function coletarOrgsComVisita(dueDateGt, dueDateLt) {
  const orgs = new Set();
  const MAX_PAGES = 20;
  for (let page = 1; page <= MAX_PAGES; page++) {
    const url = `${BASE}/tasks?per_page=100&page=${page}&dueDateGt=${dueDateGt}&dueDateLt=${dueDateLt}`;
    const json = await agendorGet(url);
    const dados = json.data || [];
    for (const t of dados) {
      if (t.type === 'Visita' && t.finishedAt && t.organization) {
        orgs.add(t.organization.id);
      }
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

// ─── Excel helpers ────────────────────────────────────────────────────────────
function aplicarEstiloHeader(row, corFundo = COR_HEADER) {
  row.eachCell(cell => {
    cell.font = { bold: true, color: { argb: BRANCO }, name: 'Arial Narrow', size: 10 };
    cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: corFundo } };
    cell.alignment = { horizontal: 'center', vertical: 'middle' };
  });
  row.height = 22;
}

// Colunas da aba de cada SDR:
// A=DATA  B=EMPRESA  C=ÁREA  D=TIPO  E=MODALIDADE  F=VALOR  G=LINK
const COL = { DATA: 1, EMPRESA: 2, AREA: 3, TIPO: 4, MODAL: 5, VALOR: 6, LINK: 7 };
const LETRA = { [COL.DATA]:'A', [COL.EMPRESA]:'B', [COL.AREA]:'C', [COL.TIPO]:'D',
                [COL.MODAL]:'E', [COL.VALOR]:'F', [COL.LINK]:'G' };

// ─── Função principal ─────────────────────────────────────────────────────────
async function main() {
  console.log('=== PREMIAÇÃO SDR — GRUPO VIGNA — 21/05 → 19/06/2026 ===\n');

  // 1. Ler planilha de reuniões realizadas
  const wb = new ExcelJS.Workbook();
  await wb.xlsx.readFile('dados/Reuniões realizada do dia 21 maio ate 19 junho.xlsx');
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

  // 2. Coletar registros de todos os SDRs dentro do período
  const reunioes = [];
  const excluidos = [];

  wsSource.eachRow((row, rowNum) => {
    if (rowNum === 1) return;
    const usuario = row.getCell(C.usuario).value;
    const sdr = SDR_MAP[usuario];
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
      empresa: row.getCell(C.empresa).value || '',
      area: extrairArea(row.getCell(C.negocio).value),
      dataAgendamento: dt,
      dataFormatada: formatDate(dt),
      modalidade: detectModalidade(row.getCell(C.descricao).value),
      tipo: null,
      valor: null,
    });
  });

  console.log(`Reuniões no período: ${reunioes.length} (${SDR_ORDER.length} SDRs)`);
  if (excluidos.length > 0) {
    console.log('Excluídos (fora do período):');
    excluidos.forEach(r => console.log(`  L${r.rowNum} | ${r.sdr} | ${r.data} | ${r.empresa}`));
  }

  // 3. Verificar histórico global nos 3 meses anteriores (lotes de 31 dias)
  console.log('\nBuscando histórico de visitas (fev–mai/2026)...');
  const lotes = [
    ['2026-02-21', '2026-03-21'],
    ['2026-03-21', '2026-04-21'],
    ['2026-04-21', '2026-05-21'],
  ];
  const orgsComHistorico = new Set();
  for (const [gt, lt] of lotes) {
    process.stdout.write(`  ${gt} → ${lt}...`);
    const orgs = await coletarOrgsComVisita(gt, lt);
    orgs.forEach(id => orgsComHistorico.add(id));
    console.log(` ${orgs.size} orgs com visita`);
    await sleep(300);
  }
  const nossoOrgIds = new Set(reunioes.filter(r => r.orgId).map(r => r.orgId));
  const orgsRelevantes = new Set([...orgsComHistorico].filter(id => nossoOrgIds.has(id)));
  console.log(`${orgsRelevantes.size} empresa(s) do dataset com visita anterior → FUP automático.`);

  // 4. Classificar NOVA vs FUP
  const sortedPorData = [...reunioes].sort((a, b) => a.dataAgendamento - b.dataAgendamento);
  const primeiraOcorrencia = new Map(); // `sdr:orgId` → timestamp
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
    r.valor = r.tipo === 'NOVA' ? VALOR_NOVA : VALOR_FUP;
  }

  // 5. Gerar Excel ──────────────────────────────────────────────────────────────
  const outWb = new ExcelJS.Workbook();
  const summaryData = {};

  // Criar Resumo primeiro (fica na primeira posição)
  const capa = outWb.addWorksheet('Resumo');

  // ── Abas por SDR ──
  for (const sdrName of SDR_ORDER) {
    const sdrReunioes = reunioes
      .filter(r => r.sdr === sdrName)
      .sort((a, b) => a.dataAgendamento - b.dataAgendamento);

    const sheet = outWb.addWorksheet(sdrName);

    // Título
    sheet.mergeCells('A1:G1');
    const titleCell = sheet.getCell('A1');
    const equipe = EQUIPE_GABRIEL.includes(sdrName) ? 'Equipe Gabriel' : 'Equipe Paola';
    titleCell.value = `PREMIAÇÃO — ${sdrName.toUpperCase()} (${equipe}) — 21/05 → 19/06/2026`;
    titleCell.font = { bold: true, size: 11, color: { argb: BRANCO }, name: 'Arial Narrow' };
    titleCell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: COR_HEADER } };
    titleCell.alignment = { horizontal: 'center', vertical: 'middle' };
    sheet.getRow(1).height = 28;
    sheet.addRow([]); // linha em branco

    // Cabeçalho de colunas
    const headerRow = sheet.addRow(['DATA', 'EMPRESA', 'ÁREA', 'TIPO', 'MODALIDADE', 'VALOR', 'LINK AGENDOR']);
    aplicarEstiloHeader(headerRow);

    // Larguras
    sheet.getColumn(COL.DATA).width   = 12;
    sheet.getColumn(COL.EMPRESA).width = 38;
    sheet.getColumn(COL.AREA).width    = 22;
    sheet.getColumn(COL.TIPO).width    = 8;
    sheet.getColumn(COL.MODAL).width   = 20;
    sheet.getColumn(COL.VALOR).width   = 14;
    sheet.getColumn(COL.LINK).width    = 50;

    // Nota sobre a fórmula
    const notaRow = sheet.addRow(['', '⚡ Altere a coluna TIPO (NOVA/FUP) → VALOR atualiza automaticamente']);
    notaRow.getCell(2).font = { italic: true, size: 9, color: { argb: 'FF555555' } };
    notaRow.height = 14;

    const PRIMEIRA_LINHA_DADOS = 5; // linha 1=título, 2=branco, 3=header, 4=nota, 5=dados

    let totalNovas = 0, totalFups = 0, totalValor = 0;
    let linhaAtual = PRIMEIRA_LINHA_DADOS;

    for (const r of sdrReunioes) {
      const link = r.orgId ? `https://beta.agendor.com.br/tasks?organizationId=${r.orgId}` : '';
      const row = sheet.addRow([r.dataFormatada, r.empresa, r.area, r.tipo, r.modalidade, null, link]);

      // ── Fórmula no VALOR (coluna F) ──
      // =IF(D{linha}="NOVA",7.5,IF(D{linha}="FUP",5,0))
      row.getCell(COL.VALOR).value = {
        formula: `=IF(${LETRA[COL.TIPO]}${linhaAtual}="NOVA",7.5,IF(${LETRA[COL.TIPO]}${linhaAtual}="FUP",5,0))`,
        result: r.valor,
      };

      // ── Dropdown NOVA/FUP no TIPO ──
      row.getCell(COL.TIPO).dataValidation = {
        type: 'list',
        allowBlank: false,
        showDropDown: false, // false = mostra a seta
        formulae: ['"NOVA,FUP"'],
      };

      // ── Formatação visual ──
      const bg = r.tipo === 'NOVA' ? COR_NOVA : COR_FUP;
      row.eachCell(cell => {
        cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: bg } };
        cell.alignment = { vertical: 'middle' };
        cell.border = { bottom: { style: 'thin', color: { argb: 'FFD0D0D0' } } };
      });
      row.getCell(COL.TIPO).alignment  = { horizontal: 'center', vertical: 'middle' };
      row.getCell(COL.MODAL).alignment = { horizontal: 'center', vertical: 'middle' };
      row.getCell(COL.VALOR).numFmt    = 'R$ #,##0.00';
      row.getCell(COL.VALOR).alignment = { horizontal: 'right', vertical: 'middle' };
      row.height = 18;

      if (r.tipo === 'NOVA') totalNovas++;
      else totalFups++;
      totalValor += r.valor;
      linhaAtual++;
    }

    // Linha de total
    sheet.addRow([]);
    linhaAtual++;
    const totalRow = sheet.addRow(['', 'TOTAL', '', `${totalNovas} N / ${totalFups} F`, '', null]);
    // Fórmula de soma da coluna VALOR
    totalRow.getCell(COL.VALOR).value = {
      formula: `=SUM(F${PRIMEIRA_LINHA_DADOS}:F${linhaAtual - 1})`,
      result: totalValor,
    };
    totalRow.getCell(COL.EMPRESA).font = { bold: true, name: 'Arial Narrow' };
    totalRow.getCell(COL.TIPO).font    = { bold: true, name: 'Arial Narrow' };
    totalRow.getCell(COL.TIPO).alignment = { horizontal: 'center' };
    totalRow.getCell(COL.VALOR).numFmt = 'R$ #,##0.00';
    totalRow.getCell(COL.VALOR).font   = { bold: true };
    totalRow.getCell(COL.VALOR).alignment = { horizontal: 'right' };

    summaryData[sdrName] = {
      novas: totalNovas,
      fups: totalFups,
      total: sdrReunioes.length,
      valor: totalValor,
    };
  }

  // ── Aba Resumo ──────────────────────────────────────────────────────────────
  capa.mergeCells('B2:I2');
  const capaTitle = capa.getCell('B2');
  capaTitle.value = 'PREMIAÇÃO SDR — GRUPO VIGNA — 21/05 → 19/06/2026';
  capaTitle.font = { bold: true, size: 14, color: { argb: BRANCO }, name: 'Arial Narrow' };
  capaTitle.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: COR_HEADER } };
  capaTitle.alignment = { horizontal: 'center', vertical: 'middle' };
  capa.getRow(2).height = 36;
  capa.addRow([]);

  const COLUNAS_RESUMO = ['', 'SDR', 'NOVAS', 'FUPs', 'TOTAL', 'TOTAL A PAGAR'];

  // Bloco Equipe Gabriel
  const hdrG = capa.addRow(['', '— EQUIPE GABRIEL —']);
  hdrG.getCell(2).font = { bold: true, color: { argb: BRANCO }, name: 'Arial Narrow', size: 10 };
  hdrG.getCell(2).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: COR_EQUIPE_G } };
  hdrG.height = 18;

  const subHdrG = capa.addRow(COLUNAS_RESUMO);
  aplicarEstiloHeader(subHdrG, COR_EQUIPE_G);

  let totG = { novas: 0, fups: 0, total: 0, valor: 0 };
  for (const sdr of EQUIPE_GABRIEL) {
    const d = summaryData[sdr] || { novas: 0, fups: 0, total: 0, valor: 0 };
    const row = capa.addRow(['', sdr, d.novas, d.fups, d.total, d.valor]);
    row.getCell(2).font = { name: 'Arial Narrow' };
    row.getCell(6).numFmt = 'R$ #,##0.00';
    row.getCell(6).font = { bold: true };
    row.getCell(6).alignment = { horizontal: 'right' };
    row.height = 18;
    totG.novas += d.novas; totG.fups += d.fups; totG.total += d.total; totG.valor += d.valor;
  }
  const subtotalG = capa.addRow(['', 'SUBTOTAL GABRIEL', totG.novas, totG.fups, totG.total, totG.valor]);
  subtotalG.getCell(2).font = { bold: true, name: 'Arial Narrow' };
  [3,4,5,6].forEach(i => subtotalG.getCell(i).font = { bold: true });
  subtotalG.getCell(6).numFmt = 'R$ #,##0.00';
  subtotalG.getCell(6).alignment = { horizontal: 'right' };
  subtotalG.height = 18;

  capa.addRow([]);

  // Bloco Equipe Paola
  const hdrP = capa.addRow(['', '— EQUIPE PAOLA —']);
  hdrP.getCell(2).font = { bold: true, color: { argb: BRANCO }, name: 'Arial Narrow', size: 10 };
  hdrP.getCell(2).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: COR_EQUIPE_P } };
  hdrP.height = 18;

  const subHdrP = capa.addRow(COLUNAS_RESUMO);
  aplicarEstiloHeader(subHdrP, COR_EQUIPE_P);

  let totP = { novas: 0, fups: 0, total: 0, valor: 0 };
  for (const sdr of EQUIPE_PAOLA) {
    const d = summaryData[sdr] || { novas: 0, fups: 0, total: 0, valor: 0 };
    const row = capa.addRow(['', sdr, d.novas, d.fups, d.total, d.valor]);
    row.getCell(2).font = { name: 'Arial Narrow' };
    row.getCell(6).numFmt = 'R$ #,##0.00';
    row.getCell(6).font = { bold: true };
    row.getCell(6).alignment = { horizontal: 'right' };
    row.height = 18;
    totP.novas += d.novas; totP.fups += d.fups; totP.total += d.total; totP.valor += d.valor;
  }
  const subtotalP = capa.addRow(['', 'SUBTOTAL PAOLA', totP.novas, totP.fups, totP.total, totP.valor]);
  subtotalP.getCell(2).font = { bold: true, name: 'Arial Narrow' };
  [3,4,5,6].forEach(i => subtotalP.getCell(i).font = { bold: true });
  subtotalP.getCell(6).numFmt = 'R$ #,##0.00';
  subtotalP.getCell(6).alignment = { horizontal: 'right' };
  subtotalP.height = 18;

  capa.addRow([]);

  // Total geral
  const gtValor = totG.valor + totP.valor;
  const gtTotal = totG.total + totP.total;
  const gtNovas = totG.novas + totP.novas;
  const gtFups  = totG.fups  + totP.fups;
  const totalCapaRow = capa.addRow(['', 'TOTAL GERAL', gtNovas, gtFups, gtTotal, gtValor]);
  totalCapaRow.getCell(2).font = { bold: true, size: 12, name: 'Arial Narrow' };
  [3,4,5,6].forEach(i => totalCapaRow.getCell(i).font = { bold: true, size: 12 });
  totalCapaRow.getCell(6).numFmt = 'R$ #,##0.00';
  totalCapaRow.getCell(6).alignment = { horizontal: 'right' };
  totalCapaRow.height = 22;

  // Legenda de valores
  capa.addRow([]);
  capa.addRow([]);
  const legHdr = capa.addRow(['', 'TABELA DE VALORES (válida para todos)']);
  legHdr.getCell(2).font = { bold: true, name: 'Arial Narrow' };
  [
    ['', 'NOVA', '', '', '', 'R$ 7,50'],
    ['', 'FUP',  '', '', '', 'R$ 5,00'],
  ].forEach(lr => {
    const r = capa.addRow(lr);
    r.getCell(6).alignment = { horizontal: 'right' };
  });

  // Larguras da capa
  capa.getColumn(1).width = 3;
  capa.getColumn(2).width = 30;
  capa.getColumn(3).width = 10;
  capa.getColumn(4).width = 10;
  capa.getColumn(5).width = 12;
  capa.getColumn(6).width = 18;

  // ── Salvar ──────────────────────────────────────────────────────────────────
  const outDir = path.join('relatorios', 'premiacao');
  fs.mkdirSync(outDir, { recursive: true });
  const outPath = path.join(outDir, 'Premiacao_SDR_MAIO_JUNHO_2026.xlsx');
  await outWb.xlsx.writeFile(outPath);

  console.log(`\n✓ Arquivo gerado: ${outPath}`);
  console.log('\n=== RESULTADO ===');
  console.log('EQUIPE GABRIEL:');
  for (const sdr of EQUIPE_GABRIEL) {
    const d = summaryData[sdr] || { novas:0, fups:0, total:0, valor:0 };
    console.log(`  ${sdr}: ${d.novas} NOVAS + ${d.fups} FUPs = ${d.total} reuniões → R$ ${d.valor.toFixed(2).replace('.',',')}`);
  }
  console.log(`  SUBTOTAL: R$ ${totG.valor.toFixed(2).replace('.',',')}`);
  console.log('\nEQUIPE PAOLA:');
  for (const sdr of EQUIPE_PAOLA) {
    const d = summaryData[sdr] || { novas:0, fups:0, total:0, valor:0 };
    console.log(`  ${sdr}: ${d.novas} NOVAS + ${d.fups} FUPs = ${d.total} reuniões → R$ ${d.valor.toFixed(2).replace('.',',')}`);
  }
  console.log(`  SUBTOTAL: R$ ${totP.valor.toFixed(2).replace('.',',')}`);
  console.log(`\nTOTAL GERAL: R$ ${gtValor.toFixed(2).replace('.',',')}`);
  if (excluidos.length > 0) {
    console.log(`\nObs.: ${excluidos.length} registro(s) excluído(s) por fora do período.`);
    excluidos.forEach(r => console.log(`  L${r.rowNum} | ${r.sdr} | ${r.data} | ${r.empresa}`));
  }
}

main().catch(e => { console.error('ERRO:', e.message); process.exit(1); });

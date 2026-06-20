// premiar-sdr.js
// Gera a planilha de premiação dos SDRs com base no export do Agendor
// Uso: node --env-file=.env automacoes/premiar-sdr.js
//
// Lógica NOVA vs FUP:
//   - Busca GLOBALMENTE todas as Visitas finalizadas nos 3 meses antes do período (lotes de 31 dias)
//   - Se a empresa já teve visita antes de 21/05 → FUP
//   - Se não → primeira ocorrência no período = NOVA, demais = FUP

const ExcelJS = require('exceljs');
const path = require('path');
const fs = require('fs');

const TOKEN = process.env.AGENDOR_API_TOKEN;
if (!TOKEN) { console.error('AGENDOR_API_TOKEN não encontrado. Rode com: node --env-file=.env automacoes/premiar-sdr.js'); process.exit(1); }

const BASE = 'https://api.agendor.com.br/v3';
const AUTH_HEADERS = { Authorization: `Token ${TOKEN}` };

const SDR_MAP = {
  'Michelle Valente Araujo - Matriz SP': 'Michelle',
  'Kailany Santos da Silva - Matriz SP': 'Kailany',
  'Ana Karolayne - Matriz SP': 'Ana Karolayne',
  'Igor Souza Ferreira - Matriz SP': 'Igor'
};
const SDR_ORDER = ['Michelle', 'Kailany', 'Ana Karolayne', 'Igor'];

const INICIO = new Date('2026-05-21T00:00:00.000Z');
const FIM    = new Date('2026-06-20T02:59:59.000Z'); // 19/06 23:59 BRT

const VALORES = {
  NOVA: { 'VIDEO CONFERÊNCIA': 7.50, PRESENCIAL: 15.00 },
  FUP:  { 'VIDEO CONFERÊNCIA': 5.00, PRESENCIAL: 7.50  }
};

const COR_HEADER = 'FF1B2A4A';
const COR_NOVA   = 'FFD6E4F7';
const COR_FUP    = 'FFFFF3CD';
const BRANCO     = 'FFFFFFFF';

async function agendorGet(url) {
  const res = await fetch(url, { headers: AUTH_HEADERS });
  if (!res.ok) throw new Error(`HTTP ${res.status} em ${url}: ${await res.text()}`);
  return res.json();
}

// Busca tarefas num intervalo (máx 20 páginas), retornando Set de orgIds com Visita finalizada
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

function estiloCelulasHeader(row) {
  row.eachCell(cell => {
    cell.font = { bold: true, color: { argb: BRANCO }, name: 'Arial Narrow', size: 10 };
    cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: COR_HEADER } };
    cell.alignment = { horizontal: 'center', vertical: 'middle' };
  });
  row.height = 22;
}

async function main() {
  console.log('=== PREMIAÇÃO SDR — GRUPO VIGNA — 21/05 → 19/06/2026 ===\n');

  // 1. Ler planilha de reuniões realizadas
  const wb = new ExcelJS.Workbook();
  await wb.xlsx.readFile('dados/Reuniões realizada do dia 21 maio ate 19 junho.xlsx');
  const ws = wb.worksheets[0];

  const hdrMap = {};
  ws.getRow(1).eachCell((cell, col) => { hdrMap[cell.value] = col; });

  const C = {
    usuario:       hdrMap['Usuário que realizou a tarefa'],
    empresa:       hdrMap['Empresa relacionada'],
    codigoEmpresa: hdrMap['Código da Empresa'],
    negocio:       hdrMap['Negócio relacionado'],
    dataAgend:     hdrMap['Data de agendamento'],
    descricao:     hdrMap['Descrição'],
  };

  // 2. Coletar registros do time dentro do período
  const reunioes = [];
  const excluidos = [];
  ws.eachRow((row, rowNum) => {
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

  console.log(`Reuniões no período: ${reunioes.length}`);
  if (excluidos.length > 0) {
    console.log('Excluídos (fora do período):');
    excluidos.forEach(r => console.log(`  L${r.rowNum} | ${r.sdr} | ${r.data} | ${r.empresa}`));
  }

  // 3. Buscar histórico global de visitas nos 3 meses anteriores ao período
  // Intervalo máximo da API: 31 dias por janela
  console.log('\nBuscando histórico de visitas (3 meses anteriores ao período)...');
  const lotes = [
    ['2026-02-21', '2026-03-21'],
    ['2026-03-21', '2026-04-21'],
    ['2026-04-21', '2026-05-21'],
  ];

  const orgsComHistorico = new Set();
  for (const [gt, lt] of lotes) {
    console.log(`  Verificando ${gt} → ${lt}...`);
    const orgs = await coletarOrgsComVisita(gt, lt);
    orgs.forEach(id => orgsComHistorico.add(id));
    await sleep(300);
  }

  // Filtrar apenas as que aparecem no nosso dataset
  const nossoOrgIds = new Set(reunioes.filter(r => r.orgId).map(r => r.orgId));
  const orgsRelevantesComHistorico = new Set([...orgsComHistorico].filter(id => nossoOrgIds.has(id)));
  console.log(`${orgsRelevantesComHistorico.size} empresa(s) do dataset com visita nos 3 meses anteriores → FUP automático.`);

  // 4. Classificar NOVA vs FUP
  const sorted = [...reunioes].sort((a, b) => a.dataAgendamento - b.dataAgendamento);
  const primeiraOcorrencia = new Map(); // `sdr:orgId` → timestamp
  for (const r of sorted) {
    if (!r.orgId) continue;
    const key = `${r.sdr}:${r.orgId}`;
    if (!primeiraOcorrencia.has(key)) primeiraOcorrencia.set(key, r.dataAgendamento.getTime());
  }

  for (const r of reunioes) {
    if (!r.orgId || orgsRelevantesComHistorico.has(r.orgId)) {
      r.tipo = 'FUP';
    } else {
      const key = `${r.sdr}:${r.orgId}`;
      const primeira = primeiraOcorrencia.get(key);
      r.tipo = (r.dataAgendamento.getTime() === primeira) ? 'NOVA' : 'FUP';
    }
    r.valor = VALORES[r.tipo][r.modalidade] ?? 0;
  }

  // 5. Gerar Excel
  const outWb = new ExcelJS.Workbook();

  const summaryData = {};

  // Criar Resumo primeiro (pra ficar na primeira posição)
  const capa = outWb.addWorksheet('Resumo');

  // Abas por SDR
  for (const sdrName of SDR_ORDER) {
    const sdrReunioes = reunioes
      .filter(r => r.sdr === sdrName)
      .sort((a, b) => a.dataAgendamento - b.dataAgendamento);

    const sheet = outWb.addWorksheet(sdrName);

    // Título
    sheet.mergeCells('A1:G1');
    const titleCell = sheet.getCell('A1');
    titleCell.value = `PREMIAÇÃO — ${sdrName.toUpperCase()} — 21/05 → 19/06/2026`;
    titleCell.font = { bold: true, size: 12, color: { argb: BRANCO }, name: 'Arial Narrow' };
    titleCell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: COR_HEADER } };
    titleCell.alignment = { horizontal: 'center', vertical: 'middle' };
    sheet.getRow(1).height = 28;

    sheet.addRow([]); // espaço

    // Cabeçalho de colunas
    const headerRow = sheet.addRow(['DATA', 'EMPRESA', 'ÁREA', 'TIPO', 'MODALIDADE', 'VALOR', 'LINK AGENDOR']);
    estiloCelulasHeader(headerRow);

    sheet.getColumn(1).width = 12;
    sheet.getColumn(2).width = 38;
    sheet.getColumn(3).width = 22;
    sheet.getColumn(4).width = 8;
    sheet.getColumn(5).width = 20;
    sheet.getColumn(6).width = 14;
    sheet.getColumn(7).width = 50;

    let totalNovas = 0, totalFups = 0, totalValor = 0;

    for (const r of sdrReunioes) {
      const link = r.orgId ? `https://beta.agendor.com.br/tasks?organizationId=${r.orgId}` : '';
      const row = sheet.addRow([r.dataFormatada, r.empresa, r.area, r.tipo, r.modalidade, r.valor, link]);

      const bg = r.tipo === 'NOVA' ? COR_NOVA : COR_FUP;
      row.eachCell(cell => {
        cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: bg } };
        cell.alignment = { vertical: 'middle' };
        cell.border = { bottom: { style: 'thin', color: { argb: 'FFD0D0D0' } } };
      });
      row.getCell(4).alignment = { horizontal: 'center', vertical: 'middle' };
      row.getCell(5).alignment = { horizontal: 'center', vertical: 'middle' };
      row.getCell(6).numFmt = 'R$ #,##0.00';
      row.getCell(6).alignment = { horizontal: 'right', vertical: 'middle' };
      row.height = 18;

      if (r.tipo === 'NOVA') totalNovas++;
      else totalFups++;
      totalValor += r.valor;
    }

    // Total
    sheet.addRow([]);
    const totalRow = sheet.addRow(['', 'TOTAL', '', `${totalNovas} N / ${totalFups} F`, '', totalValor]);
    totalRow.getCell(2).font = { bold: true, name: 'Arial Narrow' };
    totalRow.getCell(4).font = { bold: true, name: 'Arial Narrow' };
    totalRow.getCell(4).alignment = { horizontal: 'center' };
    totalRow.getCell(6).numFmt = 'R$ #,##0.00';
    totalRow.getCell(6).font = { bold: true };
    totalRow.getCell(6).alignment = { horizontal: 'right' };

    summaryData[sdrName] = {
      novas: totalNovas,
      fups: totalFups,
      total: sdrReunioes.length,
      valor: totalValor,
      valorNovas: sdrReunioes.filter(r => r.tipo === 'NOVA').reduce((s, r) => s + r.valor, 0),
      valorFups:  sdrReunioes.filter(r => r.tipo === 'FUP' ).reduce((s, r) => s + r.valor, 0),
    };
  }

  // Preencher aba Resumo (criada antes das SDR)
  capa.mergeCells('B2:I2');
  const capaTitle = capa.getCell('B2');
  capaTitle.value = 'PREMIAÇÃO SDR — GRUPO VIGNA — 21/05 → 19/06/2026';
  capaTitle.font = { bold: true, size: 14, color: { argb: BRANCO }, name: 'Arial Narrow' };
  capaTitle.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: COR_HEADER } };
  capaTitle.alignment = { horizontal: 'center', vertical: 'middle' };
  capa.getRow(2).height = 36;

  capa.addRow([]);

  const capaHdrRow = capa.addRow(['', 'SDR', 'NOVAS', 'FUPs', 'TOTAL REUNIÕES', 'VALOR NOVAS', 'VALOR FUPs', 'TOTAL A PAGAR']);
  estiloCelulasHeader(capaHdrRow);

  let gtNovas = 0, gtFups = 0, gtTotal = 0, gtValor = 0;

  for (const sdrName of SDR_ORDER) {
    const d = summaryData[sdrName];
    const row = capa.addRow(['', sdrName, d.novas, d.fups, d.total, d.valorNovas, d.valorFups, d.valor]);
    row.getCell(2).font = { bold: true, name: 'Arial Narrow' };
    [6, 7, 8].forEach(i => {
      row.getCell(i).numFmt = 'R$ #,##0.00';
      row.getCell(i).alignment = { horizontal: 'right' };
    });
    row.getCell(8).font = { bold: true };
    row.height = 20;
    gtNovas += d.novas; gtFups += d.fups; gtTotal += d.total; gtValor += d.valor;
  }

  capa.addRow([]);
  const totalCapaRow = capa.addRow(['', 'TOTAL GERAL', gtNovas, gtFups, gtTotal, '', '', gtValor]);
  totalCapaRow.getCell(2).font = { bold: true, size: 11, name: 'Arial Narrow' };
  [3, 4, 5].forEach(i => totalCapaRow.getCell(i).font = { bold: true });
  totalCapaRow.getCell(8).numFmt = 'R$ #,##0.00';
  totalCapaRow.getCell(8).font = { bold: true, size: 11 };
  totalCapaRow.height = 22;

  // Tabela de valores / legenda
  capa.addRow([]);
  capa.addRow([]);
  const legHdr = capa.addRow(['', 'TABELA DE VALORES']);
  legHdr.getCell(2).font = { bold: true, name: 'Arial Narrow' };
  [
    ['', 'NOVA — Videoconferência', '', '', '', '', '', 'R$ 7,50'],
    ['', 'FUP — Videoconferência',  '', '', '', '', '', 'R$ 5,00'],
    ['', 'NOVA — Presencial',       '', '', '', '', '', 'R$ 15,00'],
    ['', 'FUP — Presencial',        '', '', '', '', '', 'R$ 7,50'],
  ].forEach(lr => capa.addRow(lr));

  capa.getColumn(1).width = 3;
  capa.getColumn(2).width = 35;
  capa.getColumn(3).width = 10;
  capa.getColumn(4).width = 10;
  capa.getColumn(5).width = 18;
  capa.getColumn(6).width = 16;
  capa.getColumn(7).width = 16;
  capa.getColumn(8).width = 18;

  // Salvar
  const outDir = path.join('relatorios', 'premiacao');
  fs.mkdirSync(outDir, { recursive: true });
  const outPath = path.join(outDir, 'Premiacao_SDR_MAIO_JUNHO_2026.xlsx');
  await outWb.xlsx.writeFile(outPath);

  console.log(`\n✓ Arquivo gerado: ${outPath}`);
  console.log('\n=== RESULTADO ===');
  for (const sdr of SDR_ORDER) {
    const d = summaryData[sdr];
    console.log(`${sdr}: ${d.novas} NOVAS + ${d.fups} FUPs = ${d.total} reuniões → R$ ${d.valor.toFixed(2).replace('.', ',')}`);
  }
  console.log(`\nTOTAL GERAL: R$ ${gtValor.toFixed(2).replace('.', ',')}`);
  if (excluidos.length > 0) {
    console.log(`\nObs.: ${excluidos.length} registro(s) excluído(s) por fora do período:`);
    excluidos.forEach(r => console.log(`  L${r.rowNum} | ${r.sdr} | ${r.data} | ${r.empresa}`));
  }
}

main().catch(e => { console.error('ERRO:', e.message); process.exit(1); });

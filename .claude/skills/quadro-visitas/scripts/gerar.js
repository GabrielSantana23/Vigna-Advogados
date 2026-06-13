// Gera o Quadro de Visitas Grupo Vigna em .xlsx a partir de um JSON de dados.
// Uso: node gerar.js dados.json saida.xlsx

const ExcelJS = require("exceljs");
const fs = require("fs");
const path = require("path");

const [, , inputPath, outputPath] = process.argv;
if (!inputPath || !outputPath) {
  console.error("Uso: node gerar.js dados.json saida.xlsx");
  process.exit(1);
}

const dados = JSON.parse(fs.readFileSync(inputPath, "utf-8"));
const { data, linhas } = dados;

const AZUL_ESCURO = "FF1B2A4A";
const BRANCO = "FFFFFFFF";
const BORDA = "FFD8DFF0";
const UNIDADE_CORES = {
  ALPHAVILLE: "FF4A235A",
  CAMPINAS: "FF1A3A6A",
  MG: "FF1A5C3A",
  PR: "FF7A3B00",
  RS: "FF8B1A1A",
  MATRIZ: "FF1B2A4A",
};

const COL_WIDTHS = [14, 8, 34, 20, 16, 14, 28, 14, 64];
const N_COLS = COL_WIDTHS.length;

const thinBorder = {
  top: { style: "thin", color: { argb: BORDA } },
  bottom: { style: "thin", color: { argb: BORDA } },
  left: { style: "thin", color: { argb: BORDA } },
  right: { style: "thin", color: { argb: BORDA } },
};

const wb = new ExcelJS.Workbook();
const ws = wb.addWorksheet("Quadro de Visitas", {
  views: [{ showGridLines: false, zoomScale: 95 }],
});

ws.columns = COL_WIDTHS.map((w) => ({ width: w }));

// Linha 1 — vazia (espaçamento)
ws.getRow(1).height = 6;

// Linha 2 — título
ws.mergeCells(2, 1, 2, N_COLS);
const titulo = ws.getCell(2, 1);
titulo.value = `GRUPO VIGNA  ·  QUADRO DE VISITAS  ·  ${data}`;
titulo.font = { name: "Georgia", size: 12, bold: true, color: { argb: BRANCO } };
titulo.fill = { type: "pattern", pattern: "solid", fgColor: { argb: AZUL_ESCURO } };
titulo.alignment = { horizontal: "center", vertical: "middle" };
ws.getRow(2).height = 28;

// Linha 3 — subtítulo
ws.mergeCells(3, 1, 3, N_COLS);
const sub = ws.getCell(3, 1);
sub.value = "Reuniões agendadas para o dia, por unidade e horário";
sub.font = { name: "Arial", size: 9, color: { argb: "FFA9B6D6" } };
sub.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FF131F38" } };
sub.alignment = { horizontal: "center", vertical: "middle" };
ws.getRow(3).height = 20;

// Linha 4 — vazia (espaçamento)
ws.getRow(4).height = 6;

// Linha 5 — cabeçalho
const labels = [
  "UNIDADE",
  "HORÁRIO",
  "RAZÃO SOCIAL",
  "CNPJ",
  "CONTATO",
  "CARGO",
  "CONSULTOR VIGNA",
  "REGIME",
  "OPORTUNIDADES PRIORITÁRIAS",
];
labels.forEach((label, i) => {
  const c = ws.getCell(5, i + 1);
  c.value = label;
  c.font = { name: "Arial", size: 8.5, bold: true, color: { argb: BRANCO } };
  c.fill = { type: "pattern", pattern: "solid", fgColor: { argb: AZUL_ESCURO } };
  c.alignment = { horizontal: "center", vertical: "middle", wrapText: true };
  c.border = thinBorder;
});
ws.getRow(5).height = 24;

// Linhas de dados
let linhaAtual = 6;
linhas.forEach((item, idx) => {
  const fundo = idx % 2 === 0 ? "FFEAF0FB" : "FFF5F6FA";
  const valores = [
    item.unidade,
    item.horario,
    item.razao,
    item.cnpj,
    item.pessoa,
    item.cargo,
    item.consultor,
    item.regime,
    Array.isArray(item.oportunidades) ? item.oportunidades.join("\n") : item.oportunidades,
  ];

  valores.forEach((valor, col) => {
    const c = ws.getCell(linhaAtual, col + 1);
    c.value = valor;
    c.border = thinBorder;
    c.alignment = { horizontal: "center", vertical: "middle", wrapText: true };

    if (col === 0) {
      // Unidade
      c.fill = {
        type: "pattern",
        pattern: "solid",
        fgColor: { argb: UNIDADE_CORES[item.unidade] || AZUL_ESCURO },
      };
      c.font = { name: "Arial", size: 8.5, bold: true, color: { argb: BRANCO } };
    } else if (col === 1) {
      // Horário
      c.font = { name: "Arial", size: 8.5, bold: true };
      c.fill = { type: "pattern", pattern: "solid", fgColor: { argb: fundo } };
    } else if (col === 2) {
      // Razão Social
      c.font = { name: "Arial", size: 8.5, bold: true };
      c.alignment = { horizontal: "left", vertical: "middle", wrapText: true };
      c.fill = { type: "pattern", pattern: "solid", fgColor: { argb: fundo } };
    } else if (col === 8) {
      // Oportunidades
      c.font = { name: "Arial", size: 8 };
      c.alignment = { horizontal: "left", vertical: "middle", wrapText: true };
      c.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFFAFBFF" } };
    } else {
      c.font = { name: "Arial", size: 8.5 };
      c.fill = { type: "pattern", pattern: "solid", fgColor: { argb: fundo } };
    }
  });

  ws.getRow(linhaAtual).height = 38;
  linhaAtual += 1;
});

// Rodapé
ws.mergeCells(linhaAtual + 1, 1, linhaAtual + 1, N_COLS);
const rodape = ws.getCell(linhaAtual + 1, 1);
rodape.value = `Documento confidencial · Uso interno · Grupo Vigna © ${new Date().getFullYear()}`;
rodape.font = { name: "Arial", size: 7.5, italic: true, color: { argb: "FF8899BB" } };
rodape.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFF0F3FA" } };
rodape.alignment = { horizontal: "center", vertical: "middle" };

// Congela cabeçalho
ws.views = [{ showGridLines: false, zoomScale: 95, state: "frozen", ySplit: 5 }];

fs.mkdirSync(path.dirname(outputPath), { recursive: true });
wb.xlsx.writeFile(outputPath).then(() => {
  console.log(`Arquivo gerado em: ${outputPath}`);
});

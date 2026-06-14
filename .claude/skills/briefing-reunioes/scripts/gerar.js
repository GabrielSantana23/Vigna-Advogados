// Gera o Briefing Pré-Reunião Grupo Vigna em .xlsx, uma aba por empresa/reunião.
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
const { data, reunioes } = dados;

const AZUL_ESCURO = "FF1B2A4A";
const AZUL_SUBTITULO = "FF131F38";
const BRANCO = "FFFFFFFF";
const CINZA_AZULADO = "FFA9B6D6";
const BORDA = "FFD8DFF0";

const thinBorder = {
  top: { style: "thin", color: { argb: BORDA } },
  bottom: { style: "thin", color: { argb: BORDA } },
  left: { style: "thin", color: { argb: BORDA } },
  right: { style: "thin", color: { argb: BORDA } },
};

const CAMPOS = [
  { label: "Consultores Vigna", key: "consultores" },
  { label: "Horário Reunião", key: "horario" },
  { label: "Nome da Empresa", key: "empresa" },
  { label: "CNPJ", key: "cnpj" },
  { label: "Responsável da Empresa", key: "responsavel" },
  { label: "Cargo", key: "cargo" },
  { label: "Resumo da Empresa", key: "resumo" },
  { label: "Oportunidades", key: "oportunidades" },
  { label: "Regime Tributário", key: "regime" },
  { label: "Site da Empresa", key: "site" },
  { label: "Pauta", key: "pauta" },
  { label: "Link da Reunião", key: "link" },
];

const LINHAS_GRANDES = new Set(["Resumo da Empresa", "Oportunidades"]);

function nomeAba(item, idx) {
  const base = (item.empresa || `Reuniao ${idx + 1}`).replace(/[\\/?*[\]:]/g, "").trim();
  const prefixo = String(idx + 1).padStart(2, "0");
  return `${prefixo} - ${base}`.slice(0, 31);
}

const wb = new ExcelJS.Workbook();

reunioes.forEach((item, idx) => {
  const ws = wb.addWorksheet(nomeAba(item, idx), {
    views: [{ showGridLines: false, zoomScale: 100 }],
  });

  ws.columns = [{ width: 26 }, { width: 100 }];

  // Linha 1 — espaçamento
  ws.getRow(1).height = 6;

  // Linha 2 — título
  ws.mergeCells(2, 1, 2, 2);
  const titulo = ws.getCell(2, 1);
  titulo.value = `GRUPO VIGNA  ·  BRIEFING PRÉ-REUNIÃO  ·  ${data}`;
  titulo.font = { name: "Georgia", size: 12, bold: true, color: { argb: BRANCO } };
  titulo.fill = { type: "pattern", pattern: "solid", fgColor: { argb: AZUL_ESCURO } };
  titulo.alignment = { horizontal: "center", vertical: "middle" };
  ws.getRow(2).height = 28;

  // Linha 3 — subtítulo (nome da empresa)
  ws.mergeCells(3, 1, 3, 2);
  const sub = ws.getCell(3, 1);
  sub.value = item.empresa || "Empresa não identificada";
  sub.font = { name: "Arial", size: 9, color: { argb: CINZA_AZULADO } };
  sub.fill = { type: "pattern", pattern: "solid", fgColor: { argb: AZUL_SUBTITULO } };
  sub.alignment = { horizontal: "center", vertical: "middle" };
  ws.getRow(3).height = 20;

  // Linha 4 — espaçamento
  ws.getRow(4).height = 6;

  // Linha 5 — cabeçalho
  ["CAMPO", "INFORMAÇÃO"].forEach((label, i) => {
    const c = ws.getCell(5, i + 1);
    c.value = label;
    c.font = { name: "Arial", size: 9, bold: true, color: { argb: BRANCO } };
    c.fill = { type: "pattern", pattern: "solid", fgColor: { argb: AZUL_ESCURO } };
    c.alignment = { horizontal: i === 0 ? "center" : "left", vertical: "middle" };
    c.border = thinBorder;
  });
  ws.getRow(5).height = 22;

  // Linhas de dados
  let linhaAtual = 6;
  CAMPOS.forEach((campo, i) => {
    const fundo = i % 2 === 0 ? "FFEAF0FB" : "FFF5F6FA";
    const valor = item[campo.key] || "Não localizado";

    const cCampo = ws.getCell(linhaAtual, 1);
    cCampo.value = campo.label;
    cCampo.font = { name: "Arial", size: 9, bold: true };
    cCampo.alignment = { horizontal: "left", vertical: "middle", wrapText: true };
    cCampo.fill = { type: "pattern", pattern: "solid", fgColor: { argb: fundo } };
    cCampo.border = thinBorder;

    const cValor = ws.getCell(linhaAtual, 2);
    cValor.value = valor;
    cValor.font = { name: "Arial", size: 9 };
    cValor.alignment = { horizontal: "left", vertical: "middle", wrapText: true };
    cValor.fill = { type: "pattern", pattern: "solid", fgColor: { argb: fundo } };
    cValor.border = thinBorder;

    ws.getRow(linhaAtual).height = LINHAS_GRANDES.has(campo.label) ? 90 : 22;
    linhaAtual += 1;
  });

  // Rodapé
  ws.mergeCells(linhaAtual + 1, 1, linhaAtual + 1, 2);
  const rodape = ws.getCell(linhaAtual + 1, 1);
  rodape.value = `Documento confidencial · Uso interno · Grupo Vigna © ${new Date().getFullYear()}`;
  rodape.font = { name: "Arial", size: 7.5, italic: true, color: { argb: "FF8899BB" } };
  rodape.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFF0F3FA" } };
  rodape.alignment = { horizontal: "center", vertical: "middle" };

  ws.views = [{ showGridLines: false, zoomScale: 100, state: "frozen", ySplit: 5 }];
});

fs.mkdirSync(path.dirname(outputPath), { recursive: true });
wb.xlsx.writeFile(outputPath).then(() => {
  console.log(`Arquivo gerado em: ${outputPath}`);
  console.log(`${reunioes.length} aba(s) criada(s).`);
});

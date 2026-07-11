// Grava as linhas já montadas (empresa + dados enriquecidos do Agendor) na aba
// Lançamento_Reuniões do Controle_Comercial_Vigna. Faz backup do arquivo antes de gravar.
//
// Uso: node preencher-planilha.js "caminho/Controle_Comercial_Vigna - mes.xlsx" linhas.json
//
// linhas.json — objeto indexado pelo número da linha na planilha (string), ex:
// {
//   "2": {
//     "empresa": "CONSTRUTORA CLARK LTDA",
//     "data": "14/07/2026",
//     "segmento": "Construtora",
//     "cnpj": "65527939000154",
//     "responsavel": "Lourdes",
//     "cargo": "RH",
//     "cidade": "São Paulo",
//     "estado": "SP",
//     "sdr": "Ana Karolayne - Matriz SP",
//     "consultor": "Pamela"
//   },
//   "3": { "empresa": "Empresa não localizada no Agendor" },   // só nome, resto fica em branco
//   "4": { "limpar": true }                                     // apaga a linha (ex: duplicata)
// }

const fs = require("fs");
const path = require("path");
const ExcelJS = require("exceljs");

const [, , planilhaPath, linhasPath] = process.argv;
if (!planilhaPath || !linhasPath) {
  console.error('Uso: node preencher-planilha.js "planilha.xlsx" linhas.json');
  process.exit(1);
}

function cnpjFmt(v) {
  if (!v) return "Não localizado";
  const digitos = String(v).replace(/\D/g, "");
  if (digitos.length !== 14) return v;
  return digitos.replace(/(\d{2})(\d{3})(\d{3})(\d{4})(\d{2})/, "$1.$2.$3/$4-$5");
}

function parseDataBr(str) {
  const [dia, mes, ano] = str.split("/").map(Number);
  return new Date(Date.UTC(ano, mes - 1, dia, 12, 0, 0));
}

async function main() {
  const linhas = JSON.parse(fs.readFileSync(linhasPath, "utf-8"));

  const backupPath = planilhaPath.replace(/\.xlsx$/i, `.BACKUP-${Date.now()}.xlsx`);
  fs.copyFileSync(planilhaPath, backupPath);

  const wb = new ExcelJS.Workbook();
  await wb.xlsx.readFile(planilhaPath);
  const ws = wb.getWorksheet("Lançamento_Reuniões");
  if (!ws) throw new Error("Aba 'Lançamento_Reuniões' não encontrada na planilha.");

  for (const [rowNum, dados] of Object.entries(linhas)) {
    const r = parseInt(rowNum, 10);
    const row = ws.getRow(r);

    if (dados.limpar) {
      for (let c = 1; c <= 10; c++) row.getCell(c).value = null;
      row.commit();
      continue;
    }

    row.getCell(2).value = dados.empresa;

    // Empresa não localizada no Agendor (nem organização, nem tarefa): só grava o nome.
    if (!dados.data && !dados.cnpj && !dados.segmento && !dados.cidade && !dados.estado) {
      row.commit();
      continue;
    }

    // Organização localizada no Agendor mas sem tarefa de Reunião no período: grava os
    // dados cadastrais da empresa e deixa os campos da reunião (data, responsável, cargo,
    // SDR, consultor) em branco em vez de "Não informado" — não houve reunião pra descrever.
    row.getCell(3).value = dados.segmento || "Não localizado";
    row.getCell(4).value = cnpjFmt(dados.cnpj);
    row.getCell(7).value = dados.cidade || "Não localizado";
    row.getCell(8).value = dados.estado || "Não localizado";

    if (dados.data) {
      row.getCell(1).value = parseDataBr(dados.data);
      row.getCell(1).numFmt = "dd/mm/yyyy";
      row.getCell(5).value = dados.responsavel || "Não informado";
      row.getCell(6).value = dados.cargo || "Não informado";
      row.getCell(9).value = dados.sdr || "Não informado";
      row.getCell(10).value = dados.consultor || "Não informado";
    }
    row.commit();
  }

  await wb.xlsx.writeFile(planilhaPath);
  console.log(`Planilha atualizada: ${planilhaPath}`);
  console.log(`Backup salvo em: ${backupPath}`);
}

main().catch((e) => {
  console.error(e);
  process.exit(1);
});

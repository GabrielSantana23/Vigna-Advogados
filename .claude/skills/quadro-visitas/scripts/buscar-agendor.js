// Busca as reuniões/visitas do dia no Agendor e monta um JSON intermediário
// pra enriquecimento (CNPJ, regime, unidade, oportunidades) e geração do Quadro de Visitas.
//
// Uso: node --env-file=../../../../.env buscar-agendor.js DD/MM/AAAA saida.json

const fs = require("fs");
const path = require("path");

const [, , dataArg, outputPath] = process.argv;
if (!dataArg || !outputPath) {
  console.error("Uso: node buscar-agendor.js DD/MM/AAAA saida.json");
  process.exit(1);
}

const TOKEN = process.env.AGENDOR_API_TOKEN;
if (!TOKEN) {
  console.error("AGENDOR_API_TOKEN não encontrado. Rode com: node --env-file=<caminho para .env> buscar-agendor.js ...");
  process.exit(1);
}

const [dia, mes, ano] = dataArg.split("/").map(Number);
const dataAlvo = `${ano}-${String(mes).padStart(2, "0")}-${String(dia).padStart(2, "0")}`;

const hoje = new Date(Date.UTC(ano, mes - 1, dia));
const amanha = new Date(Date.UTC(ano, mes - 1, dia + 1));
const fmt = (d) => d.toISOString().slice(0, 10);

const BASE = "https://api.agendor.com.br/v3";
const headers = { Authorization: `Token ${TOKEN}` };

async function agendorGet(url) {
  const res = await fetch(url, { headers });
  if (!res.ok) {
    throw new Error(`Erro ${res.status} ao chamar ${url}: ${await res.text()}`);
  }
  return res.json();
}

async function main() {
  const tarefas = [];
  for (let page = 1; page <= 30; page++) {
    const url = `${BASE}/tasks?per_page=100&page=${page}&dueDateGt=${fmt(hoje)}&dueDateLt=${fmt(amanha)}`;
    const json = await agendorGet(url);
    tarefas.push(...json.data);
    if (json.data.length < 100) break;
  }

  // Filtra: apenas tarefas do tipo Reunião, vencimento exatamente no dia alvo (UTC date)
  const doDia = tarefas.filter(
    (t) => t.type === "Reunião" && t.dueDate && t.dueDate.slice(0, 10) === dataAlvo
  );

  const reunioes = [];
  for (const tarefa of doDia) {
    let org = null;
    let contatoPrincipal = null;

    if (tarefa.organization) {
      try {
        const orgData = await agendorGet(`${BASE}/organizations/${tarefa.organization.id}?withCustomFields=true`);
        org = orgData.data;
        if (org.people && org.people.length > 0) {
          try {
            const pessoaData = await agendorGet(`${BASE}/people/${org.people[0].id}`);
            contatoPrincipal = { nome: pessoaData.data.name, cargo: pessoaData.data.role };
          } catch (e) {
            contatoPrincipal = { nome: org.people[0].name, cargo: null };
          }
        }
      } catch (e) {
        console.error(`Aviso: não foi possível buscar organização ${tarefa.organization.id}: ${e.message}`);
      }
    }

    reunioes.push({
      taskId: tarefa.id,
      tipo: tarefa.type,
      texto: tarefa.text,
      dueDate: tarefa.dueDate,
      consultorResponsavel: tarefa.user ? tarefa.user.name : null,
      organizationId: tarefa.organization ? tarefa.organization.id : null,
      nomeFantasia: org ? org.name : (tarefa.organization ? tarefa.organization.name : null),
      razaoSocial: org ? org.legalName : null,
      cnpj: org ? org.cnpj : null,
      endereco: org
        ? { cidade: org.address.city, estado: org.address.state }
        : null,
      regimeTributario:
        org && org.customFields && org.customFields.regime_tributario
          ? org.customFields.regime_tributario.value
          : null,
      contatoPrincipal,
    });
  }

  fs.mkdirSync(path.dirname(outputPath), { recursive: true });
  fs.writeFileSync(outputPath, JSON.stringify({ data: dataAlvo, reunioes }, null, 2));
  console.log(`${reunioes.length} reunião(ões) encontradas para ${dataAlvo}.`);
  console.log(`Salvo em: ${outputPath}`);
}

main().catch((e) => {
  console.error(e);
  process.exit(1);
});

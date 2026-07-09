// Busca reuniões no Agendor por nome de empresa e monta um JSON intermediário
// com todos os dados necessários pra preencher a aba Lançamento_Reuniões.
//
// Uso: node --env-file=../../../../.env buscar-reunioes.js empresas.json saida.json
//
// empresas.json (entrada) — array de nomes de empresa como vieram do usuário:
// ["CONSTRUTORA CLARK LTDA", "XILOLITE S/A", ...]

const fs = require("fs");

const [, , entradaPath, saidaPath] = process.argv;
if (!entradaPath || !saidaPath) {
  console.error("Uso: node buscar-reunioes.js empresas.json saida.json");
  process.exit(1);
}

const TOKEN = process.env.AGENDOR_API_TOKEN;
if (!TOKEN) {
  console.error("AGENDOR_API_TOKEN não encontrado. Rode com: node --env-file=<caminho para .env> buscar-reunioes.js ...");
  process.exit(1);
}

const BASE = "https://api.agendor.com.br/v3";
const headers = { Authorization: `Token ${TOKEN}` };

// A API do Agendor não deixa buscar tarefas com mais de 31 dias de janela a partir de hoje.
function dataLimite() {
  const d = new Date();
  d.setDate(d.getDate() - 30);
  return d.toISOString().slice(0, 10);
}

async function agendorGet(url) {
  const res = await fetch(url, { headers });
  if (!res.ok) {
    return { error: `${res.status} ${await res.text()}` };
  }
  return res.json();
}

async function buscarOrganizacao(nome) {
  // Usar o parâmetro "name" (não "q" — "q" no /organizations não filtra corretamente).
  const r = await agendorGet(`${BASE}/organizations?name=${encodeURIComponent(nome)}`);
  if (r.error || !r.data || r.data.length === 0) return null;
  return r.data[0];
}

async function main() {
  const empresas = JSON.parse(fs.readFileSync(entradaPath, "utf-8"));
  const limite = dataLimite();
  const resultado = [];

  for (const nomeOriginal of empresas) {
    let org = await buscarOrganizacao(nomeOriginal);

    // Se não achar de primeira, tenta com só a primeira palavra significativa
    // (nomes muito longos ou com sufixo LTDA/S.A costumam falhar na busca exata).
    if (!org) {
      const termoCurto = nomeOriginal.split(/\s+/).slice(0, 2).join(" ");
      org = await buscarOrganizacao(termoCurto);
    }

    if (!org) {
      resultado.push({ nomeOriginal, encontrado: false });
      continue;
    }

    const tasksRes = await agendorGet(`${BASE}/organizations/${org.id}/tasks?dueDateGt=${limite}`);
    const reunioes = (tasksRes.data || []).filter((t) => t.type === "Reunião");

    resultado.push({
      nomeOriginal,
      encontrado: true,
      orgId: org.id,
      orgNome: org.name,
      cnpj: org.cnpj || null,
      segmento: org.sector ? org.sector.name : null,
      cidade: org.address ? org.address.city : null,
      estado: org.address ? org.address.state : null,
      reunioes: reunioes.map((t) => ({
        taskId: t.id,
        sdr: t.user ? t.user.name : null,
        dueDate: t.dueDate,
        finishedAt: t.finishedAt,
        texto: t.text,
      })),
    });

    await new Promise((r) => setTimeout(r, 150)); // não martelar a API
  }

  fs.writeFileSync(saidaPath, JSON.stringify(resultado, null, 2));
  const semResultado = resultado.filter((r) => !r.encontrado).length;
  console.log(`${resultado.length} empresa(s) processada(s), ${semResultado} sem organização encontrada no Agendor.`);
  console.log(`Salvo em: ${saidaPath}`);
}

main().catch((e) => {
  console.error(e);
  process.exit(1);
});

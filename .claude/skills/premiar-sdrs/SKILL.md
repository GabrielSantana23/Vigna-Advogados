---
name: premiar-sdrs
description: >
  Gera a planilha de premiação mensal dos SDRs do Grupo Vigna em Excel (.xlsx), com
  comissões de reuniões realizadas (NOVA/FUP × modalidade × senioridade) e negócios
  fechados (faixas progressivas para sênior, flat para júnior), separada por Equipe
  Gabriel e Equipe Paola. Usa dois exports do Agendor como entrada e consulta o
  histórico via API para classificar NOVA vs FUP automaticamente.
  Use quando o usuário pedir "premiação", "gerar premiação", "planilha de premiação",
  "calcular comissões do time", "premiar SDRs" ou "comissão do mês".
---

# /premiar-sdrs — Premiação Mensal dos SDRs

## O que essa skill faz

Gera automaticamente a planilha de premiação mensal dos SDRs do Grupo Vigna, combinando:
- **Reuniões realizadas** — classificadas como NOVA ou FUP via histórico do Agendor, com valor por modalidade (VC ou Presencial) e nível de senioridade
- **Negócios fechados** — com comissão por faixas progressivas (sênior) ou valor flat (júnior)

Saída: Excel com aba **Resumo**, aba **Negócios Fechados**, aba **Igor Vasconcelos** e uma aba por SDR com dropdowns interativos e fórmulas automáticas.

---

## Saída obrigatória

- **Formato:** `.xlsx`
- **Local:** `relatorios/premiacao/Premiacao_SDR_[MES]_[ANO].xlsx`
- **Nunca entregar:** tabela no chat, CSV ou texto corrido. Sempre o arquivo Excel.

---

## Tabela de valores

### Júnior (menos de 1 ano de casa)
| Tipo de comissão | Valor |
|------------------|-------|
| Reunião NOVA — Videoconferência | R$ 7,50 |
| Reunião NOVA — Presencial | R$ 15,00 |
| Reunião FUP — Videoconferência | R$ 5,00 |
| Reunião FUP — Presencial | R$ 7,50 |
| Negócio fechado (flat) | R$ 100,00 |

### Sênior (1 ano ou mais de casa)
| Tipo de comissão | Valor |
|------------------|-------|
| Reunião NOVA — Videoconferência | R$ 10,00 |
| Reunião NOVA — Presencial | R$ 25,00 |
| Reunião FUP — Videoconferência | R$ 5,00 |
| Reunião FUP — Presencial | R$ 10,00 |
| Negócios fechados 1–5 | R$ 120,00/cada |
| Negócios fechados 6–10 | R$ 150,00/cada |
| Negócios fechados 11+ | R$ 180,00/cada |

**SDRs sênior (+1 ano):** Gabriel, Ana Beatriz, Leticia

---

## Equipes e SDRs

**Equipe Gabriel:** Gabriel, Michelle, Kailany, Ana Karolayne, Igor (Souza Ferreira)

**Equipe Paola:** Larissa, Ana Belle, Leticia, Ana Beatriz, Lawanny, Railanne

**Negócios fechados (mapeamento extra):** Igor Vasconcelos (júnior, aba dedicada)

---

## Lógica NOVA vs FUP

- Busca globalmente no Agendor tarefas do tipo **Visita finalizadas** nos **3 meses anteriores** ao período (em lotes de 31 dias — limite da API)
- Se a empresa já teve Visita finalizada antes do período → **FUP**
- Se não: primeira ocorrência da empresa por SDR no período → **NOVA**; demais → **FUP**
- Se não tiver empresa vinculada na tarefa → **FUP** por padrão

---

## Workflow passo a passo

### Passo 1 — Confirmar o período com o usuário

Se não informado, perguntar:
> "Qual o período da premiação? (ex: 21/05 a 19/06)"

### Passo 2 — Exportar os dois arquivos do Agendor

Orientar o usuário a exportar:

**A) Reuniões realizadas:**
- Agendor → Relatórios → Tarefas → filtrar por tipo **Visita**, status **Finalizada**, período informado
- Exportar como `.xlsx` e salvar em `dados/` com nome descritivo
  - Exemplo: `dados/Reuniões realizada do dia 21 maio ate 19 junho.xlsx`

**B) Negócios fechados:**
- Agendor → Negócios → filtrar por status **Ganho**, período informado
- Exportar como `.xlsx` e salvar em `dados/`
  - Exemplo: `dados/Negócios fechados - MAIO - JUNHO.xlsx`

### Passo 3 — Atualizar o script para o novo período

Abrir `automacoes/premiar-sdr.js` e atualizar:

```javascript
// Linha ~57 — datas do período
const INICIO = new Date('AAAA-MM-DDT00:00:00.000Z');
const FIM    = new Date('AAAA-MM-DDT02:59:59.000Z'); // último dia 23:59 BRT = +3h UTC

// Linha ~186 — lotes de histórico (3 meses anteriores ao período)
const lotes = [
  ['AAAA-MM-DD', 'AAAA-MM-DD'], // mês -3
  ['AAAA-MM-DD', 'AAAA-MM-DD'], // mês -2
  ['AAAA-MM-DD', 'AAAA-MM-DD'], // mês -1
];
```

E atualizar os nomes dos arquivos de entrada:

```javascript
// Linha ~163 — arquivo de reuniões
await wb.xlsx.readFile('dados/NOME_DO_ARQUIVO_REUNIOES.xlsx');

// Linha ~237 — arquivo de negócios
await wb.xlsx.readFile('dados/NOME_DO_ARQUIVO_NEGOCIOS.xlsx');
```

E o nome do arquivo de saída:

```javascript
// Linha ~390 — nome do arquivo gerado
const outPath = path.join(outDir, 'Premiacao_SDR_MES_ANO.xlsx');
```

### Passo 4 — Rodar o script

```bash
node --env-file=.env automacoes/premiar-sdr.js
```

O script consulta o Agendor automaticamente (~30s) e gera o Excel.

### Passo 5 — Conferir e entregar

- Verificar o resumo no terminal (NOVA/FUP por SDR, totais)
- Abrir o arquivo gerado em `relatorios/premiacao/`
- Revisar os registros excluídos por fora do período (listados no terminal)
- Enviar para a diretoria

---

## Como atualizar a lista de SDRs

Se entrar ou sair alguém do time, editar em `automacoes/premiar-sdr.js`:

```javascript
// SDR_MAP — reuniões (nome do Agendor → nome curto)
const SDR_MAP = { ... };

// NEGOCIOS_SDR_MAP — negócios fechados
const NEGOCIOS_SDR_MAP = { ... };

// Equipes
const EQUIPE_GABRIEL = [...];
const EQUIPE_PAOLA   = [...];

// Sênior (+1 ano)
const SENIOR_SDRS = new Set(['Gabriel', 'Ana Beatriz', 'Leticia', /* adicionar aqui */]);
```

O nome exato do Agendor vem na primeira coluna do export (ex: `"Gabriel Santana Barra - Matriz SP"`).

---

## Como atualizar os valores de comissão

Editar em `automacoes/premiar-sdr.js`:

```javascript
// Tabela de reuniões (~linha 74)
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

// Faixas de fechamento (~linha 93)
function comissaoFechamento(total, isSenior) {
  if (!isSenior) return total * 100; // júnior: alterar aqui
  const t1 = Math.min(total, 5)                * 120; // bloco 1–5: alterar aqui
  const t2 = Math.max(0, Math.min(total-5, 5)) * 150; // bloco 6–10: alterar aqui
  const t3 = Math.max(0, total - 10)           * 180; // bloco 11+: alterar aqui
  return t1 + t2 + t3;
}
```

---

## Dependências

- **API:** `AGENDOR_API_TOKEN` em `.env` na raiz do projeto
- **Node.js:** disponível no sistema
- **ExcelJS:** instalado em `node_modules/` (já instalado)
- **Arquivos de entrada:** exports do Agendor salvos em `dados/`

---

## Arquivos relacionados

| Arquivo | Função |
|---------|--------|
| `automacoes/premiar-sdr.js` | Script principal — toda a lógica de premiação |
| `dados/Reuniões realizada*.xlsx` | Export de Visitas finalizadas do Agendor |
| `dados/Negócios fechados*.xlsx` | Export de Negócios ganhos do Agendor |
| `relatorios/premiacao/` | Pasta de saída dos Excels gerados |

---

## Regras importantes

- Nunca alterar a lógica de NOVA/FUP sem entender o histórico do Agendor — a classificação automática pode mudar se o histórico mudar
- Registros com data fora do período são excluídos automaticamente e listados no terminal — conferir antes de entregar
- Se o Agendor retornar erro na consulta de histórico, o script assume "sem histórico" (NOVA) e avisa no terminal
- Os dropdowns de TIPO e MODALIDADE no Excel são funcionais — ao trocar, VALOR recalcula automaticamente via fórmula

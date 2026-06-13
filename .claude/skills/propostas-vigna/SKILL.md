---
name: propostas-vigna
description: >
  Gera propostas comerciais das empresas do Grupo Vigna (Vigna Advogados, VignaTax, Compliance
  Control, Legal Control) direto no Canva. Recebe um relatório/transcrição de reunião com
  advogado (colado no chat) ou uma solicitação direta, identifica qual(is) proposta(s) fazer,
  busca os dados do lead via CNPJ, seleciona o template certo no Canva e edita a cópia. Use
  quando o usuário colar um relatório de reunião com lead, pedir "proposta", "proposta
  comercial", "PRT_XXX" ou mencionar algum dos serviços do catálogo abaixo.
---

# /propostas-vigna — Geração de Propostas Comerciais

## O que essa skill faz

Recebe um relatório/transcrição de advogado (colado no chat) ou uma solicitação direta,
identifica automaticamente qual(is) proposta(s) precisam ser feitas, seleciona o template certo
no Canva e executa a edição.

## Dependências

- **Canva MCP** — já conectado (ver "Ferramentas conectadas" no CLAUDE.md)
- `references/portfolio.md` — serviços e persona ideal de cada empresa do grupo
- `references/objetos-exemplo.md` — exemplos de objeto, pra quando for necessário customizar

---

## Catálogo completo de templates no Canva

Todos os IDs abaixo são `design_id` pra usar diretamente com `start-editing-transaction`. **Não
é necessário resolver shortlinks** — use os IDs diretamente.

### 🔵 Vigna Advogados Associados (VAA)

Pasta: `FAHLEvgbxj8`

| Template | design_id | Palavras-chave pra reconhecer |
|---|---|---|
| PROPOSTA AÇÃO ESPECÍFICA | `DAHC097433U` | ação específica, demanda pontual, ajuizamento |
| PROPOSTA CONSULTIVO FEE MENSAL | `DAHC0I9ozYI` | consultivo, fee mensal, assessoria mensal, retainer |
| PROPOSTA CONSULTIVO TIME SHEET | `DAHCz5kX4Ys` | time sheet, horas, consultivo por hora |
| PROPOSTA CONTENCIOSO CÍVEL - ESCALONADO | `DAHC04b78BA` | contencioso cível, cível, escalonado |
| PROPOSTA CONTENCIOSO TRABALHISTA - ESCALONADO | `DAHC03YfrDM` | trabalhista, contencioso trabalhista, CLT, escalonado |
| PROPOSTA CONTRATUAL TABELA | `DAHCy0Wz-LY` | contratual, tabela, contratos |
| PROPOSTA DE AUDITORIA / SANEAMENTO | `DAHC0-xzmLA` | auditoria, saneamento, due diligence jurídica |
| PROPOSTA INVESTIGAÇÃO PATRIMONIAL/DOSSIÊ | `DAHLEoQHhY0` | investigação patrimonial, dossiê, patrimônio |
| PROPOSTA PARECER | `DAHK-smaiSs` | parecer, opinião legal, consulta jurídica |
| PROPOSTA PROPRIEDADE INTELECTUAL - Registro de Marca | `DAHC1SP-WBM` | propriedade intelectual, marca, registro de marca, INPI |
| PROPOSTA RECUPERAÇÃO DE CRÉDITO JUDICIAL E EXTRAJUDICIAL | `DAHK-kqheak` | recuperação de crédito, crédito judicial, crédito extrajudicial |

### 🟢 Compliance Control (CC)

Pasta: `FAHLEQQ4rAE`

| Template | design_id | Palavras-chave pra reconhecer |
|---|---|---|
| INVESTIGAÇÃO CORPORATIVA | `DAHK88Xb4ic` | investigação corporativa, fraude interna, apuração |
| CÓDIGO DE CONDUTA | `DAHLEALxU6I` | código de conduta, ética empresarial |
| GESTÃO DO CANAL DE DENÚNCIAS | `DAHLEKsZRw8` | canal de denúncias, compliance, denúncias |
| TREINAMENTOS | `DAHK4yx25wg` | treinamento, capacitação, workshop compliance |
| IMPLEMENTAÇÃO COMPLIANCE | `DAHHbCdNFz0` | implementação compliance, programa de compliance, LGPD, GRC |
| ELABORAÇÃO DE POLÍTICA | `DAHK36SGUCI` | política interna, elaboração de política, normas internas |

### 🟡 VignaTax

Pasta: `FAHLEbmJAEk`

| Template | design_id | Palavras-chave pra reconhecer |
|---|---|---|
| COMPLIANCE PREVIDENCIÁRIO | `DAHLEXeqIpM` | compliance previdenciário, previdência, INSS, contribuições |
| DIAGNÓSTICO DA REFORMA TRIBUTÁRIA | `DAHLEekTdWc` | diagnóstico reforma tributária, impacto tributário |
| MS - Ilegalidade da Lei 12.808/2025 | `DAHLERQvp98` | mandado de segurança, Lei 12.808, ilegalidade |
| REVISÃO FISCAL COMPLETA | `DAHLEYxy5ro` | revisão fiscal completa, revisão tributária |
| REVISÃO FISCAL PIS E COFINS | `DAHKrdZv60I` | PIS, COFINS, revisão PIS/COFINS |
| REFORMA TRIBUTÁRIA COMPLETA | `DAHLEfz-F04` | reforma tributária, IVA, CBS, IBS |

### 🔴 Legal Control (LC)

Pasta: `FAHLEjGk67s`

| Template | design_id | Palavras-chave pra reconhecer |
|---|---|---|
| CANAL DE DENÚNCIAS - LEGAL ÉTICA | `DAHLEnuxeGo` | Legal Ética, canal de denúncias software, ética digital |
| NR-1+ | `DAHKZ1qopi0` | NR-1, NR1+, saúde mental, psicossocial, riscos ocupacionais |
| LEGAL TRAINING | `DAHD1RilV8s` | Legal Training, treinamento jurídico, capacitação jurídica |
| LEGAL AUDIT | `DAGfwTa54Kk` | Legal Audit, auditoria jurídica, audit |

---

## Fluxo automático — relatório do advogado

### Passo 1 — Ler e identificar as propostas

Quando o usuário colar um relatório/transcrição de advogado:

1. Leia o texto integralmente
2. Identifique cada proposta mencionada (pode ser mais de uma)
3. Para cada proposta, determine qual empresa do Grupo Vigna e qual template usar (consulte o
   catálogo acima e `references/portfolio.md` em caso de dúvida sobre o serviço)
4. Extraia do texto os dados do lead disponíveis: nome da empresa, CNPJ, contato

### Passo 2 — Buscar dados da empresa automaticamente

⚠️ **Obrigatório:** sempre consulte a API pública de CNPJ pra obter razão social, endereço e
dados cadastrais completos. Não dependa só do que está no relatório — o texto pode ter nome
fantasia, abreviações ou dados incompletos.

**Se o CNPJ estiver no relatório:**

```
GET https://publica.cnpj.ws/cnpj/{CNPJ_SEM_FORMATACAO}
```

Extraia: `razao_social`, `cnpj`, `logradouro`, `numero`, `complemento`, `bairro`, `municipio`,
`uf`, `cep`.

**Se só o nome da empresa estiver no relatório (sem CNPJ):**

1. Use web search pra encontrar o CNPJ: `"[nome da empresa]" CNPJ site:receitafederal.gov.br`
   ou `"[nome da empresa]" CNPJ razão social`
2. Depois consulte a API com o CNPJ encontrado

**Se não conseguir encontrar o CNPJ:** pergunte ao usuário apenas o CNPJ — não peça razão social
nem endereço (a API retorna tudo).

Após a consulta, apresente um resumo rápido antes de agir:

```
📋 Identifiquei as seguintes propostas:
1. [Empresa Vigna] → [Nome do Template]
   Lead: [Razão Social] | CNPJ: [XX.XXX.XXX/XXXX-XX]
   Endereço: [logradouro, número, bairro, cidade/UF]

Preciso confirmar:
- Numeração(ões) da(s) proposta(s): ?
- Nome do contato/responsável (Att.): ? ← se não estiver no relatório
```

### Passo 3 — Coletar apenas o que falta

Após a busca automática, pergunte somente:

- **Número da proposta** (ex: `PRT_058_26`) — **obrigatório, sempre perguntar**. O prefixo é
  sempre `PRT_`, independente da empresa do Grupo Vigna.
- **Nome do contato (Att.)** — se não mencionado no relatório
- A data é sempre a data atual, salvo indicação contrária
- **Número de colaboradores/vidas/usuários** — obrigatório para propostas de sistemas/plataformas
  (Legal Control: NR-1+, Legal Ética, Legal Training, Legal Audit; e demais propostas da
  Compliance Control com cobrança por vida/usuário). Esse número alimenta a tabela de condições
  comerciais (ex: página de "Proposta para X vidas"). Se não estiver no relatório, perguntar.

### Passo 4 — Formatar dados do lead

Monte com os dados obtidos da API + relatório:

```
[RAZÃO SOCIAL — conforme Receita Federal]
CNPJ: [XX.XXX.XXX/XXXX-XX]
[Logradouro, Nº, Complemento — Bairro — CEP]
Att.: [Nome do contato]
[Município/UF], [Data por extenso]
```

### Passo 5 — Gerar objeto de proposta (somente quando necessário)

⚠️ **Regra padrão: não criar objeto do zero.** Os templates já têm objeto embutido.

- **VignaTax, Legal Control, Compliance Control:** não criar objeto. Usar o do template.
- **Vigna Advogados:** perguntar ao usuário se quer customizar o objeto.
- Se o usuário pedir explicitamente, criar conforme solicitado.

Leia `references/objetos-exemplo.md` pra estilo e formato quando necessário.

### Passo 6 — Editar no Canva

Para cada proposta identificada, siga obrigatoriamente esta sequência:

**6.1 — Copiar o template (sempre)**

⚠️ Nunca edite o template original. Use `copy-design` com o `design_id` do catálogo pra criar
uma cópia limpa.

**6.2 — Renomear a cópia**

Use `update_title` (dentro de `perform-editing-operations`) com o padrão:

```
PRT_[NÚMERO]_26 - [RAZÃO SOCIAL ABREVIADA] - [SERVIÇO]
```

Exemplos:
- `PRT_391_26 - CERBRAS - REVISÃO FISCAL`
- `PRT_392_26 - CERBRAS - COMPLIANCE PREVIDENCIÁRIO`
- `PRT_393_26 - CERBRAS - REFORMA TRIBUTÁRIA`
- `PRT_058_26 - VIGNA ADV - CONTENCIOSO TRABALHISTA`
- `PRT_059_26 - ILUMEO - CANAL DE DENÚNCIAS`

⚠️ O prefixo é sempre `PRT_` para qualquer empresa do Grupo Vigna.

**6.3 — Aplicar as substituições de dados**

Na cópia criada, abra com `start-editing-transaction` e aplique via `find_and_replace_text`:

- Número da proposta → número informado pelo usuário
- Nome da empresa anterior → razão social do lead
- CNPJ anterior → CNPJ do lead
- Data anterior → data atual
- Endereço anterior → endereço do lead
- Att./contato anterior → contato do lead (se disponível)

**6.4 — Mostrar preview e salvar**

Mostre o preview (thumbnail da página 1) ao usuário e salve com `commit-editing-transaction`.

**6.5 — Disponibilizar o link editável no chat**

Após salvar, sempre exiba o link de edição direta:

```
✅ [PRT_XXX_26 - EMPRESA - SERVIÇO](edit_url)
```

Use a `edit_url` retornada por `get-design` ou `start-editing-transaction`.

**6.6 — Repetir para cada proposta adicional**

---

## Regras gerais

- Nunca mencione valores ou honorários no objeto
- Número da proposta é **obrigatório** — nunca edite o Canva sem ele. Prefixo sempre `PRT_`
- **Sempre** copiar o template — jamais editar o original diretamente
- **Sempre** renomear a cópia com o padrão `PRT_XXX_26 - EMPRESA - SERVIÇO`
- **Sempre** disponibilizar o link editável no chat após salvar cada proposta
- Objeto só é criado quando solicitado ou quando for VAA (perguntar)
- Se houver múltiplas propostas no mesmo relatório, processe uma de cada vez
- Se a proposta for ambígua entre dois templates, aponte as duas opções e deixe o usuário
  escolher
- Leia `references/portfolio.md` pra dúvidas sobre serviços ou personas de cada empresa do grupo

---

## Output padrão (quando não há relatório — solicitação direta)

```
Empresa Vigna: [nome]
Serviço: [serviço específico]
Template: [nome do template identificado]
Persona: [cargo ideal para abordar]

DADOS DO LEAD (para o Canva)
[dados formatados]

Objeto de proposta: [somente se necessário/solicitado]
Próximo passo: [confirmar numeração e editar no Canva]
```

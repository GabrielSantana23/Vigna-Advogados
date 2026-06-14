---
name: briefing-reunioes
description: >
  Gera o Briefing Pré-Reunião do Grupo Vigna em Excel (.xlsx), com uma aba por
  empresa/reunião, contendo dados cadastrais, resumo executivo da empresa (pesquisa
  pública) e oportunidades comerciais priorizadas por persona. Busca as reuniões do dia
  seguinte direto no Agendor (CRM). Use quando o usuário pedir "briefing", "briefing de
  reunião", "briefing pré-reunião", "preparar as reuniões de amanhã" ou pedir o briefing
  de uma data específica.
---

# /briefing-reunioes — Briefing Pré-Reunião Grupo Vigna

## O que essa skill faz

No fim do dia, monta um briefing executivo pra preparar os consultores Vigna pras
reuniões do dia seguinte (ou de outra data informada): dados cadastrais da empresa,
resumo executivo baseado em pesquisa pública e oportunidades comerciais priorizadas pra
cada empresa do Grupo Vigna.

## Saída obrigatória

- **Formato:** `.xlsx`
- **Local:** `relatorios/briefing-reunioes/Briefing_DDMMAAAA.xlsx` (DDMMAAAA = data das reuniões)
- **Uma aba (worksheet) por empresa/reunião** — nunca consolidar tudo numa aba só
- **Nunca entregar:** HTML, markdown, tabela no chat, JSON, CSV ou texto corrido. O
  entregável é sempre o arquivo Excel.

---

## Estrutura de cada aba

Duas colunas: **Campo** | **Informação**. Cada aba tem exatamente estas 12 linhas, nesta ordem:

| # | Campo | Como obter |
|---|-------|-----------|
| 1 | Consultores Vigna | campo `Participação:` do relatório (separar nomes por `/` ou "e") |
| 2 | Horário Reunião | campo `Hora:` + cidade/UF da matriz da empresa, ex: `14h00 — Campinas/SP` |
| 3 | Nome da Empresa | `razaoSocial` do Agendor (fallback `nomeFantasia` ou linha `Reunião: Pessoa Jurídica -`) |
| 4 | CNPJ | `cnpj` do Agendor, formatado `00.000.000/0001-00` |
| 5 | Responsável da Empresa | campo `Reunião: Com a Sr(a). [Nome]` — remover "Sr.", "Sra.", "responsável" etc, manter só o nome |
| 6 | Cargo | mesmo campo, ou `contatoPrincipal.cargo` do Agendor |
| 7 | Resumo da Empresa | pesquisa pública obrigatória — ver regras abaixo |
| 8 | Oportunidades | matriz estratégica (`references/matriz-oportunidades.md`), 3 a 6 itens |
| 9 | Regime Tributário | customField `regimeTributario` do Agendor, ou pesquisa pública |
| 10 | Site da Empresa | pesquisa pública |
| 11 | Pauta | campo `Pauta:` do relatório (resumir se for muito longo) |
| 12 | Link da Reunião | campo `Local:`, se for link de Teams/Meet/Zoom |

**Regra absoluta:** quando não localizar uma informação, usar exatamente `Não localizado`
— nunca inventar CNPJ, site, regime tributário, dados financeiros, filiais ou
oportunidades genéricas.

---

## Fluxo de execução

1. **Determinar a data alvo.** Como o briefing é feito no fim do dia, a data alvo padrão é
   o **dia seguinte** ao de hoje. Se o usuário informar outra data, usar essa.
2. **Buscar as reuniões no Agendor** rodando `buscar-agendor.js` (ver "Busca no Agendor"
   abaixo). Se o usuário colar manualmente o relatório de uma reunião, pular esse passo e
   extrair os dados do texto colado.
3. **Filtrar reuniões inválidas** — igual ao `/quadro-visitas`:
   - Se o `texto` contiver `DECLINADA`, `CANCELADA`, `CANCELOU` ou equivalente → excluir
   - Se o campo `Data:` do texto existir e for diferente da data alvo → excluir
4. **Para cada reunião válida**, extrair do campo `texto` os campos da seção "Extração
   automática" abaixo.
5. **Pesquisar publicamente** (WebSearch/WebFetch) o que faltar: resumo da empresa (sempre
   obrigatório), site oficial, regime tributário (se vier `null` do Agendor).
6. **Selecionar as oportunidades** (3 a 6) cruzando cargo do contato + segmento da empresa
   com `references/matriz-oportunidades.md`.
7. **Montar o JSON de dados** no formato esperado pelo gerador (ver "Geração do arquivo").
8. **Gerar o arquivo `.xlsx`** rodando `gerar.js`.
9. Salvar em `relatorios/briefing-reunioes/Briefing_DDMMAAAA.xlsx` e apagar os arquivos
   temporários (`agendor-DDMMAAAA.json` e `dados-DDMMAAAA.json`).
10. Confirmar pro usuário onde o arquivo foi salvo e quantas reuniões/abas foram geradas.

---

## Busca no Agendor

Reaproveita o mesmo script do `/quadro-visitas`, que já busca as tarefas do tipo "Reunião"
no Agendor e resolve organização (CNPJ, razão social, endereço, regime tributário) e
contato principal.

**Token:** fica em `.env` na raiz do projeto, na variável `AGENDOR_API_TOKEN`. Não expor
esse valor em nenhum output.

### Como rodar

```bash
cd .claude/skills/briefing-reunioes/scripts
node --env-file=../../../../.env ../../quadro-visitas/scripts/buscar-agendor.js DD/MM/AAAA agendor-DDMMAAAA.json
```

Formato de saída: ver `.claude/skills/quadro-visitas/SKILL.md` (seção "Busca no Agendor")
— mesmo JSON (`reunioes[]` com `texto`, `cnpj`, `razaoSocial`, `regimeTributario`,
`endereco`, `contatoPrincipal`, `consultorResponsavel` etc).

---

## Extração automática de campos

A partir do campo `texto` de cada reunião (formato livre, vindo do Agendor):

- **Nome da Empresa** — linha `Reunião: Pessoa Jurídica - [Nome]`
- **Consultores Vigna** — campo `Participação:` (separar corretamente os nomes)
- **Responsável da Empresa** — campo `Reunião: Com a Sr(a). [Nome] ([Cargo])`. Remover
  automaticamente "Sr.", "Sra.", "responsável" e textos desnecessários — manter só o nome
- **Cargo** — extraído do mesmo campo da reunião
- **Horário Reunião** — campo `Hora:`, normalizado pra `HHhMM` + cidade/UF da matriz da
  empresa (endereço do Agendor)
- **Pauta** — campo `Pauta:`

---

## Pesquisa pública obrigatória — Resumo da Empresa

O campo **Resumo da Empresa** nunca pode ficar vazio. Pesquisar (WebSearch/WebFetch) e
gerar um resumo executivo real, **entre 80 e 200 palavras**.

**Fontes prioritárias:** site oficial, LinkedIn da empresa, Receita Federal (CNAE,
situação cadastral), notícias públicas, apresentações institucionais, materiais
corporativos.

**O resumo deve conter:** o que a empresa faz, segmento de atuação, principais
produtos/serviços, mercado atendido, tempo de atuação (se localizado), diferenciais
operacionais, estrutura percebida, sinais de maturidade/crescimento — sempre com foco em
preparar o consultor pra reunião.

**Estilo:** consultivo, executivo, objetivo, linguagem profissional. Sem parecer texto de
IA, propaganda ou opinião pessoal. Sem frases genéricas, sem exagero, sem inventar dados.

**Site da Empresa** e **Regime Tributário** seguem a mesma regra: pesquisa pública,
`Não localizado` se não encontrar — nunca inventar.

---

## Oportunidades — seleção

Ver `references/matriz-oportunidades.md` pra matriz completa por empresa do Grupo Vigna
(Vigna Advogados, VignaTax, Compliance Control, Legal Control), com personas prioritárias
e oportunidades por empresa.

**Regras:**
- 3 a 6 oportunidades por reunião, formato numerado: `1. Oportunidade; 2. Oportunidade; ...`
- Priorizar pelo cargo do contato e pelo segmento/porte da empresa
- Cross-sell é bem-vindo: uma mesma reunião pode gerar oportunidades de mais de uma
  empresa do Grupo Vigna
- Em caso de dúvida sobre o serviço, consultar também
  `.claude/skills/propostas-vigna/references/portfolio.md`

---

## Geração do arquivo (Node.js + exceljs)

O gerador está em `.claude/skills/briefing-reunioes/scripts/gerar.js` e implementa a
formatação visual (uma aba por reunião, faixa de título, cabeçalho, linhas alternadas,
bordas, congelamento de cabeçalho, rodapé).

### Passo a passo

1. Montar um JSON com os dados do dia, no formato:

```json
{
  "data": "DD/MM/AAAA",
  "reunioes": [
    {
      "consultores": "Cleber e Rafael M.",
      "horario": "14h00 — Campinas/SP",
      "empresa": "Empresa XYZ Ltda",
      "cnpj": "00.000.000/0001-00",
      "responsavel": "João Silva",
      "cargo": "Controller",
      "resumo": "Resumo executivo de 80 a 200 palavras...",
      "oportunidades": "1. Revisão fiscal; 2. LGPD; 3. Compliance trabalhista",
      "regime": "Lucro Presumido",
      "site": "empresa.com.br",
      "pauta": "Apresentação institucional",
      "link": "Não localizado"
    }
  ]
}
```

2. Salvar esse JSON em um arquivo temporário (ex:
   `.claude/skills/briefing-reunioes/scripts/dados-DDMMAAAA.json`)

3. Rodar o gerador:

```bash
cd .claude/skills/briefing-reunioes/scripts
node gerar.js dados-DDMMAAAA.json ../../../../relatorios/briefing-reunioes/Briefing_DDMMAAAA.xlsx
```

4. Apagar o JSON temporário depois de gerar o arquivo.

---

## Confiabilidade

- CNPJ, razão social e regime tributário vêm do Agendor quando disponíveis — não pesquisar
  se já vierem preenchidos
- Resumo da Empresa, Site e Regime Tributário (quando faltando) sempre via pesquisa
  pública — nunca inventar
- Se faltar algum dado (consultor, cargo, pauta, link etc), usar `Não localizado`
- Nunca expor o conteúdo de `AGENDOR_API_TOKEN` em outputs, mensagens ou arquivos
  versionados

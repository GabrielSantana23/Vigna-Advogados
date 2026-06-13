---
name: quadro-visitas
description: >
  Gera o Quadro de Visitas diário do Grupo Vigna em Excel (.xlsx), compacto e pronto
  para envio por e-mail, com as reuniões do dia agrupadas por Unidade e ordenadas por
  horário. Use quando o usuário pedir "quadro de visitas", "planilha de visitas do dia",
  "agenda do dia em planilha" ou colar a lista de reuniões/agenda do dia.
---

# /quadro-visitas — Quadro de Visitas Grupo Vigna

## Saída obrigatória

- **Formato:** `.xlsx`
- **Local:** `relatorios/quadro-visitas/Quadro_Visitas_Vigna_DDMMAAAA.xlsx` (DDMMAAAA = data das reuniões)
- **Nunca entregar:** HTML, markdown, tabela no chat, JSON, CSV ou texto corrido. O entregável é sempre o arquivo Excel.

---

## Colunas da planilha

| # | Coluna | Observação |
|---|--------|------------|
| 1 | Unidade | Cor sólida por unidade; texto branco negrito centralizado |
| 2 | Horário | Negrito; centralizado |
| 3 | Razão Social | Negrito; alinhado à esquerda |
| 4 | CNPJ | Centralizado |
| 5 | Pessoa da Empresa | Nome do contato |
| 6 | Cargo | — |
| 7 | Consultor Vigna | Separados por `/` se mais de um |
| 8 | Regime Tributário | Centralizado |
| 9 | Oportunidades | Máximo 2, formato `① Título curto (Serviço Vigna)` |

⚠️ Não incluir coluna "Resumo Estratégico" — torna o arquivo pesado para envio por e-mail.

---

## Ordenação obrigatória

- Agrupar por Unidade na ordem: **ALPHAVILLE → CAMPINAS → MG → PR → RS → MATRIZ**
- Dentro de cada unidade: ordenar por horário crescente
- Nunca misturar linhas de unidades diferentes

---

## Formatação visual

### Cabeçalho (linha 5)
- Fundo: `#1B2A4A` (azul escuro)
- Fonte: Arial, 8.5pt, branco, negrito, centralizado
- Altura da linha: 24px

### Faixa de título (linhas 2–3)
- Linha 2: `"GRUPO VIGNA  ·  QUADRO DE VISITAS  ·  [DIA DD/MM/AAAA]"` — fundo `#1B2A4A`, fonte Georgia 12pt, branca, negrito
- Linha 3: subtítulo em cinza-azul sobre fundo `#131F38`
- Linhas 1 e 4 ficam em branco (espaçamento)

### Cores por unidade (coluna 1)
| Unidade | Cor | Hex |
|---------|-----|-----|
| ALPHAVILLE | roxo escuro | `#4A235A` |
| CAMPINAS | azul marinho | `#1A3A6A` |
| MG | verde escuro | `#1A5C3A` |
| PR | marrom/terra | `#7A3B00` |
| RS | vinho | `#8B1A1A` |
| MATRIZ | azul padrão | `#1B2A4A` |

### Linhas de dados
- Altura: 38px (compacto pra e-mail)
- Alternância de fundo: `#EAF0FB` (par) / `#F5F6FA` (ímpar)
- Alinhamento vertical: centralizado
- `wrap_text=True` em todas as células
- Coluna Oportunidades: fundo `#FAFBFF`, fonte 8pt

### Configurações gerais
- `showGridLines = False`
- `freeze_panes = "A6"` (congela cabeçalho)
- `zoomScale = 95`
- Bordas finas `#D8DFF0` em todas as células

### Larguras de coluna (aproximadas)
Unidade=14 | Horário=8 | Razão Social=34 | CNPJ=20 | Pessoa=16 | Cargo=14 | Consultor=28 | Regime=14 | Oportunidades=64

### Rodapé
- Texto: `"Documento confidencial · Uso interno · Grupo Vigna © AAAA"`
- Fonte 7.5pt itálico, cor `#8899BB`, fundo `#F0F3FA`

---

## Oportunidades — regra

- Máximo 2 por empresa
- Formato obrigatório: `① Título curto (Serviço Vigna)`
- Separadas por `\n` dentro da célula
- Serviços possíveis: Vigna Advogados, VignaTax, Compliance Control, Legal Control
- Priorizar pela persona do contato (cargo), conforme a matriz abaixo

### Linhas adicionais (sempre incluir)

Abaixo das oportunidades, na mesma célula, sempre incluir mais duas linhas:

- `Pauta: [resumo curto da pauta da reunião]`
- `Link: [link da reunião — Teams/Meet/Zoom, se houver]`

Se o link não foi informado, usar `Link: Não informado`.

### Matriz de priorização por cargo

| Cargo | Prioridades |
|-------|-------------|
| Diretor / CEO / CFO | Blindagem patrimonial, Recuperação Tributária, Planejamento Sucessório |
| Jurídico / Advogado | Compliance Trabalhista, Defesa Trabalhista, LGPD |
| Tributário / Controller | RAT/FAP, ICMS-ST, PIS/COFINS, Sistema SRH / DPNR-1, Auditoria Trabalhista, Compliance Trabalhista |
| Compliance | Due Diligence, Matriz de Riscos, ESG, LGPD |

---

## Regra de unidade

Sempre usar a cidade da **MATRIZ** da empresa — nunca filial, DDD do contato ou local da reunião.

| Estado / Região | Unidade |
|-----------------|---------|
| MG | MG |
| PR | PR |
| RS | RS |
| SP — cidades da região de Alphaville/Grande SP (ex: Barueri, Osasco, Santana de Parnaíba, Itapevi, Jandira, Carapicuíba, Cotia) | ALPHAVILLE |
| SP — DDD 12/13/14/15/16/17/18/19 (fora da região Alphaville) | CAMPINAS |
| SP — demais cidades / Capital | MATRIZ |
| Outros estados não mapeados | MATRIZ |

Se a lista acima não for suficiente pra classificar uma empresa, usar o melhor julgamento com base na cidade da matriz e seguir. Pode-se editar esta tabela depois pra refinar a lista de cidades.

---

## Fluxo de execução

1. **Extrair dados** das reuniões a partir do texto/tabela que o usuário colar (empresa, horário, contato, cargo, consultor responsável)
2. **Pesquisar CNPJ e regime tributário** de cada empresa em fontes públicas (WebSearch/WebFetch). Se não localizar: `"Não localizado"` — nunca inventar
3. **Definir a unidade** de cada empresa via cidade da matriz, usando a tabela de regra de unidade acima
4. **Ordenar** por unidade (ALPHAVILLE → CAMPINAS → MG → PR → RS → MATRIZ) e depois por horário crescente
5. **Selecionar até 2 oportunidades** por empresa com base no cargo do contato, usando a matriz de priorização
6. **Montar o JSON de dados** no formato esperado pelo gerador (ver abaixo)
7. **Gerar o arquivo `.xlsx`** rodando o script Node, seguindo todas as especificações de formatação acima
8. Salvar em `relatorios/quadro-visitas/Quadro_Visitas_Vigna_DDMMAAAA.xlsx`
9. Confirmar pro usuário que o arquivo foi gerado e onde está salvo

---

## Geração do arquivo (Node.js + exceljs)

Essa máquina não tem Python instalado, então a geração do `.xlsx` é feita com Node.js (já disponível) e a lib `exceljs` (já instalada em `scripts/node_modules`).

O gerador está em `.claude/skills/quadro-visitas/scripts/gerar.js` e já implementa toda a formatação visual descrita acima (cores por unidade, faixa de título, cabeçalho, alternância de linhas, bordas, freeze panes, rodapé etc).

### Passo a passo

1. Montar um JSON com os dados do dia, no formato:

```json
{
  "data": "DD/MM/AAAA",
  "linhas": [
    {
      "unidade": "MATRIZ",
      "horario": "14:00",
      "razao": "RAZÃO SOCIAL DA EMPRESA",
      "cnpj": "00.000.000/0001-00",
      "pessoa": "Nome do Contato",
      "cargo": "Cargo do Contato",
      "consultor": "Nome do Consultor / Outro Consultor",
      "regime": "Regime tributário ou 'Não localizado'",
      "oportunidades": [
        "① Título curto (Serviço Vigna)",
        "② Título curto (Serviço Vigna)"
      ],
      "pauta": "Resumo curto da pauta da reunião",
      "link": "Link da reunião (Teams/Meet/Zoom) ou 'Não informado'"
    }
  ]
}
```

`linhas` já deve estar na ordem final (unidade → horário). `unidade` deve ser um dos valores: `ALPHAVILLE`, `CAMPINAS`, `MG`, `PR`, `RS`, `MATRIZ`.

2. Salvar esse JSON em um arquivo temporário (ex: `.claude/skills/quadro-visitas/scripts/dados-DDMMAAAA.json`)

3. Rodar o gerador:

```bash
cd .claude/skills/quadro-visitas/scripts
node gerar.js dados-DDMMAAAA.json ../../../../relatorios/quadro-visitas/Quadro_Visitas_Vigna_DDMMAAAA.xlsx
```

4. Apagar o JSON temporário depois de gerar o arquivo (não precisa manter histórico de dados, só o `.xlsx` final)

---

## Confiabilidade

- Usar apenas informações públicas verificáveis (CNPJ, razão social, regime tributário)
- Se CNPJ ou regime não localizado: `"Não localizado"`
- Nunca inventar dados financeiros, filiais ou estrutura societária
- Se faltar algum dado da reunião (consultor, cargo, etc), deixar em branco — não inventar

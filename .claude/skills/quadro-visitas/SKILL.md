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
6. **Gerar o arquivo `.xlsx`** com openpyxl, seguindo todas as especificações de formatação acima
7. Salvar em `relatorios/quadro-visitas/Quadro_Visitas_Vigna_DDMMAAAA.xlsx`
8. Confirmar pro usuário que o arquivo foi gerado e onde está salvo

---

## Template de código Python (openpyxl)

Usar como base — adaptar com os dados extraídos do dia:

```python
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from datetime import datetime

wb = Workbook()
ws = wb.active
ws.title = "Quadro de Visitas"
ws.sheet_view.showGridLines = False
ws.sheet_view.zoomScale = 95

# Paleta
AZUL_ESCURO = "1B2A4A"
BRANCO = "FFFFFF"
BORDA = "D8DFF0"
UNIDADE_CORES = {
    "ALPHAVILLE": "4A235A", "CAMPINAS": "1A3A6A",
    "MG": "1A5C3A", "PR": "7A3B00", "RS": "8B1A1A", "MATRIZ": "1B2A4A",
}
ORDEM_UNIDADES = ["ALPHAVILLE", "CAMPINAS", "MG", "PR", "RS", "MATRIZ"]

thin = Side(border_style="thin", color=BORDA)
borda = Border(left=thin, right=thin, top=thin, bottom=thin)

# Larguras
col_widths = [14, 8, 34, 20, 16, 14, 28, 14, 64]
for i, w in enumerate(col_widths, 1):
    ws.column_dimensions[get_column_letter(i)].width = w

n_cols = len(col_widths)

# Linha 2 — título
ws.merge_cells(start_row=2, start_column=1, end_row=2, end_column=n_cols)
data_str = datetime.now().strftime("%d/%m/%Y")  # usar a data das reuniões, não a data de hoje
titulo = ws.cell(row=2, column=1, value=f"GRUPO VIGNA  ·  QUADRO DE VISITAS  ·  {data_str}")
titulo.font = Font(name="Georgia", size=12, bold=True, color=BRANCO)
titulo.fill = PatternFill("solid", fgColor=AZUL_ESCURO)
titulo.alignment = Alignment(horizontal="center", vertical="center")
ws.row_dimensions[2].height = 28

# Linha 3 — subtítulo
ws.merge_cells(start_row=3, start_column=1, end_row=3, end_column=n_cols)
sub = ws.cell(row=3, column=1, value="Reuniões agendadas para o dia, por unidade e horário")
sub.font = Font(name="Arial", size=9, color="A9B6D6")
sub.fill = PatternFill("solid", fgColor="131F38")
sub.alignment = Alignment(horizontal="center", vertical="center")
ws.row_dimensions[3].height = 20

# Linha 5 — cabeçalho
labels = ["UNIDADE", "HORÁRIO", "RAZÃO SOCIAL", "CNPJ", "CONTATO",
          "CARGO", "CONSULTOR VIGNA", "REGIME", "OPORTUNIDADES PRIORITÁRIAS"]
for i, label in enumerate(labels, 1):
    c = ws.cell(row=5, column=i, value=label)
    c.font = Font(name="Arial", size=8.5, bold=True, color=BRANCO)
    c.fill = PatternFill("solid", fgColor=AZUL_ESCURO)
    c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    c.border = borda
ws.row_dimensions[5].height = 24

# Dados — `linhas` é uma lista de dicts já ordenada por unidade > horário
# cada dict: unidade, horario, razao, cnpj, pessoa, cargo, consultor, regime, oportunidades (string com \n)
linha_atual = 6
for idx, item in enumerate(linhas):
    fundo = "EAF0FB" if idx % 2 == 0 else "F5F6FA"
    valores = [item["unidade"], item["horario"], item["razao"], item["cnpj"],
               item["pessoa"], item["cargo"], item["consultor"], item["regime"],
               item["oportunidades"]]
    for col, valor in enumerate(valores, 1):
        c = ws.cell(row=linha_atual, column=col, value=valor)
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        c.border = borda
        if col == 3:  # Razão Social
            c.font = Font(name="Arial", size=8.5, bold=True)
            c.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)
        elif col == 2:  # Horário
            c.font = Font(name="Arial", size=8.5, bold=True)
        elif col == 9:  # Oportunidades
            c.font = Font(name="Arial", size=8)
            c.fill = PatternFill("solid", fgColor="FAFBFF")
            c.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)
        else:
            c.font = Font(name="Arial", size=8.5)
        if col == 1:  # Unidade
            c.fill = PatternFill("solid", fgColor=UNIDADE_CORES.get(item["unidade"], AZUL_ESCURO))
            c.font = Font(name="Arial", size=8.5, bold=True, color=BRANCO)
        elif col != 9:
            c.fill = PatternFill("solid", fgColor=fundo)
    ws.row_dimensions[linha_atual].height = 38
    linha_atual += 1

# Rodapé
ws.merge_cells(start_row=linha_atual + 1, start_column=1, end_row=linha_atual + 1, end_column=n_cols)
rodape = ws.cell(row=linha_atual + 1, column=1,
                  value=f"Documento confidencial · Uso interno · Grupo Vigna © {datetime.now().year}")
rodape.font = Font(name="Arial", size=7.5, italic=True, color="8899BB")
rodape.fill = PatternFill("solid", fgColor="F0F3FA")
rodape.alignment = Alignment(horizontal="center", vertical="center")

ws.freeze_panes = "A6"

# Nome do arquivo usa a data DAS REUNIÕES (não necessariamente hoje)
nome_arquivo = f"Quadro_Visitas_Vigna_{data_str.replace('/', '')}.xlsx"
wb.save(f"relatorios/quadro-visitas/{nome_arquivo}")
```

---

## Confiabilidade

- Usar apenas informações públicas verificáveis (CNPJ, razão social, regime tributário)
- Se CNPJ ou regime não localizado: `"Não localizado"`
- Nunca inventar dados financeiros, filiais ou estrutura societária
- Se faltar algum dado da reunião (consultor, cargo, etc), deixar em branco — não inventar

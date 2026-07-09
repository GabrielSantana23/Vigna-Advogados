---
name: lancamento-reunioes-agendor
description: >
  Extrai reuniões do Agendor (CRM) do Grupo Vigna e preenche a aba "Lançamento_Reuniões"
  do arquivo Controle_Comercial_Vigna - [mês].xlsx sem precisar de digitação manual.
  Use quando o usuário pedir para "importar reuniões", "lançar as reuniões no controle",
  "puxar as reuniões do Agendor", "preencher o controle comercial com as reuniões", ou
  colar uma lista de empresas com status ("agendada"/"realizada") pra registrar no
  controle mensal. Também ativa quando o usuário mencionar "quadro de reuniões do mês"
  ou pedir pra cruzar uma lista de empresas com os dados do Agendor. Sempre confirme o
  escopo de SDR, o período e a lista de empresas antes de gravar — nunca sobrescreva a
  planilha sem antes checar divergências com o usuário.
---

# Lançamento de Reuniões do Agendor (Grupo Vigna)

Automatiza o trabalho manual de: pegar as reuniões que os SDRs marcaram/realizaram no
Agendor, enriquecer cada uma com os dados da empresa (CNPJ, segmento, cidade/estado,
contato, cargo, consultor Vigna que participou) e gravar tudo na aba `Lançamento_Reuniões`
do arquivo `Controle_Comercial_Vigna - [mês].xlsx` (normalmente em `relatorios/`).

## Quando usar

- O usuário pede pra "importar", "lançar" ou "puxar" as reuniões do Agendor no controle comercial.
- O usuário cola uma lista de empresas (às vezes com status tipo "agendada"/"realizada" do lado)
  pra virar linha na planilha.
- A aba `Lançamento_Reuniões` já tem nomes de empresa colados brutos na coluna B (às vezes
  literalmente "NOME DA EMPRESA - agendada") esperando enriquecimento.

## Colunas da aba Lançamento_Reuniões

| Coluna | Campo | Origem |
|---|---|---|
| A | Data da Reunião | `dueDate` da tarefa (converter UTC → horário de Brasília, cuidado com virada de dia perto da meia-noite) |
| B | Empresa | Nome oficial da organização no Agendor |
| C | Segmento | `organization.sector.name` |
| D | CNPJ | `organization.cnpj`, formatado `00.000.000/0001-00` |
| E | Nome do Responsável | Extraído do texto livre da tarefa (campo "Reunião: com Sr(a). NOME...") |
| F | Cargo | Extraído do texto livre (ex: "responsável pelo RH", "responsável jurídico") |
| G | Cidade | `organization.address.city` |
| H | Estado | `organization.address.state` |
| I | SDR | `task.user.name` (nome oficial com filial, ex: "Kailany Santos da Silva - Matriz SP") |
| J | Consultor/Especialista | Campo "Participação:" do texto livre da tarefa |
| K | Status | Realizada / Cancelada / Remanejada / Futura — **só preencher se o usuário pedir explicitamente** (ver seção "Status" abaixo) |
| L | Teve Proposta | Sim/Não — não vem do Agendor, não preencher salvo pedido explícito |
| M | Tipo de Proposta | Não vem do Agendor, não preencher salvo pedido explícito |
| N | Negócio Fechado | Não vem do Agendor, não preencher salvo pedido explícito |

**Importante (aprendido em sessão real com o Gabriel):** por padrão, ele prefere deixar
K/L/M/N em branco e só usar essa skill pra popular os dados cadastrais da empresa e da
reunião (A a J). Proposta e negócio fechado ele controla por outro processo. Só preencha
essas colunas se o usuário pedir na hora.

## Pré-requisitos que você precisa confirmar antes de rodar

1. **Escopo de SDR.** O time tem Igor Souza, Kailainy, Ana Karolayne e Michelle, e o
   próprio Gabriel também aparece como responsável por reuniões no Agendor. Pergunte
   quais SDRs entram nessa leva — não assuma o time todo nem assuma só quem foi citado
   se a fala for ambígua (ditado por voz costuma repetir ou trocar nomes).
2. **Período.** As reuniões "agendadas" podem ter `dueDate` bem à frente do dia de hoje
   (reunião marcada com 1-2 semanas de antecedência é normal) — não limite a busca ao
   intervalo "até hoje" achando que isso é o período da reunião. O período é sobre
   quando a reunião foi marcada/está no radar do mês, não sobre a data em si estar no
   passado.
3. **Fonte da lista de empresas.** Se o usuário colar uma lista de nomes, usar essa lista
   como escopo de busca. Se pedir pra puxar tudo do Agendor pro período, buscar direto
   pelas tarefas tipo "Reunião" de cada SDR (ver `buscar-agendor.js` da skill
   `quadro-visitas` como referência de paginação por `dueDate`).
4. **Caminho do arquivo.** Normalmente `relatorios/Controle_Comercial_Vigna - [mês].xlsx`.
   Confirme o nome exato com o usuário (varia por mês).

## Passo a passo

### 1. Buscar organização por nome

Usar o parâmetro `name` na busca — **não `q`**, que não filtra corretamente e pode
devolver uma organização qualquer sem relação com o termo buscado:

```
GET https://api.agendor.com.br/v3/organizations?name=NOME DA EMPRESA
```

Se não achar de primeira, tentar de novo só com a primeira ou duas palavras do nome
(nomes muito longos com sufixo LTDA/S.A costumam falhar na busca exata). Se ainda assim
não achar, marcar como "não localizado" e perguntar ao usuário o nome exato — **nunca
inventar CNPJ, segmento ou qualquer dado de uma empresa não encontrada**.

### 2. Buscar as tarefas de reunião da organização

```
GET https://api.agendor.com.br/v3/organizations/{orgId}/tasks?dueDateGt=AAAA-MM-DD
```

**Atenção:** a API do Agendor não deixa `dueDateGt` ficar mais de 31 dias no passado em
relação a hoje — erro `"must be less than or equal to 31 days"`. Usar uma data dentro
desse limite (ex: hoje menos 30 dias) e filtrar client-side por `type === "Reunião"`.

Rodar com:
```bash
node --env-file=../../../../.env scripts/buscar-reunioes.js empresas.json saida.json
```
onde `empresas.json` é um array simples dos nomes de empresa a buscar.

### 3. Resolver múltiplas tarefas pra mesma empresa

Uma organização pode ter mais de uma tarefa tipo "Reunião" na janela buscada. Antes de
escolher qual usar:
- Se duas tarefas têm a mesma data/hora e texto quase idêntico → provável duplicata no
  Agendor. Usar a mais recente (maior `taskId`) e ignorar a outra.
- Se as tarefas são de SDRs diferentes → checar se ambas fazem sentido pro escopo pedido
  pelo usuário (uma pode pertencer a outro SDR fora do time, de outra filial).
- Se há uma tarefa antiga já finalizada e outra nova em aberto pra mesma reunião →
  normalmente é reagendamento; usar a mais recente.
- Nunca decidir sozinho quando a ambiguidade for material (duas reuniões genuinamente
  diferentes, por exemplo) — perguntar ao usuário.

### 4. Extrair contato, cargo e consultor do texto livre

O campo `text` da tarefa é texto livre digitado pelo SDR, sem padrão rígido, mas segue
um formato recorrente:
```
Reunião: [Pessoa Jurídica/Recursos Humanos/...]
Ativo: [nome do SDR]
Reunião: com [Sr./Sra./Dr./Dra.] NOME, responsável [pelo/pela] CARGO na empresa EMPRESA
Data: DD/MM/AAAA
Hora: HH:MM
Local: ...
Telefone: ...
Participação: NOME(S) DO(S) CONSULTOR(ES)
Pauta: ...
```
- **Nome do Responsável / Cargo:** extrair da linha "Reunião: com...". Se o cargo não
  estiver explícito, usar `"Não informado"` — não inventar a partir do segmento da empresa.
- **Consultor/Especialista:** vem do campo "Participação:". Copiar como está (separado
  por "e"/"/" se houver mais de um nome), sem tentar normalizar capitalização de forma
  agressiva.
- Se o nome da empresa que aparece dentro do texto livre for diferente do nome oficial
  da organização no Agendor (acontece — é erro de cadastro do time), usar o nome oficial
  da organização na coluna Empresa e não se preocupar em corrigir o texto livre.

### 5. Checar a lista de SDR da planilha (aba oculta `Listas`, intervalo `ListaSDR`)

Se o `SDR` (nome do usuário que finalizou/é dono da tarefa no Agendor) não estiver na
lista suspensa `ListaSDR` da aba `Listas`, avisar o usuário e perguntar se deve:
- adicionar o nome à lista suspensa antes de gravar, ou
- gravar mesmo assim sem mexer na lista (pode gerar aviso de validação no Excel depois,
  mas não trava a gravação).

### 6. Montar o JSON de linhas e gravar

Montar um objeto indexado por número de linha (ver formato em
`scripts/preencher-planilha.js`) e rodar:

```bash
node scripts/preencher-planilha.js "relatorios/Controle_Comercial_Vigna - mes.xlsx" linhas.json
```

O script já faz backup automático do arquivo antes de gravar (salvo ao lado do original
com sufixo `.BACKUP-<timestamp>.xlsx`) e preserva as outras abas (`Visão_Geral`,
`Lançamento_Ligações`, `Propostas_Enviadas`, `Listas`) intocadas.

**Antes de rodar**, sempre mostrar uma prévia resumida pro usuário (tabela ou lista) com
pelo menos: empresa, data, SDR e qualquer divergência/problema encontrado (empresa não
localizada, SDR fora da lista suspensa, tarefa duplicada, SDR fora do escopo pedido). Só
gravar depois da confirmação — é um arquivo de controle comercial oficial, erro ali tem
custo real.

### 7. Reportar pro usuário

Ao final, resumir: quantas reuniões foram gravadas, quantas empresas não foram
localizadas no Agendor (e pedir o nome certo se fizer sentido), e qualquer inconsistência
de dado encontrada no Agendor em si (nome de empresa divergente dentro do texto livre,
tarefas duplicadas, etc.) — sem tentar "consertar" o cadastro do Agendor por conta própria.

## Limitações importantes (sempre comunique isso ao usuário)

- A extração de Nome do Responsável, Cargo e Consultor depende de texto livre digitado
  pelo SDR — não é um campo estruturado do Agendor. Trate como primeira leitura.
- Segmento vem do campo `sector` da organização no Agendor, que reflete a classificação
  que já está cadastrada lá (às vezes genérica ou aparentemente errada, ex: "órgãos
  públicos" pra uma transportadora privada) — não é a skill que decide isso, é o
  cadastro do CRM.
- CNPJ, cidade e estado só vêm preenchidos se o cadastro da organização no Agendor já
  tiver esses dados. Quando faltar, usar `"Não localizado"` — nunca pesquisar/inventar.
- Essa skill cobre apenas tarefas do tipo `"Reunião"`. Ligação tem skill própria
  (`tabulacao-ligacoes-agendor`).

## Referência rápida da API do Agendor

- **Token:** `AGENDOR_API_TOKEN` no `.env` da raiz do projeto. Nunca expor esse valor em
  outputs.
- **Buscar organização por nome:** `GET /v3/organizations?name=TERMO` (usar `name`, não `q`)
- **Tarefas de uma organização:** `GET /v3/organizations/{id}/tasks?dueDateGt=AAAA-MM-DD`
  (janela máxima de 31 dias a partir de hoje)
- **Tarefas do dia (todas as organizações):** `GET /v3/tasks?dueDateGt=...&dueDateLt=...`
  (ver `scripts/buscar-agendor.js` da skill `quadro-visitas` para paginação)

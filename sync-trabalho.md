# Sincronização — Computador do Trabalho

Checklist pra deixar este workspace funcionando no PC do trabalho, com outra conta Claude.

## O que já vem pronto ao clonar o repositório

- Skills (`.claude/skills/`) e comandos (`.claude/commands/`)
- Contexto do negócio (`_contexto/`), guia de marca (`marca/`) e templates (`templates/`)
- `CLAUDE.md` com as instruções do sistema
- Hook de auto-sync (salva no GitHub automaticamente ao final de cada sessão)

## Passo a passo no PC do trabalho

1. **Clonar o repositório** (mesma conta GitHub de casa, já tem acesso):
   ```
   git clone https://github.com/GabrielSantana23/Vigna-Advogados.git
   ```

2. **Abrir a pasta no VSCode** com a extensão do Claude Code instalada.

3. **Fazer login no Claude Code** com a conta do trabalho.

4. **Reconectar os MCPs na nova conta** (em claude.ai → Configurações → Connectors):
   - Canva
   - Notion
   - Gamma
   - Microsoft 365

   Esses ficam vinculados à conta Claude (não ao projeto), por isso precisam ser reconectados na conta nova.

5. **Recriar o arquivo `.env`** na raiz do projeto com:
   ```
   AGENDOR_API_TOKEN=<colar o valor copiado de casa>
   ```
   Esse arquivo nunca vai pro GitHub por segurança (token usado pela skill `quadro-visitas`).

6. **Instalar as dependências da skill `quadro-visitas`**:
   ```
   cd .claude/skills/quadro-visitas/scripts
   npm install
   ```

7. **Testar**: roda `/iniciar` pra confirmar que o contexto carregou certo.

## Observações

- A memória de longo prazo do Claude (aprendizados de sessões anteriores) fica fora do repositório, vinculada à máquina/conta de casa. No trabalho ela começa do zero e vai se reconstruindo com o uso.
- Depois desse setup inicial, da próxima vez que abrir no trabalho já está tudo pronto — só `git pull` pra puxar mudanças feitas aqui.
- Pode apagar este arquivo depois que o setup estiver concluído.

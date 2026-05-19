Usar com a skill /implementation-plan

<tarefa>
Implementar uma feature significativa no aplicativo Gerencie Carteira: substituir o
fluxo atual de erros `#N/D` na coluna de gerente por um fluxo de captura de gerentes
órfãos em runtime, remover a abertura do DIRECIONA, e adicionar confirmação de
fechamento com rollback transacional do `.xlsm`.
</tarefa>

<contexto>
O Gerencie Carteira é uma automação Windows-only: extrai dados de monitoramento do
Serasa via e-mails do Outlook (COM), parseia HTML e atualiza planilhas Excel `.xlsm`
macro-habilitadas via xlwings. Backend Python v3 com pipeline JSON Lines em stdout,
consumido por shell Electron 33 + React 18 + Vite 5 + TypeScript. Versão atual v3.0.1.

**Comportamento atual (a ser substituído):** ao colar os dados na planilha, CNPJs
sem cadastro de gerente geram `#N/D` na coluna de verificação. O programa avisa o
usuário e abre o executável DIRECIONA para o usuário preencher manualmente na
planilha depois.

**Comportamento desejado:** o programa resolve os gerentes órfãos ANTES de colar,
através de um formulário em runtime, e reinjeta os dados na sheet de PROCX. O usuário
não precisa mais abrir a planilha — exceto para olhar a tabela dinâmica.
</contexto>

<arquivos_criticos>
| Arquivo | Papel |
|---------|-------|
| `backend/src/gerencie_carteira.py` | Entrypoint do pipeline; `main()` define a invariante de ordem |
| `backend/src/log_emitter.py` | Contrato JSON Lines (`emit`, `emit_result`) — única saída sancionada Python→UI |
| `backend/src/directory_bootstrap.py` | Workspace idempotente, verificação de share |
| `app/electron/main.ts` | spawn do core, readline IPC, dialog da cascata (exit 4), whitelist `shell.openPath`. Contém `BASE_FILENAME_PATTERN` |
| `app/electron/preload.ts` | contextBridge — superfície `window.electronAPI` |
| `app/src/hooks/useScriptRunner.ts` | Reducer de estado; detecção de exit codes e auto-rerun |
| `app/src/App.tsx` | Composição da UI |
| `config/config.ini` | Paths via `%USERPROFILE%`, aba/coluna do Excel, assunto do e-mail |
| `app/package.json` | Fonte única da versão (SemVer) |
</arquivos_critos>

<fases>

<fase numero="1" nome="Leitura da sheet PROCX GERENTES">
- Objetivo: carregar o mapeamento CNPJ → Gerente existente.
- Ler a sheet `PROCX GERENTES`: gerentes na coluna B, CNPJs na coluna C.
- Montar uma estrutura em memória (array/dict) CNPJ → gerente.
- Entregável: estrutura carregada e logada via `emit()` com contagem de entradas.
</fase>

<fase numero="2" nome="Detecção de CNPJs órfãos">
- Objetivo: identificar quais CNPJs dos e-mails não têm gerente.
- Montar o array de CNPJs a serem colados, advindos dos e-mails obtidos.
- Cruzar com o mapeamento da Fase 1; um CNPJ é "órfão" se não existir no PROCX.
- Entregável: lista de CNPJs órfãos. Se vazia, pular a Fase 3.
</fase>

<fase numero="3" nome="Formulário de input em runtime">
- Objetivo: capturar o gerente de cada CNPJ órfão com o usuário.
- O Python emite um evento JSON Lines solicitando input (novo `step`/evento no
  protocolo); o Electron/React abre uma interface listando cada CNPJ órfão com um
  campo para o gerente.
- O botão de confirmar só fica habilitado quando TODOS os campos estão preenchidos.
- Ao confirmar, a UI devolve o mapeamento ao Python via IPC.
- Entregável: mapeamento CNPJ → gerente completo para todos os órfãos.
</fase>

<fase numero="4" nome="Reinjeção no PROCX GERENTES">
- Objetivo: persistir os novos gerentes na planilha.
- Acrescentar as novas linhas ao array (Fase 1 + respostas da Fase 3).
- Injetar de volta na sheet `PROCX GERENTES`: gerentes na coluna B, CNPJs na coluna C.
- Preencher dinamicamente as demais colunas (replicar a fórmula/preenchimento da
  linha anterior, como já é feito para o VLOOKUP da coluna de verificação).
- Entregável: sheet `PROCX GERENTES` atualizada sem `#N/D`.
</fase>

<fase numero="5" nome="Colagem em Email-BD">
- Objetivo: colar os dados dos e-mails normalmente.
- Inserir as linhas em `Email-BD` exatamente como é feito atualmente.
- Como os gerentes já foram resolvidos na Fase 4, não deve restar nenhum `#N/D`.
- Entregável: `Email-BD` atualizada, sem pendências de gerente.
</fase>

<fase numero="6" nome="Remoção do DIRECIONA">
- Objetivo: eliminar a abertura do executável DIRECIONA.
- Remover a lógica que dispara `executavel_direciona` no caso `#N/D`/pendências.
- Remover config/whitelist relacionada se ficar órfã (`config.ini`, whitelist do
  `shell.openPath` em `main.ts`).
- Entregável: nenhum caminho de código abre o DIRECIONA.
</fase>

<fase numero="7" nome="Confirmação de fechamento + rollback transacional">
- Objetivo: impedir perda silenciosa de estado se o usuário fechar durante o runtime.
- Se o usuário tentar fechar o programa em execução, abrir janela perguntando
  "SIM"/"NÃO" — "o programa está em execução, deseja realmente fechar?".
- Se confirmar o fechamento: descartar tudo; o estado deve ficar como se o programa
  não tivesse rodado.
- Mecanismo transacional do `.xlsm`: criar um backup da planilha antiga ANTES de
  qualquer escrita. Só deletar o backup APÓS `wb.save()` bem-sucedido. Se o
  fechamento/erro ocorrer durante a colagem no novo `.xlsm`, descartar o novo e
  restaurar a partir do backup. Nunca delete a planilha antiga antes de ter
  produzido a nova com sucesso.
- Entregável: fechar no meio do processo nunca corrompe nem perde a planilha.
</fase>

</fases>

<restricoes_inviolaveis>
1. **Protocolo JSON Lines:** todo output Python→UI passa por `emit()`/`emit_result()`
   em `log_emitter.py`. NENHUM `print()` solto — corromperia o stream.
2. **Invariante de ordem:** em `main()`, a base DEVE ser resolvida ANTES de
   `extrair_dados_dos_anexos()` (que marca e-mails `UnRead=False`). NÃO reordenar.
   A captura de gerentes órfãos não pode quebrar essa ordem.
3. **Preservação de macros:** arquivos permanecem `.xlsm`; manipulação via xlwings.
   Forçar `app.calculate()` antes de ler colunas calculadas (já feito na v3.0.1).
4. **Rollback transacional:** backup criado antes de escrever, removido só após save
   bem-sucedido. A planilha antiga só é deletada após a nova existir com sucesso.
5. **Não agravar dívida conhecida:** e-mails são marcados como lidos antes de
   `wb.save()`. Não introduzir nova lógica que amplie essa janela de perda; idealmente
   o novo fluxo de input acontece ANTES da marcação de lidos.
6. **SemVer:** esta é uma mudança funcional MINOR → bump v3.0.1 para v3.1.0 em
   `app/package.json` (fonte única). Não hardcodar versão em outro lugar.
7. **Idioma:** todo código, comentário, log e UI em Português Brasileiro.
8. **Sincronia de pattern:** se mexer em `BASE_FILENAME_PATTERN`, manter `main.ts` e
   o core Python em sincronia.
</restricoes_inviolaveis>

<criterios_de_aceitacao>
1. Sheet `PROCX GERENTES` é lida e o mapeamento CNPJ→gerente é montado em memória.
2. CNPJs órfãos são detectados antes de qualquer escrita na planilha.
3. Formulário de input aparece SOMENTE se houver órfãos; confirmar só habilitado com
   todos os campos preenchidos.
4. Após confirmação, `PROCX GERENTES` é reinjetada (B=gerente, C=CNPJ) com
   preenchimento dinâmico das demais colunas; `Email-BD` colada sem `#N/D`.
5. DIRECIONA nunca é aberto em nenhum caminho de código.
6. Fechar durante o runtime exibe diálogo SIM/NÃO e, se confirmado, restaura o estado
   anterior (backup) sem corromper a planilha antiga.
7. Testes pytest passando (incluir novos testes para detecção de órfãos e rollback).
8. Versão bumpada para v3.1.0 e docs de contexto atualizados conforme `AGENTS.md`.
</criterios_de_aceitacao>

<edge_cases>
- Sheet `PROCX GERENTES` vazia ou inexistente → todos os CNPJs são órfãos; tratar sem crash.
- Nenhum CNPJ órfão → não exibir formulário; seguir direto para colagem.
- Backup já existente de execução anterior abortada → sobrescrever/limpar com segurança.
- Falha ao criar o backup → abortar antes de qualquer escrita (não prosseguir sem rede de segurança).
- CNPJ duplicado entre e-mails → resolver gerente uma única vez.
- `wb.save()` falha após input do usuário → restaurar backup, emitir erro, não perder os dados digitados (re-emitir se possível).
</edge_cases>

<verificacao>
- [ ] `PROCX GERENTES` lida (B=gerente, C=CNPJ)
- [ ] Array de CNPJs dos e-mails montado antes de colar
- [ ] Órfãos detectados por cruzamento
- [ ] Formulário só aparece com órfãos
- [ ] Confirmar só habilitado com todos os campos preenchidos
- [ ] Novo evento JSON Lines documentado e tratado em `useScriptRunner.ts`
- [ ] Reinjeção em `PROCX GERENTES` com preenchimento dinâmico das colunas extras
- [ ] `Email-BD` colada normalmente, sem `#N/D`
- [ ] DIRECIONA removido (código + config + whitelist órfã)
- [ ] Diálogo SIM/NÃO ao fechar em runtime
- [ ] Backup criado antes de escrever
- [ ] Backup deletado SOMENTE após `wb.save()` bem-sucedido
- [ ] Planilha antiga deletada só após nova criada com sucesso
- [ ] Invariante de ordem de `main()` preservada
- [ ] Versão bumpada para v3.1.0 em `app/package.json`
- [ ] pytest passing
</verificacao>

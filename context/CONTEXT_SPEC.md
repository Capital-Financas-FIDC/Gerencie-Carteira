# CONTEXT_SPEC — Especificacao Canonica para Documentos de Contexto

> **IMUTAVEL.** Este documento define a estrutura canonica dos documentos de contexto do projeto.
> Qualquer alteracao em `metaspec.md`, `index.md` ou `timeline.md` DEVE seguir este spec.
> Para alterar ESTE documento, e necessaria aprovacao explicita do usuario.
>
> Versao: 2.0 | Criado: 2026-05-18

---

## Principios Fundamentais

Baseado nas diretrizes oficiais da Anthropic para CLAUDE.md, nas melhores praticas de
AI-Driven Development da industria (Google, Meta, Stripe), e nos padroes de documentacao
token-eficiente para LLMs.

### 1. Token Economy

Cada linha consome tokens do contexto. Linhas que nao mudam o comportamento do agente
sao custo sem retorno. Regra de ouro: **se remover a linha nao causa erro, ela nao deveria existir.**

### 2. Single Ownership

Cada informacao existe em EXATAMENTE UM documento. Redundancia entre arquivos cria
divergencia inevitavel e desperdica tokens.

### 3. Derivabilidade

Informacao que o agente pode obter lendo o codigo (via Glob, Grep, Read) NAO deve ser
documentada. Documentar apenas o que NAO e inferivel: decisoes, motivacoes, regras de
negocio, e estado que requer contexto humano.

### 4. Anti-Frankenstein

Documentos de contexto NAO sao logs. Eles representam o ESTADO ATUAL do projeto.
Atualizacoes SUBSTITUEM informacao — nunca acumulam. Historico pertence ao git e
ao `timeline.md`.

---

## Mapa de Ownership

Cada tipo de informacao tem um unico dono. Violar ownership cria redundancia.

| Informacao | Dono | NAO colocar em |
|------------|------|----------------|
| Identidade, dominio, proposito | `metaspec.md` | index, timeline |
| Stack e versoes | `metaspec.md` | index, CLAUDE.md |
| Arquitetura e fluxo de dados | `metaspec.md` | docs/ |
| Estado atual (versao, branch, saude) | `metaspec.md` | index, timeline |
| Dividas tecnicas abertas | `metaspec.md` | timeline |
| Autenticacao e autorizacao | `metaspec.md` | index, timeline |
| Fontes de dados e integracoes | `metaspec.md` | index, timeline |
| Regras de negocio criticas | `metaspec.md` | timeline, index |
| Lista de artefatos ativos | `index.md` | metaspec |
| Navegacao entre documentos | `index.md` | metaspec, timeline |
| Inventario de arquivos criticos | `index.md` | metaspec |
| Resumo de testes (status) | `index.md` | metaspec |
| Infraestrutura e CI/CD | `index.md` | metaspec |
| Documentacao tecnica (referencias) | `index.md` | metaspec |
| Historia evolutiva por fases | `timeline.md` | metaspec, index |
| Metricas acumuladas historicas | `timeline.md` | metaspec, index |
| Comandos de dev (build, test, lint) | `CLAUDE.md` | metaspec, index |
| Convencoes de codigo/workflow | `CLAUDE.md` | metaspec |

---

## Formato: metaspec.md

**Proposito:** Contexto essencial do projeto para o agente. Responde: "O que e este projeto,
como funciona, e o que eu preciso saber para nao quebrar nada?"

**Limite:** 200 linhas. Se ultrapassar, algo esta redundante ou derivavel.

### Secoes Obrigatorias (nesta ordem)

```
# MetaSpec — {NomeProjeto}
> Contexto para agentes AI. Versao: X.Y | Atualizado: YYYY-MM-DD

## IDENTIDADE
  - Nome, dominio, proposito, usuarios-alvo, idioma
  - Maximo: 5 bullets

## STACK
  - Bloco de codigo com tecnologias e versoes
  - Uma linha por camada (ex: runtime, framework, banco, infra)
  - NAO listar LOC — isso muda a cada commit

## ARQUITETURA
  - Diagramas ASCII de fluxo de dados (maximo 2)
  - Tabelas de camadas: Camada | Diretorio | Responsabilidade
  - NAO listar arquivos individuais (isso e do index.md)
  - NAO listar LOC

## ESTADO ATUAL (vX.Y — DD/MM/YYYY)
  - Branch atual e saude geral (1 linha)
  - "Pronto" — lista concisa do que funciona (bullets)
  - "Dividas tecnicas" — lista com acao necessaria
  - NAO incluir historico de versoes anteriores (isso e do timeline)
  - NAO incluir "o que foi removido em vX" (isso e do timeline)
  - NAO incluir "o que mudou na vX.Y" (isso e do timeline)
```

### Secoes Condicionais (incluir apenas se aplicavel ao projeto)

```
## DADOS
  - Incluir se: o projeto integra com fontes externas ou armazena dados
  - Fontes externas (APIs, bancos, filas, streams)
  - Armazenamento local (caches, TTLs, file storage)
  - Configs que afetam comportamento

## AUTH
  - Incluir se: o projeto tem autenticacao/autorizacao
  - Tabela unica: Aspecto | Detalhe
  - Provider, algoritmo, roles, fluxo de acesso

## REGRAS DE NEGOCIO CRITICAS
  - Incluir se: o dominio tem regras que o agente PRECISA saber para nao quebrar logica
  - Formato: subsecao por dominio
  - Bullets concisos, sem explicacao tutorial
  - APENAS regras nao-obvias que afetam implementacao
```

### Secoes PROIBIDAS no metaspec

- `REFERENCIAS` — O index ja faz isso
- `CONVENCOES` — O CLAUDE.md ja faz isso
- Historico de versoes ("Removido em v11", "Mudou em v12.1")
- LOC de arquivos individuais
- Cross-references para outros docs de contexto

### Regra de Atualizacao

Ao atualizar `metaspec.md`, o agente DEVE:
1. **Substituir** a secao relevante com o estado novo (nao adicionar ao existente)
2. Manter o limite de 200 linhas
3. Atualizar a versao e data no header
4. Se uma informacao migrou para outro doc, REMOVER do metaspec

---

## Formato: index.md

**Proposito:** Mapa de navegacao. Responde: "Onde encontro X?"

**Limite:** 150 linhas. Se ultrapassar, ha detalhes demais.

### Secoes Obrigatorias (nesta ordem)

```
# Context Index — {NomeProjeto}
> Mapa de artefatos. Atualizado: YYYY-MM-DD

## Navegacao Rapida
  - Tabela 4-5 linhas: Artefato | Caminho | Descricao
  - Apenas documentos de contexto core

## Artefatos Ativos
  ### Analises — context/analysis/
  ### Walkthroughs — context/walkthroughs/
  ### Plans — plans/
  - Tabela: Arquivo | Data | Descricao | Status
  - APENAS arquivos na raiz (nao old/)
  - Plans implementados devem ser movidos para old/ — nao listados aqui

## Arquivos Criticos
  - Agrupar por camada do projeto (categorias livres, adaptaveis)
  - Exemplos de categorias: "Core", "API", "UI", "CLI", "Config", "Workers"
  - Tabela: Arquivo | Responsabilidade
  - APENAS arquivos que o agente precisa conhecer para nao errar
  - NAO listar LOC (muda constantemente)
  - NAO listar TODOS os arquivos (o agente sabe usar Glob)
```

### Secoes Condicionais (incluir apenas se aplicavel ao projeto)

```
## Testes
  - Incluir se: o projeto tem test suites
  - Tabela resumo: Camada | Diretorio | Status
  - NAO listar contagem exata de assertions
  - Usar "passing" / "N failing" como status

## Infraestrutura
  - Incluir se: o projeto tem CI/CD ou deploy configs
  - Tabela: Arquivo | Descricao
  - Apenas configs de CI, scripts de deploy, .env locations

## Documentacao Tecnica
  - Incluir se: existem docs em docs/
  - Tabela: Arquivo | Descricao
  - Referencias para documentacao tecnica detalhada do projeto
```

### Secoes PROIBIDAS no index

- Contagens exatas de LOC, assertions, ou arquivos
- Notas historicas ("Removido em vX", "Era Y, agora e Z")
- Subsecoes `old/` com contagens de artefatos historicos
- Duplicacao de informacao do metaspec (stack, config details)

### Regra de Atualizacao

Ao atualizar `index.md`, o agente DEVE:
1. **Adicionar** novos artefatos criados
2. **Remover** artefatos movidos para `old/` ou deletados
3. NUNCA incrementar contadores ("agora sao 42 testes") — usar status qualitativo
4. Atualizar a data no header

---

## Formato: timeline.md

**Proposito:** Historia evolutiva do projeto. Responde: "Como chegamos aqui e por que?"

**Limite:** 250 linhas. Fases antigas devem ser comprimidas quando o documento crescer.

### Secoes Obrigatorias (nesta ordem)

```
# Timeline — {NomeProjeto}
> Historia evolutiva. {N} fases | {periodo}.

## Fase N: {Nome} ({periodo})
  - Commits: hash_inicio -> hash_fim
  - 3-5 bullets descrevendo O QUE mudou e POR QUE
  - NAO incluir "Marco:" — e redundante com os bullets
  - NAO listar walkthroughs individuais — consultar old/ se necessario

## Metricas Snapshot ({data})
  - Tabela resumo com metricas APROXIMADAS (usar "~")
  - NAO prometer precisao que decai (contagens exatas viram mentira)
```

### Compressao de Fases Antigas

Quando `timeline.md` ultrapassar 250 linhas:
1. Comprimir fases antigas (>3 meses) para formato de 2 linhas:
   ```
   ## Fase 0: Fundacao (Dez/2025) — Setup inicial + API funcional
   ## Fase 1: Refatoracao (Jan/2026) — Arquitetura em camadas
   ```
2. Manter fases recentes (<3 meses) no formato completo
3. Detalhes de fases comprimidas ficam acessiveis via `git log`

### Secoes PROIBIDAS no timeline

- Listas de walkthroughs por fase (existem no `old/`)
- "Marco:" como subsecao (redundante)
- LOC antes/depois de refatoracoes
- Contagens exatas de commits por fase (usar "~N")

### Regra de Atualizacao

Ao atualizar `timeline.md`, o agente DEVE:
1. **Adicionar** nova fase OU atualizar fase em andamento
2. NUNCA editar fases completas (>1 mes) exceto para comprimir
3. Atualizar metricas snapshot com valores aproximados
4. Se >250 linhas, comprimir fases mais antigas

---

## Regras Anti-Frankenstein

Estas regras previnem acumulo de cruft nos documentos de contexto.

### 1. Atualizacao = Substituicao

Quando o estado do projeto muda, a secao correspondente e REESCRITA com o estado novo.
Nunca adicionar "Atualizado em DD/MM: agora X e Y" — isso e log, nao estado.

**Errado:**
```
## ESTADO ATUAL (v2.0)
- API REST funcional
- Auth JWT

### O que mudou em v2.1
- Adicionado rate limiting
- Corrigido bug no refresh token
```

**Correto:**
```
## ESTADO ATUAL (v2.1 — 30/03/2026)
- API REST funcional com rate limiting
- Auth JWT com refresh token funcional
```

### 2. Uma Informacao, Um Lugar

Se a contagem de testes aparece no `metaspec.md` E no `index.md`, qual e o correto quando
divergem? Resposta: nenhum — porque a contagem nao deveria estar em nenhum. O agente
sabe contar sozinho.

### 3. Numeros Exatos Decaem

Contagens exatas (661 assertions, 484 LOC, 418 commits) ficam desatualizadas apos o
proximo commit. Usar:
- "~160 tests" ao inves de "162"
- "~500 assertions" ao inves de "499"
- Ou omitir numeros e usar status: "tests passing", "2 tests failing"

### 4. Historico Pertence ao Git

"O que foi removido em v2" nao e estado atual — e historia. Se precisa ser registrado,
vai no `timeline.md` dentro da fase correspondente. Se ja esta no timeline, NAO duplicar
no metaspec.

### 5. Prune Test

Antes de salvar qualquer atualizacao, para cada linha adicionada perguntar:
> "Se eu remover esta linha, o agente vai cometer um erro?"

Se a resposta e "nao", a linha nao deve existir.

---

## Checklist de Validacao

Usar este checklist ao atualizar qualquer documento de contexto:

- [ ] Documento esta dentro do limite de linhas?
- [ ] Nenhuma informacao esta duplicada em outro doc de contexto?
- [ ] Nenhum numero exato que vai decair no proximo commit?
- [ ] Nenhuma secao de historico/changelog no metaspec?
- [ ] Nenhuma secao "O que mudou em vX.Y" (exceto no timeline)?
- [ ] Ownership respeitado conforme tabela acima?
- [ ] Informacao derivavel do codigo foi omitida?
- [ ] Data de atualizacao atualizada no header?

---

## Referencias

Este spec foi construido a partir de:

- **Anthropic CLAUDE.md Best Practices**: target <200 linhas, prune test, conditional loading
- **Anthropic Prompting Guide**: XML tags, declarativo > descritivo, primacy effect
- **Google AI Documentation Standards**: structured context, single source of truth
- **Stripe Engineering Practices**: minimal viable documentation, ownership model
- **Meta AI-Assisted Development**: token-efficient context, derivability principle
- **Community Consensus** (Cursor Rules, Windsurf Rules, Cline Rules): short, specific, verifiable
- **Anti-patterns documentados**: Frankenstein docs, kitchen sink sessions, governance theater
- **Principio Anthropic**: "If your CLAUDE.md is too long, Claude ignores half of it"
- **Token-efficient documentation**: maximizar signal-to-noise ratio em contexto limitado

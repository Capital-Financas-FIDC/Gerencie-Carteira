# Funcionalidades:

Este programa serve para automatização da obtenção de dados vindos do Seresa, pelo programa do Gerencie Carteira.
 
# 🗂️ Guia de utilização

Aqui temos um guia sobre como algumas configurações importantes e um guia de versionamento, sobre como o programa foi versionado até agora, e como ele deve ser versionado para versões futuras.

## ⚙️ Arquivo Config 

Este arquivo gerencia os diretórios e caminhos que o programa vai utilizar. Disponível a partir da `versão 2.8.0`, lá está indicado onde deve ser colocado cada diretório. 

Aqui no GitHub temos um arquivo chamado `config.example.init`. Nele temos um **exemplo**. Ele contém indicações do que ser feito, e **DEVE** ser alterado antes do uso, colocando dentro dele os diretórios que o usuário vai usar. O arquivo também deve ser renomeado para `config.ini`.

---

## 🗂️ Guia de Versionamento

Esta parte define a convenção de versionamento usada neste repositório para organizar as versões principais, branches alternativas e builds experimentais do programa.

---

### 📁 Estrutura Geral

Cada versão teve um commit próprio, gerando um histórico. Ele pode ser acessado no histórico de commits do repositório.

---

### 🧩 Versões Principais (`main/`)

As versões principais seguem um esquema numérico no formato `MAJOR.MINOR.PATCH`, **sem compromisso com o padrão SemVer**, mas com algumas ideias semelhantes:

- `MAJOR`: Mudanças grandes, reestruturações ou mudanças de compatibilidade.
- `MINOR`: Novas funcionalidades que mantêm compatibilidade com a versão anterior.
- `PATCH`: Correções de bugs ou melhorias pequenas.

**Exemplos:**
- `2.5.0`: Nova funcionalidade adicionada.
- `2.5.1`: Correção ou pequeno ajuste na versão anterior.
- `2.6.0`: Mudanças significativas, mas ainda compatível com a linha 2.x.

---

### 🌿 Branches Alternativas

Branches são variações paralelas de versões principais, como implementações alternativas, mudanças de interface, ou builds experimentais.

#### 🔖 Formato recomendado:

`MAJOR.MINOR.PATCH+nome-da-branch`


#### 📌 Exemplos:
- `2.5.1+streamlit`: Versão com interface feita em Streamlit, baseada na 2.5.1.
- `2.5.1+tkinter`: Versão com interface Tkinter, também baseada na 2.5.1.

Essas versões não substituem a principal e não devem ter numeração superior à linha base (evitar `2.6.0`, `2.7.0` se forem apenas variações paralelas da 2.5.1).

---

### ⚙️ Builds em descontinuadas e em desenvolvimento

Aqui trataremos de versões do programa que não funcionam ou estão em desenvolvimento como protótipo

#### 📌 Exemplos:
- `2.4.0+broken`: Versão com bugs críticos, mantida apenas como histórico.
- `3.0.0-dev`: Protótipo de uma futura linha 3.x.

Essas versões **não devem ser tratadas como releases** e não devem ser confundidas com versões principais.


---

### ✅ Regras Gerais de Versionamento

- Use `main/` para versões oficiais e estáveis.
- Use uma branch para projetos alternativos que partem de alguma versão da main.
- Use `dev/` para versões de teste, protótipos ou versões quebradas.
- Utilize `+sufixo` para identificar branches ou builds especiais.
- Não crie versões com números maiores se elas não forem de fato evoluções da linha principal.
- Evite nomes genéricos como `final`, `nova`, `teste`, `última`.

---

### 📁 Exemplo de Estrutura Final

├── main/
│ ├── 2.5.0/
│ ├── 2.5.1/
│ └── 2.6.0/
├── streamlit
│ ├── 2.5.1+streamlit/
├── tkinter
│ └── 2.5.1+tkinter/
└── dev/
  └── 2.4.0+broken/


---

### 🧪 Alterações futuras

Qualquer modificação nas regras de versionamento deve ser registrada neste documento.

Para dúvidas, consulte o mantenedor principal do projeto.

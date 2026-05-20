# FAQ Automation

Automação local em Python que lê a planilha de release do Jira e gera automaticamente:

- **FAQ em Markdown** (`faq_432.1.md`) — estrutura de Melhorias e Correções agrupadas por Menu
- **Nova planilha** (`release_432.1.xlsx`) — mesmo layout e formatação, dados limpos

---

## Estrutura do FAQ gerado

```
# FAQ — Versão 432.1

• Melhorias 🗒️
    • Financeiro:
        • Relatórios: Novo filtro de busca por período
        • Cobrança: Exportação de boletos em lote
    • Estoque:
        • Dashboard: Novo painel de movimentação

• Correções 🗒️
    • Cadastro:
        • Usuários: Correção na edição de permissões
```

**Regras aplicadas automaticamente:**
- Coluna `Tipo = Melhoria` → seção Melhorias
- Coluna `Tipo = Correção` → seção Correções
- Agrupamento por coluna `Menu`
- Subitem por coluna `Especifico` (se preenchido)
- Linhas vazias e tipos inválidos são ignorados
- Ordem original da planilha é preservada

---

## Colunas esperadas na planilha

| Coluna | Obrigatória | Descrição |
|---|---|---|
| **Tipo** | ✅ | `Melhoria` ou `Correção` |
| **Menu** | ✅ | Módulo/área (ex: Financeiro, Cadastro) |
| **Descrição** | ✅ | Texto que aparece na FAQ |
| Especifico | — | Submenu (ex: Relatórios, Usuários) |
| Link | — | Link do ticket no Jira |
| Merge | — | Controle interno |
| Faq | — | Controle interno |
| Versão | — | Versão atual |
| Faq de versão | — | Versão do FAQ |

---

## Instalação

```bash
# 1. Entre na pasta do projeto
cd faq-automation

# 2. Crie o ambiente virtual
python3 -m venv venv

# 3. Ative o ambiente virtual
source venv/bin/activate        # Linux / macOS
venv\Scripts\activate           # Windows

# 4. Instale as dependências
pip install -r requirements.txt
```

> **Toda vez que abrir um novo terminal**, repita o passo 3 antes de rodar o script.

---

## Como usar

### Com sua planilha real

```bash
python run.py --file exemplos/Chamados-do-Mes.xlsx --version 431.1 --next-version 432.1
```

### Parâmetros

| Parâmetro | Obrigatório | Descrição |
|---|---|---|
| `--file` | ✅ | Caminho da planilha (.xlsm ou .xlsx) |
| `--version` | ✅ | Versão atual (ex: `431.1`) |
| `--next-version` | ✅ | Próxima versão (ex: `432.1`) |
| `--output` | — | Pasta de saída (padrão: `./output`) |

---

## Estrutura do projeto

```
jira-faq-automation/
├── run.py                      ← Ponto de entrada
├── requirements.txt
├── README.md
├── modules/
│   ├── reader.py               ← Lê e valida a planilha
│   ├── faq_generator.py        ← Gera o .md com a estrutura de FAQ
│   ├── sheet_generator.py      ← Gera a nova planilha .xlsx
│   └── utils.py                ← Helpers e constantes
├── exemplos/
│   └── release_431.1.xlsx      ← Planilha com os chamados
└── output/                     ← Arquivos gerados aqui
    ├── faq_432.1.md
    └── release_432.1.xlsx
```

---

## Personalizações comuns

**Adicionar novo tipo além de Melhoria/Correção:**
Edite `modules/faq_generator.py` — adicione uma nova seção com `_renderizar_secao()`.

**Mudar o emoji das seções:**
Em `modules/faq_generator.py`, edite as chamadas `_renderizar_secao("Melhorias", "🗒️", ...)`.

**Mudar o agrupamento (ex: agrupar por Especifico ao invés de Menu):**
Edite `_agrupar_por_menu()` em `modules/faq_generator.py`.

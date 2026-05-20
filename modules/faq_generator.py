"""
Gera o FAQ no formato:

📄 Melhorias
➡ Menu
    • Específico: Descrição        ← item único
    • Específico:                  ← múltiplos itens
        ◦ Descrição 1
        ◦ Descrição 2

📄 Correções
➡ Menu
    • Específico: Descrição
"""
import pandas as pd
from collections import defaultdict, OrderedDict
from modules.utils import safe_str


def _agrupar(df: pd.DataFrame):
    """
    Retorna estrutura ordenada:
    { menu: { especifico: [descricao, ...] } }
    preservando a ordem de aparição.
    """
    grupos = OrderedDict()

    for _, row in df.iterrows():
        menu       = safe_str(row.get("Menu"))      or "Geral"
        especifico = safe_str(row.get("Especifico")) or ""
        descricao  = safe_str(row.get("Descrição"))

        if not descricao:
            continue

        if menu not in grupos:
            grupos[menu] = OrderedDict()

        chave = especifico if especifico else "__sem_especifico__"
        if chave not in grupos[menu]:
            grupos[menu][chave] = []
        grupos[menu][chave].append(descricao)

    return grupos


def _renderizar_secao(titulo: str, emoji: str, df_secao: pd.DataFrame) -> list:
    linhas = []
    linhas.append(f"{emoji} {titulo}")

    if df_secao.empty:
        linhas.append("    *(nenhum item)*")
        linhas.append("")
        return linhas

    grupos = _agrupar(df_secao)

    for menu, especificos in grupos.items():
        linhas.append(f"➡ {menu}")

        for especifico, descricoes in especificos.items():
            chave = "" if especifico == "__sem_especifico__" else especifico

            if len(descricoes) == 1:
                # Item único: • Específico: Descrição  (ou só descrição se sem específico)
                if chave:
                    linhas.append(f"    • {chave}: {descricoes[0]}")
                else:
                    linhas.append(f"    • {descricoes[0]}")
            else:
                # Múltiplos itens: específico na linha + ◦ para cada descrição
                if chave:
                    linhas.append(f"    • {chave}: ")
                for desc in descricoes:
                    linhas.append(f"        ◦ {desc}")

    linhas.append("")
    return linhas


def gerar_faq_markdown(df: pd.DataFrame, versao: str, caminho_saida: str):
    melhorias = df[df["Tipo"].str.strip().str.lower() == "melhoria"]
    correcoes = df[df["Tipo"].str.strip().str.lower() == "correção"]

    linhas = []
    linhas += _renderizar_secao("Melhorias", "📄", melhorias)
    linhas.append("")
    linhas += _renderizar_secao("Correções", "📄", correcoes)

    with open(caminho_saida, "w", encoding="utf-8") as f:
        f.write("\n".join(linhas))

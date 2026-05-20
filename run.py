"""
Jira FAQ Automation — release_431.1
Uso:
    python run.py --file release_431.1.xlsm --version 431.1 --next-version 432.1
"""
import argparse
import sys
from pathlib import Path
from modules.reader import ler_planilha
from modules.faq_generator import gerar_faq_markdown
from modules.sheet_generator import gerar_planilha_nova_versao
from modules.utils import log


def main():
    parser = argparse.ArgumentParser(
        description="Gera FAQ Markdown e nova planilha a partir de uma release do Jira."
    )
    parser.add_argument("--file", required=True, help="Planilha atual (.xlsm ou .xlsx)")
    parser.add_argument("--version", required=True, help="Versão atual (ex: 431.1)")
    parser.add_argument("--next-version", required=True, help="Próxima versão (ex: 432.1)")
    parser.add_argument("--output", default="output", help="Pasta de saída (padrão: ./output)")
    args = parser.parse_args()

    arquivo = Path(args.file)
    if not arquivo.exists():
        log(f"Arquivo não encontrado: {arquivo}", nivel="erro")
        sys.exit(1)

    versao_atual   = args.version
    proxima_versao = args.next_version
    pasta_saida    = Path(args.output)
    pasta_saida.mkdir(parents=True, exist_ok=True)

    log(f"Lendo planilha: {arquivo}")
    df, avisos = ler_planilha(arquivo)

    if avisos:
        for av in avisos:
            log(av, nivel="aviso")

    if df.empty:
        log("Nenhum dado encontrado na planilha.", nivel="erro")
        sys.exit(1)

    melhorias = df[df["Tipo"].str.strip().str.lower() == "melhoria"]
    correcoes = df[df["Tipo"].str.strip().str.lower() == "correção"]
    log(f"{len(df)} itens lidos → {len(melhorias)} melhorias, {len(correcoes)} correções")

    faq_path = pasta_saida / f"faq_{versao_atual}.md"
    log(f"Gerando FAQ → {faq_path}")
    gerar_faq_markdown(df, versao_atual, str(faq_path))
    log("FAQ gerado com sucesso!")

    planilha_path = pasta_saida / f"release_{proxima_versao}.xlsx"
    log(f"Gerando planilha → {planilha_path}")
    gerar_planilha_nova_versao(arquivo, proxima_versao, str(planilha_path))
    log("Planilha gerada com sucesso!")

    log("=" * 52)
    log(f"Concluído! Arquivos em '{pasta_saida}/':")
    log(f"  faq_{versao_atual}.md")
    log(f"  release_{proxima_versao}.xlsx")


if __name__ == "__main__":
    main()

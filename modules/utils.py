from datetime import datetime

COLUNAS_OBRIGATORIAS = ["Tipo", "Menu", "Descrição"]

def log(mensagem: str, nivel: str = "info"):
    ts = datetime.now().strftime("%H:%M:%S")
    prefixo = {"info": "INFO", "aviso": "AVISO", "erro": "ERRO"}.get(nivel, "INFO")
    print(f"[{ts}] [{prefixo}] {mensagem}")

def safe_str(valor) -> str:
    if valor is None:
        return ""
    try:
        import pandas as pd
        if pd.isna(valor):
            return ""
    except Exception:
        pass
    s = str(valor).strip()
    # Ignorar valores de fórmulas booleanas (=FALSE(), =TRUE())
    if s.lower() in ("false", "true", "0", "1"):
        return ""
    return s

import pandas as pd
import os

CAMINHO_PLANILHAS = "planilhas"
if not os.path.exists(CAMINHO_PLANILHAS):
    os.makedirs(CAMINHO_PLANILHAS)

COLUNAS_MULTIPLAS = [
    "Última atualização:",
    "Imagem Registrada:"
] + [
    f"({b}) Pergunta {i:02d} [Monitoramento]"
    for b in ["PC", "CE", "DF", "DQ"]
    for i in range(1, 32)
]

def extrair_info(df, coluna):
    coluna_busca = "Imagem do Equipamento Registrada em:" if coluna == "Imagem Registrada:" else coluna
    if coluna_busca not in df.columns:
        return pd.DataFrame()

    log = df[coluna_busca].dropna().astype(str)
    if log.empty:
        return pd.DataFrame()

    extra = log.str.extract(r"(?P<nome>.+) - (?P<datahora>\d{2}/\d{2}/\d{4} \d{2}:\d{2}:\d{2})")
    extra["datahora"] = pd.to_datetime(extra["datahora"], format="%d/%m/%Y %H:%M:%S", errors="coerce")
    extra = extra.dropna(subset=["nome", "datahora"])
    
    if extra.empty:
        return pd.DataFrame()

    extra["data"] = extra["datahora"].dt.normalize()
    extra["hora"] = extra["datahora"].dt.strftime("%H:%M:%S")
    extra["semana"] = extra["datahora"].dt.to_period("W").astype(str)
    
    nome_limpo = coluna.split(" [Monitoramento]")[0].strip()
    extra["alterado_no_momento"] = "Imagem Registrada:" if "Imagem do Equipamento" in nome_limpo else nome_limpo
    extra["index_base"] = extra.index
    return extra

def listar_arquivos():
    return [f for f in os.listdir(CAMINHO_PLANILHAS) if f.endswith(".xlsx")]
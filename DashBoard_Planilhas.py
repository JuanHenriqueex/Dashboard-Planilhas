import pandas as pd
import plotly.express as px
import dash
from dash import dcc, html, Input, Output, dash_table, State
import os
import re

app = dash.Dash(__name__)

CAMINHO_PLANILHAS = "planilhas"
ARQUIVOS = [f for f in os.listdir(CAMINHO_PLANILHAS) if f.endswith(".xlsx")]

def extrair_info(df, coluna):
    if coluna not in df.columns:
        return pd.DataFrame()
    log = df[coluna].dropna()
    if log.empty:
        return pd.DataFrame()
    extra = log.str.extract(r"(?P<nome>.+) - (?P<datahora>\d{2}/\d{2}/\d{4} \d{2}:\d{2}:\d{2})")
    extra["coluna"] = coluna
    extra["datahora"] = pd.to_datetime(extra["datahora"], format="%d/%m/%Y %H:%M:%S", errors="coerce")
    extra["data"] = extra["datahora"].dt.date
    extra["index_base"] = log.index
    return extra.dropna(subset=["nome", "datahora"])

COLUNAS_MULTIPLAS = [
    "Última atualização:", "Imagem do Equipamento Registrada em:"
] + [f"({bloco}) Pergunta {i:02d} [Monitoramento]" for bloco in ["PC", "CE", "DF", "DQ"] for i in range(1, 32)]

app.layout = html.Div(
    style={"fontFamily": "Arial", "padding": "20px"},
    children=[
        html.H2("Análise de Alterações"),
        dcc.Dropdown(
            id="arquivo_excel",
            options=[{"label": nome, "value": nome} for nome in ARQUIVOS],
            value=ARQUIVOS[0] if ARQUIVOS else None,
            style={"width": "500px", "marginBottom": "20px"},
        ),
        dcc.Dropdown(
            id="tipo_grafico",
            options=[
                {"label": "Rosca (Última atualização)", "value": "rosquinha_atualizacao"},
                {"label": "Rosca (Imagem do Equipamento)", "value": "rosquinha_imagem"},
                {"label": "Rosca (Multi-Colunas)", "value": "rosquinha_multi"},
                {"label": "Colunas (por dia)", "value": "colunas"},
                {"label": "Colunas (por mês)", "value": "colunas_mes"},
            ],
            value="rosquinha_atualizacao",
            clearable=False,
            style={"width": "400px", "marginBottom": "20px"},
        ),
        dcc.Graph(id="grafico"),
        html.H4("Resumo Mensal de Atualizações", id="titulo_resumo", style={"marginTop": "30px"}),
        dash_table.DataTable(
            id="tabela_quedas_mensais",
            data=[],
            columns=[],
            page_size=12,
            style_table={"overflowX": "auto"},
            style_cell={"textAlign": "left"},
            style_header={"fontWeight": "bold"},
        ),
        html.H4("Detalhes:", id="titulo_tabela", style={"marginTop": "30px"}),
        dash_table.DataTable(
            id="tabela_detalhes",
            data=[], columns=[], page_size=12,
            style_table={"overflowX": "auto"},
            style_cell={"textAlign": "left"},
            style_header={"fontWeight": "bold"},
        ),
        html.Button(
            "Exportar Tabela",
            id="btn_exportar",
            n_clicks=0,
            style={"marginTop": "10px", "marginBottom": "10px"}
        ),
        dcc.Download(id="download_dataframe_csv"),
    ]
)

@app.callback(
    Output("grafico", "figure"),
    Output("tabela_detalhes", "data"),
    Output("tabela_detalhes", "columns"),
    Output("titulo_tabela", "children"),
    Output("tabela_quedas_mensais", "data"),
    Output("tabela_quedas_mensais", "columns"),
    Input("arquivo_excel", "value"),
    Input("tipo_grafico", "value"),
    Input("grafico", "clickData"),
)
def atualizar_tudo(arquivo_nome, tipo, clickData):
    if not arquivo_nome:
        return dash.no_update, [], [], "Detalhes:", [], []
    
    try:
        df = pd.read_excel(os.path.join(CAMINHO_PLANILHAS, arquivo_nome), sheet_name=0)
    except Exception as e:
        return dash.no_update, [], [], f"Erro ao abrir o arquivo: {e}", [], []

    usar_multiplas = tipo in ["rosquinha_multi", "colunas", "colunas_mes"]
    df_uso = pd.DataFrame()

    dados_quedas = []
    colunas_quedas = []

    if usar_multiplas:
        colunas_existentes = list(set(COLUNAS_MULTIPLAS) & set(df.columns))
        if not colunas_existentes:
            return dash.no_update, [], [], "Colunas necessárias não encontradas no arquivo.", [], []
        
        df_info = pd.concat([extrair_info(df, col) for col in colunas_existentes], ignore_index=True)
        if df_info.empty:
            return dash.no_update, [], [], "Nenhum dado para mostrar.", [], []

        df_merged = df_info.merge(df.reset_index(), left_on="index_base", right_on="index", how="left")
        df_uso = df_merged

        if tipo == "rosquinha_multi":
            contagem = df_uso["nome"].value_counts().reset_index()
            contagem.columns = ["nome", "alteracoes"]
            fig = px.pie(contagem, names="nome", values="alteracoes", hole=0.4,
                         title="Total de Alterações por Pessoa (Multi-Colunas)")

        elif tipo == "colunas":
            por_dia = df_uso.groupby(["data", "nome"]).size().reset_index(name="qtd")
            if por_dia.empty:
                return dash.no_update, [], [], "Nenhum dado para mostrar.", [], []

            diario = por_dia.groupby("data").agg(
                pessoas=("nome", lambda x: ", ".join(sorted(set(x)))),
                pessoas_unicas=("nome", "nunique")
            ).reset_index()

            fig = px.bar(
                diario, x="data", y="pessoas_unicas", custom_data=["pessoas"],
                labels={"data": "Data", "pessoas_unicas": "Nº de Pessoas"},
                title="Nº de Pessoas que Alteraram por Dia (Multi-Colunas)",
                color_discrete_sequence=["cornflowerblue"]
            )
            fig.update_traces(hovertemplate="<b>Data:</b> %{x}<br><b>Pessoas:</b> %{y}<br><b>Quem Alterou:</b> %{customdata[0]}<extra></extra>")
            fig.update_layout(xaxis_tickangle=-45)

        else:  # tipo == "colunas_mes"
            df_uso['mes'] = df_uso['datahora'].dt.to_period('M').astype(str)
            por_mes = df_uso.groupby(["mes", "nome"]).size().reset_index(name="qtd")
            if por_mes.empty:
                return dash.no_update, [], [], "Nenhum dado para mostrar.", [], []

            mensal = por_mes.groupby("mes").agg(
                total_atualizacoes=("qtd", "sum"),
                pessoas_unicas=("nome", "nunique"),
                pessoas=("nome", lambda x: ", ".join(sorted(set(x))))
            ).reset_index()

            fig = px.bar(
                mensal, x="mes", y="pessoas_unicas", custom_data=["pessoas"],
                labels={"mes": "Mês", "pessoas_unicas": "Nº de Pessoas"},
                title="Nº de Pessoas que Alteraram por Mês (Multi-Colunas)",
                color_discrete_sequence=["mediumaquamarine"]
            )
            fig.update_traces(hovertemplate="<b>Mês:</b> %{x}<br><b>Pessoas:</b> %{y}<br><b>Quem Alterou:</b> %{customdata[0]}<extra></extra>")
            fig.update_layout(xaxis_tickangle=-45)

            dados_quedas = mensal.to_dict("records")
            colunas_quedas = [
                {"name": "Mês", "id": "mes"},
                {"name": "Pessoas Únicas", "id": "pessoas_unicas"},
                {"name": "Total de Atualizações", "id": "total_atualizacoes"},
                {"name": "Pessoas", "id": "pessoas"},
            ]
            
    else:
        coluna_escolhida = {
            "rosquinha_atualizacao": "Última atualização:",
            "rosquinha_imagem": "Imagem do Equipamento Registrada em:"
        }.get(tipo, "Última atualização:")

        if coluna_escolhida not in df.columns:
            return dash.no_update, [], [], f"Coluna '{coluna_escolhida}' não encontrada.", [], []

        extra = extrair_info(df, coluna_escolhida)
        if extra.empty:
            return dash.no_update, [], [], "Nenhum dado para mostrar.", [], []

        df["nome"] = pd.NA
        df["datahora"] = pd.NaT
        df["data"] = pd.NaT
        df.loc[extra.index, ["nome", "datahora", "data"]] = extra[["nome", "datahora", "data"]]
        df_uso = df

        contagem = df["nome"].value_counts().reset_index()
        contagem.columns = ["nome", "alteracoes"]
        titulo = "Total de Alterações por Pessoa"
        titulo += " (Última atualização)" if tipo == "rosquinha_atualizacao" else " (Imagem do Equipamento)"
        fig = px.pie(contagem, names="nome", values="alteracoes", hole=0.4, title=titulo)

    if not clickData:
        return fig, [], [], "Detalhes:", dados_quedas, colunas_quedas

    if tipo.startswith("rosquinha"):
        pessoa = clickData["points"][0]["label"]
        linhas = df_uso[df_uso["nome"].str.contains(rf"\b{re.escape(pessoa)}\b", na=False, regex=True)]
        if linhas.empty:
            return fig, [], [], f"Detalhes de {pessoa}:", dados_quedas, colunas_quedas

        col_fixas = ["nome", "TAG", "Descrição", "Área", "datahora"]
        col_dinamicas = [c for c in linhas.columns if c not in col_fixas and linhas[c].notna().any()]
        col_final = [c for c in col_fixas + col_dinamicas if c in linhas.columns]

        data = linhas[col_final].dropna(how='all').to_dict("records")
        columns = [{"name": "Datahora" if c == "datahora" else c.capitalize(), "id": c} for c in col_final]
        titulo = f"Detalhes de {pessoa}:"

    else:
        if tipo == "colunas":
            periodo = pd.to_datetime(clickData["points"][0]["x"]).date()
            linhas = df_uso[df_uso["data"] == periodo]
            titulo = f"Alterações em {periodo}:"
            periodo_para_tabela = periodo
        else: # "colunas_mes"
            periodo_str = clickData["points"][0]["x"]
            linhas = df_uso[df_uso["mes"] == periodo_str]
            titulo = f"Alterações em {periodo_str}:"
            periodo_para_tabela = periodo_str

        if linhas.empty:
            return fig, [], [], f"Detalhes de {periodo_para_tabela}:", dados_quedas, colunas_quedas

        col_fixas = ["nome", "TAG", "Descrição", "Área", "datahora"]
        col_dinamicas = [c for c in linhas.columns if c not in col_fixas and linhas[c].notna().any()]
        col_final = [c for c in col_fixas + col_dinamicas if c in linhas.columns]

        data = linhas[col_final].dropna(how='all').to_dict("records")
        columns = [{"name": "Datahora" if c == "datahora" else c.capitalize(), "id": c} for c in col_final]

    return fig, data, columns, titulo, dados_quedas, colunas_quedas

@app.callback(
    Output("download_dataframe_csv", "data"),
    Input("btn_exportar", "n_clicks"),
    State("tabela_detalhes", "data"),
    State("tabela_quedas_mensais", "data"),
    State("tipo_grafico", "value"),
    prevent_initial_call=True,
)
def exportar_tabela(n_clicks, tabela_detalhes_data, tabela_mensal_data, tipo_grafico):
    if n_clicks > 0:
        df_exportar = pd.DataFrame()
        filename = "dados_tabela.csv"
        
        # Lógica para exportar os nomes das pessoas que alteraram no mês
        if tipo_grafico == "colunas_mes" and tabela_mensal_data:
            df_mensal = pd.DataFrame(tabela_mensal_data)
            df_exportar = df_mensal[['mes', 'pessoas']]
            df_exportar.rename(columns={'mes': 'Mês', 'pessoas': 'Pessoas'}, inplace=True)
            filename = "pessoas_por_mes.csv"
        # Comportamento padrão para os outros filtros
        elif tabela_detalhes_data:
            df_exportar = pd.DataFrame(tabela_detalhes_data)
            filename = "detalhes_tabela.csv"
        
        if not df_exportar.empty:
            return dcc.send_data_frame(
                df_exportar.to_csv,
                filename,
                index=False,
                sep=';'
            )
    return None

if __name__ == "__main__":
    app.run(debug=True)
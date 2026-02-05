import pandas as pd
import plotly.express as px
import dash
from dash import dcc, html, Input, Output, dash_table, State
import os

app = dash.Dash(__name__)

# Configuração de pastas e arquivos
CAMINHO_PLANILHAS = "planilhas"
if not os.path.exists(CAMINHO_PLANILHAS):
    os.makedirs(CAMINHO_PLANILHAS)

ARQUIVOS = [f for f in os.listdir(CAMINHO_PLANILHAS) if f.endswith(".xlsx")]

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

COLUNAS_MULTIPLAS = [
    "Última atualização:",
    "Imagem Registrada:"
] + [
    f"({b}) Pergunta {i:02d} [Monitoramento]"
    for b in ["PC", "CE", "DF", "DQ"]
    for i in range(1, 32)
]

app.layout = html.Div(
    style={"fontFamily": "Arial", "padding": "20px"},
    children=[
        html.H2("Dashboard de Alterações"),

        html.Div([
            html.Label("1. Selecione o Arquivo:"),
            dcc.Dropdown(
                id="arquivo_excel",
                options=[{"label": a, "value": a} for a in ARQUIVOS],
                value=ARQUIVOS[0] if ARQUIVOS else None
            ),
        ], style={"width": "400px", "marginBottom": "10px"}),

        html.Div([
            html.Label("2. Filtro de Semana:"),
            dcc.Dropdown(id="filtro_semana", placeholder="Todas as semanas"),
            html.Button("Limpar Semana", id="btn_limpar_semana", n_clicks=0, style={"marginTop": "5px"}),
        ], style={"width": "400px", "marginBottom": "10px"}),

        # --- NOVO FILTRO DE INTERVALO DE DATAS ---
        html.Div([
            html.Label("3. Filtro por Intervalo de Datas:"),
            html.Br(),
            dcc.DatePickerRange(
                id="filtro_data_range",
                min_date_allowed=None,
                max_date_allowed=None,
                start_date=None,
                end_date=None,
                display_format="DD/MM/YYYY",
                clearable=True,
                style={"marginTop": "5px"}
            ),
        ], style={"width": "400px", "marginBottom": "10px"}),

        html.Div([
            html.Label("4. Tipo de Gráfico:"),
            dcc.Dropdown(
                id="tipo_grafico",
                options=[
                    {"label": "Colunas (Tags Únicas/Semana)", "value": "colunas_tags_semana"},
                    {"label": "Rosca (Usuários)", "value": "rosquinha_multi"},
                    {"label": "Colunas (Dias)", "value": "colunas_dias"},
                    {"label": "Colunas (Por TAGs)", "value": "rosquinha_tags"},
                ],
                value="colunas_tags_semana",
                clearable=False
            ),
        ], style={"width": "400px", "marginBottom": "20px"}),

        dcc.Graph(id="grafico"),

        html.H4("Resumo de Alterações por TAG", id="titulo_tags", style={"marginTop": "30px"}),
        dash_table.DataTable(
            id="tabela_tags",
            style_table={'overflowX': 'auto', 'width': '50%'},
            style_cell={'textAlign': 'center', 'padding': '10px'},
            style_header={'fontWeight': 'bold', 'backgroundColor': '#f2f2f2'},
            page_size=10,
            sort_action="native"
        ),

        html.H4("Resumo Dinâmico", id="titulo_resumo", style={"marginTop": "30px"}),
        dash_table.DataTable(id="tabela_resumo_dinamico", style_table={'overflowX': 'auto'}, sort_action="native"),
        
        html.Div(id="container_soma_total", style={"marginTop": "10px", "width": "450px"}, children=[
            dash_table.DataTable(
                id="tabela_soma_total",
                style_cell={'textAlign': 'center', 'fontWeight': 'bold', 'backgroundColor': '#f9f9f9'},
                style_header={'backgroundColor': '#e1e1e1'}
            )
        ]),

        html.H4("Detalhes das Alterações", id="titulo_tabela", style={"marginTop": "30px"}),
        dash_table.DataTable(
            id="tabela_detalhes",
            style_table={'overflowX': 'auto'},
            style_cell={'textAlign': 'left', 'padding': '10px'},
            style_header={'fontWeight': 'bold'},
            sort_action="native"
        ),

        html.Br(),
        html.Button("Exportar CSV", id="btn_exportar", style={
            "marginTop": "20px", 
            "padding": "10px 20px", 
            "backgroundColor": "#28a745", 
            "color": "white", 
            "border": "none", 
            "borderRadius": "5px",
            "cursor": "pointer"
        }),
        dcc.Download(id="download_dataframe_csv"),
    ]
)

# --- CALLBACK PARA ATUALIZAR OPÇÕES DE SEMANA E LIMITES DO CALENDÁRIO ---
@app.callback(
    [Output("filtro_semana", "options"), Output("filtro_semana", "value"),
     Output("filtro_data_range", "min_date_allowed"), Output("filtro_data_range", "max_date_allowed")],
    [Input("arquivo_excel", "value"), Input("btn_limpar_semana", "n_clicks")],
    State("filtro_semana", "value")
)
def atualizar_limites_filtros(arquivo, n_clicks, semana_atual):
    if not arquivo: return [], None, None, None
    
    df = pd.read_excel(os.path.join(CAMINHO_PLANILHAS, arquivo), sheet_name=0)
    lista_dfs = [extrair_info(df, c) for c in COLUNAS_MULTIPLAS if c in df.columns or (c=="Imagem Registrada:" and "Imagem do Equipamento Registrada em:" in df.columns)]
    lista_dfs = [d for d in lista_dfs if not d.empty]
    
    if not lista_dfs: return [], None, None, None
    
    df_total = pd.concat(lista_dfs, ignore_index=True)
    
    semanas = sorted(df_total["semana"].unique())
    options = [{"label": f"Semana {s}", "value": s} for s in semanas]
    
    # Define os limites do calendário baseados nos dados
    min_d = df_total["data"].min()
    max_d = df_total["data"].max()

    return options, (None if n_clicks > 0 else semana_atual), min_d, max_d

# --- CALLBACK PRINCIPAL ---
@app.callback(
    [Output("grafico", "figure"), Output("tabela_detalhes", "data"), Output("tabela_detalhes", "columns"),
     Output("titulo_tabela", "children"), Output("tabela_resumo_dinamico", "data"), Output("tabela_resumo_dinamico", "columns"),
     Output("tabela_tags", "data"), Output("tabela_tags", "columns"), Output("titulo_tags", "style"),
     Output("titulo_resumo", "children"), Output("tabela_soma_total", "data"), Output("tabela_soma_total", "columns")],
    [Input("arquivo_excel", "value"), 
     Input("tipo_grafico", "value"), 
     Input("filtro_semana", "value"), 
     Input("filtro_data_range", "start_date"), # Data Início
     Input("filtro_data_range", "end_date"),   # Data Fim
     Input("grafico", "clickData")]
)
def atualizar_dashboard(arquivo, tipo, semana_sel, start_date, end_date, click):
    if not arquivo: return dash.no_update, [], [], "", [], [], [], [], {"display": "none"}, "Resumo", [], []

    df_original = pd.read_excel(os.path.join(CAMINHO_PLANILHAS, arquivo), sheet_name=0)
    lista_eventos = [extrair_info(df_original, c) for c in COLUNAS_MULTIPLAS if c in df_original.columns or (c=="Imagem Registrada:" and "Imagem do Equipamento Registrada em:" in df_original.columns)]
    lista_eventos = [d for d in lista_eventos if not d.empty]

    if not lista_eventos: return px.pie(title="Sem dados"), [], [], "Sem dados", [], [], [], [], {"display": "none"}, "Resumo", [], []

    df_eventos = pd.concat(lista_eventos, ignore_index=True)
    
    # --- FILTRAGEM ---
    if semana_sel: 
        df_eventos = df_eventos[df_eventos["semana"] == semana_sel]
    
    if start_date and end_date:
        # Converter dates do componente para datetime
        df_eventos = df_eventos[(df_eventos["data"] >= start_date) & (df_eventos["data"] <= end_date)]
    
    df_vinculado = df_eventos.merge(df_original[["TAG"]].reset_index(), left_on="index_base", right_on="index", how="left")

    # Lógica de processamento (Igual à anterior, usando o df_vinculado filtrado)
    tags_data, tags_cols, tags_style = [], [], {"display": "none"}
    res_data, res_cols, titulo_res = [], [], "Resumo Diário"
    soma_data, soma_cols = [], []

    if tipo == "colunas_tags_semana":
        df_semanal = df_vinculado.groupby("semana").agg(tags_unicas=("TAG", "nunique")).reset_index()
        fig = px.bar(df_semanal, x="semana", y="tags_unicas", text="tags_unicas", title="Tags Únicas Alteradas por Semana")
        titulo_res = "Resumo Semanal"
        res_df = df_vinculado.groupby("semana").agg(total_alteracoes=("nome", "count"), tags_unicas=("TAG", "nunique"), pessoas_unicas=("nome", "nunique"), usuarios=("nome", lambda x: ", ".join(sorted(set(x))))).reset_index()
        res_data = res_df.to_dict("records")
        res_cols = [{"name": n, "id": i} for n, i in zip(["Semana", "Total Alt.", "Tags Únicas", "Pessoas", "Usuários"], ["semana", "total_alteracoes", "tags_unicas", "pessoas_unicas", "usuarios"])]
        soma_data = [{"label": "SOMA TOTAL", "valor_alt": res_df["total_alteracoes"].sum(), "valor_tags": df_vinculado["TAG"].nunique()}]
        soma_cols = [{"name": "", "id": "label"}, {"name": "Total Alterações", "id": "valor_alt"}, {"name": "Total Tags Únicas", "id": "valor_tags"}]

    elif tipo == "rosquinha_multi":
        res_user = df_vinculado.groupby("nome").agg(total_alt=("nome", "count"), tags_unicas=("TAG", "nunique")).reset_index().sort_values("total_alt", ascending=False)
        fig = px.pie(res_user, names="nome", values="total_alt", hole=0.4, title="Alterações por Usuário")
        titulo_res = "Resumo Geral por Usuário"
        res_data = res_user.to_dict("records")
        res_cols = [{"name": "Nome do Usuário", "id": "nome"}, {"name": "Total de Alterações", "id": "total_alt"}, {"name": "Tags Únicas Alteradas", "id": "tags_unicas"}]
        soma_data = [{"label": "SOMA TOTAL", "valor_alt": res_user["total_alt"].sum(), "valor_tags": df_vinculado["TAG"].nunique()}]
        soma_cols = [{"name": "", "id": "label"}, {"name": "Total Alterações", "id": "valor_alt"}, {"name": "Total Tags Únicas", "id": "valor_tags"}]

    elif tipo == "colunas_dias":
        diario = df_vinculado.groupby("data").agg(total_alteracoes=("nome", "count"), tags_unicas=("TAG", "nunique"), pessoas_unicas=("nome", "nunique"), usuarios=("nome", lambda x: ", ".join(sorted(set(x))))).reset_index()
        diario["data_str"] = diario["data"].dt.strftime("%d/%m/%Y")
        fig = px.bar(diario, x="data_str", y="total_alteracoes", text="total_alteracoes", title="Volume de Alterações por Dia")
        res_data = diario.to_dict("records")
        res_cols = [{"name": n, "id": i} for n, i in zip(["Data", "Total Alt.", "Tags Únicas", "Pessoas", "Usuários"], ["data_str", "total_alteracoes", "tags_unicas", "pessoas_unicas", "usuarios"])]
        titulo_res = "Resumo Diário"
        soma_data = [{"label": "SOMA TOTAL", "valor_alt": diario["total_alteracoes"].sum(), "valor_tags": df_vinculado["TAG"].nunique()}]
        soma_cols = [{"name": "", "id": "label"}, {"name": "Total Alterações", "id": "valor_alt"}, {"name": "Total Tags Únicas", "id": "valor_tags"}]

    elif tipo == "rosquinha_tags":
        contagem_tags = df_vinculado.groupby("TAG").size().reset_index(name="Total").sort_values("Total", ascending=False)
        fig = px.bar(contagem_tags, x="TAG", y="Total", text="Total", title="Alterações por TAG")
        tags_data = contagem_tags.to_dict("records")
        tags_cols = [{"name": i, "id": i} for i in contagem_tags.columns]
        tags_style = {"display": "block", "marginTop": "30px"}
        soma_data = [{"label": "SOMA TOTAL", "valor_alt": contagem_tags["Total"].sum(), "valor_tags": df_vinculado["TAG"].nunique()}]
        soma_cols = [{"name": "", "id": "label"}, {"name": "Total Alterações", "id": "valor_alt"}, {"name": "Total Tags Únicas", "id": "valor_tags"}]

    data_det, cols_det, titulo_det = [], [], "Clique no gráfico para ver detalhes"
    if click:
        valor = click["points"][0].get("label") or click["points"][0].get("x")
        if tipo == "colunas_dias": fil = df_vinculado[df_vinculado["data"].dt.strftime("%d/%m/%Y") == valor]
        elif tipo == "colunas_tags_semana": fil = df_vinculado[df_vinculado["semana"] == valor]
        elif tipo == "rosquinha_tags": fil = df_vinculado[df_vinculado["TAG"] == valor]
        else: fil = df_vinculado[df_vinculado["nome"] == valor]
        
        if not fil.empty:
            det = fil.merge(df_original[["TAG", "Área", "Tipo", "Sistema"]].reset_index(), left_on="index_base", right_on="index", how="left")
            cols_det = [{"name": n, "id": i} for n, i in zip(["TAG", "Área", "Sistema", "Tipo", "Pergunta/Campo", "Data", "Hora"], ["TAG_y", "Área", "Sistema", "Tipo", "alterado_no_momento", "data", "hora"])]
            data_det = det.to_dict("records")
            titulo_det = f"Detalhes de: {valor}"

    return fig, data_det, cols_det, titulo_det, res_data, res_cols, tags_data, tags_cols, tags_style, titulo_res, soma_data, soma_cols

if __name__ == "__main__":
    app.run(debug=True)
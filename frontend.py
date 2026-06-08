from dash import dcc, html, dash_table

# Definição do Layout
layout = html.Div(
    style={"fontFamily": "Arial", "padding": "20px"},
    children=[
        html.H2("Dashboard de Alterações"),

        html.Div([
            html.Label("1. Selecione o Arquivo:"),
            dcc.Dropdown(id="arquivo_excel"),
        ], style={"width": "400px", "marginBottom": "10px"}),

        html.Div([
            html.Label("2. Filtro de Semana:"),
            dcc.Dropdown(id="filtro_semana", placeholder="Todas as semanas"),
            html.Button("Limpar Semana", id="btn_limpar_semana", n_clicks=0, style={"marginTop": "5px"}),
        ], style={"width": "400px", "marginBottom": "10px"}),

        html.Div([
            html.Label("3. Filtro por Intervalo de Datas:"),
            html.Br(),
            dcc.DatePickerRange(
                id="filtro_data_range",
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
            "marginTop": "20px", "padding": "10px 20px", "backgroundColor": "#28a745", 
            "color": "white", "borderRadius": "5px", "cursor": "pointer"
        }),
        dcc.Download(id="download_dataframe_csv"),
    ]
)
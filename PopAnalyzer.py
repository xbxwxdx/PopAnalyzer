import pandas as pd
import dash
from dash import dcc, html, Input, Output, dash_table
import plotly.express as px
import plotly.graph_objects as go
from datetime import datetime

# Загрузка данных (предполагаем, что данные уже в DataFrame df)
# В вашем случае нужно загрузить Excel
# df = pd.read_excel('your_file.xlsx')
file_path = 'Данные для тестового.xlsx'

sheet_name = 'Data'
# Read the Excel file
df = pd.read_excel(file_path, sheet_name=sheet_name, skiprows=1)

sdf = df.columns.tolist()

print (sdf)

# Для примера создадим синтетические данные на основе вашей структуры
# Здесь я создам данные за июль и август 2025 для двух дистрибьюторов
# В реальности замените на загрузку вашего файла

# Создадим DataFrame с данными за два месяца
# data = {
#     'month': ['2025-07-01', '2025-08-01'] * 10,
#     'distr_name': ['Дистрибьютор 1', 'Дистрибьютор 2'] * 10,
#     'pos_code': list(range(1, 21)),
#     'revenue': [1000, 1200, 800, 1500, 900, 1100, 1300, 1400, 1000, 1200] * 2,
#     'sales_quantity': [10, 12, 8, 15, 9, 11, 13, 14, 10, 12] * 2,
#     'sku_name': ['SKU1', 'SKU2', 'SKU1', 'SKU3', 'SKU2', 'SKU1', 'SKU3', 'SKU4', 'SKU2', 'SKU1'] * 2,
#     'brand_name': ['Бренд 1', 'Бренд 2', 'Бренд 1', 'Бренд 3', 'Бренд 2', 'Бренд 1', 'Бренд 3', 'Бренд 4', 'Бренд 2', 'Бренд 1'] * 2,
#     'category_name': ['Категория 1', 'Категория 2', 'Категория 1', 'Категория 3', 'Категория 2', 'Категория 1', 'Категория 3', 'Категория 4', 'Категория 2', 'Категория 1'] * 2
# }

# df = pd.DataFrame(data)
df['month'] = pd.to_datetime(df['month'])

# Расчет факторов
def calculate_factors(data):
    factors = []
    for (month, distr), group in data.groupby([data['month'].dt.to_period('M'), 'distr_name']):
        total_revenue = group['revenue'].sum()
        total_quantity = group['sales_quantity'].sum()
        unique_tt = group['pos_code'].nunique()
        unique_sku_per_tt = group.groupby('pos_code')['sku_name'].nunique().mean()
        avg_offtake_per_sku = total_quantity / (unique_tt * unique_sku_per_tt) if unique_tt * unique_sku_per_tt > 0 else 0
        avg_price = total_revenue / total_quantity if total_quantity > 0 else 0
        
        factors.append({
            'month': str(month),
            'distr_name': distr,
            'revenue': total_revenue,
            'tt_count': unique_tt,
            'depth': unique_sku_per_tt,
            'offtake_sku': avg_offtake_per_sku,
            'avg_price': avg_price
        })
    return pd.DataFrame(factors)

factors_df = calculate_factors(df)

# Создание дашборда
app = dash.Dash(__name__)

app.layout = html.Div([
    html.H1("📈 Факторный анализ вторичных продаж дистрибьюторов (PoP)", style={'textAlign': 'center'}),
    
    html.Div([
        html.Label("Выберите дистрибьютора:"),
        dcc.Dropdown(
            id='distr-dropdown',
            options=[{'label': d, 'value': d} for d in factors_df['distr_name'].unique()],
            value=factors_df['distr_name'].unique()[0],
            clearable=False
        )
    ], style={'width': '30%', 'margin': '20px'}),
    
    html.Hr(),
    
    html.Div([
        html.Div([
            html.H4("Ключевые показатели (Август 2025)"),
            html.Div(id='kpi-display')
        ], style={'padding': '20px', 'border': '1px solid #ddd', 'borderRadius': '5px'}),
    ]),
    
    html.Hr(),
    
    html.Div([
        html.H3("Факторы выручки (PoP)"),
        dcc.Graph(id='factor-bars')
    ]),
    
    html.Div([
        html.H3("Детализация факторов по месяцам"),
        dash_table.DataTable(
            id='factor-table',
            columns=[
                {'name': 'Месяц', 'id': 'month'},
                {'name': 'Дистрибьютор', 'id': 'distr_name'},
                {'name': 'Выручка', 'id': 'revenue', 'type': 'numeric', 'format': {'specifier': ',.0f'}},
                {'name': 'Кол-во ТТ', 'id': 'tt_count'},
                {'name': 'Глубина (SKU/ТТ)', 'id': 'depth', 'format': {'specifier': '.2f'}},
                {'name': 'Off-take SKU (ед./ТТ)', 'id': 'offtake_sku', 'format': {'specifier': '.2f'}},
                {'name': 'Ср. цена (руб.)', 'id': 'avg_price', 'format': {'specifier': ',.2f'}}
            ],
            style_table={'overflowX': 'auto'},
            style_cell={'textAlign': 'center', 'padding': '10px'},
            style_header={'backgroundColor': '#f4f4f4', 'fontWeight': 'bold'}
        )
    ], style={'marginTop': '30px'}),
    
    html.Hr(),
    
    html.Div([
        html.H3("Верификация: произведение факторов = выручка"),
        dcc.Graph(id='verification-chart')
    ])
])

@app.callback(
    [Output('kpi-display', 'children'),
     Output('factor-bars', 'figure'),
     Output('factor-table', 'data'),
     Output('verification-chart', 'figure')],
    [Input('distr-dropdown', 'value')]
)
def update_dashboard(selected_distr):
    # Фильтрация данных
    distr_data = factors_df[factors_df['distr_name'] == selected_distr].sort_values('month')
    
    # Расчет PoP
    if len(distr_data) >= 2:
        latest = distr_data.iloc[-1]
        previous = distr_data.iloc[-2]
        
        pop_revenue = ((latest['revenue'] - previous['revenue']) / previous['revenue']) * 100
        pop_tt = ((latest['tt_count'] - previous['tt_count']) / previous['tt_count']) * 100 if previous['tt_count'] > 0 else 0
        pop_depth = ((latest['depth'] - previous['depth']) / previous['depth']) * 100 if previous['depth'] > 0 else 0
        pop_offtake = ((latest['offtake_sku'] - previous['offtake_sku']) / previous['offtake_sku']) * 100 if previous['offtake_sku'] > 0 else 0
        pop_price = ((latest['avg_price'] - previous['avg_price']) / previous['avg_price']) * 100 if previous['avg_price'] > 0 else 0
    else:
        latest = distr_data.iloc[0]
        pop_revenue = pop_tt = pop_depth = pop_offtake = pop_price = 0
    
    # KPI блок
    kpi_display = html.Div([
        html.P(f"Выручка: {latest['revenue']:,.0f} руб. ({pop_revenue:+.1f}%)"),
        html.P(f"Кол-во ТТ: {latest['tt_count']} ({pop_tt:+.1f}%)"),
        html.P(f"Глубина: {latest['depth']:.2f} SKU/ТТ ({pop_depth:+.1f}%)"),
        html.P(f"Off-take SKU: {latest['offtake_sku']:.2f} ед./ТТ ({pop_offtake:+.1f}%)"),
        html.P(f"Средняя цена: {latest['avg_price']:.2f} руб. ({pop_price:+.1f}%)")
    ])
    
    # График факторов
    fig_bars = go.Figure()
    for factor in ['tt_count', 'depth', 'offtake_sku', 'avg_price']:
        values = distr_data[factor].tolist()
        months = distr_data['month'].tolist()
        fig_bars.add_trace(go.Bar(
            x=months,
            y=values,
            name=factor.replace('_', ' ').title()
        ))
    
    fig_bars.update_layout(
        title=f"Динамика факторов для {selected_distr}",
        xaxis_title="Месяц",
        yaxis_title="Значение",
        barmode='group'
    )
    
    # Таблица
    table_data = distr_data.to_dict('records')
    
    # График верификации
    distr_data['calculated_revenue'] = (
        distr_data['tt_count'] * 
        distr_data['depth'] * 
        distr_data['offtake_sku'] * 
        distr_data['avg_price']
    )
    
    fig_verify = go.Figure()
    fig_verify.add_trace(go.Scatter(
        x=distr_data['month'],
        y=distr_data['revenue'],
        mode='lines+markers',
        name='Фактическая выручка'
    ))
    fig_verify.add_trace(go.Scatter(
        x=distr_data['month'],
        y=distr_data['calculated_revenue'],
        mode='lines+markers',
        name='Рассчитанная (факторы)'
    ))
    
    fig_verify.update_layout(
        title="Верификация: произведение факторов = выручка",
        xaxis_title="Месяц",
        yaxis_title="Выручка, руб.",
        hovermode='x unified'
    )
    
    return kpi_display, fig_bars, table_data, fig_verify

if __name__ == '__main__':
    app.run(debug=True, port=8050)

import dash
from dash import dcc, html, Input, Output, dash_table
import pandas as pd
import plotly.express as px
import re

# โหลดข้อมูล (สมมติมีคอลัมน์ tuition_fee ที่เป็นข้อความ เช่น '30000 บาท/เทอม')
DATA_PATH = "mytcas_scrape.xlsx"
df = pd.read_excel(DATA_PATH, engine='openpyxl')


# แปลงค่าใช้จ่ายเป็นตัวเลขเพื่อกรองช่วงราคา
def extract_price(text):
    if not isinstance(text, str):
        return None
    m = re.search(r'(\d{3,})', text.replace(',', ''))
    if m:
        return int(m.group(1))
    return None

df['tuition_fee_num'] = df['tuition_fee'].apply(extract_price)

# กรองแถวที่มี tuition_fee_num เป็น None ออกสำหรับกราฟและ slider
df_fee = df[df['tuition_fee_num'].notnull()]

# สร้างแอป Dash
app = dash.Dash(__name__)
app.title = "MyTCAS Tuition Dashboard"

# ตัวเลือกคำค้น
keyword_options = [{'label': k, 'value': k} for k in sorted(df['keyword'].unique())]

app.layout = html.Div(style={'font-family': 'Segoe UI, Tahoma, Geneva, Verdana, sans-serif', 'margin': '20px'}, children=[
    html.H1("📊 MyTCAS Tuition Fee Dashboard", style={'textAlign': 'center', 'color': '#003366'}),
    
    # ตัวกรองคำค้น
    html.Div([
        html.Label("กรองคำค้น:", style={'fontWeight': 'bold'}),
        dcc.Dropdown(
            id='filter-keyword',
            options=keyword_options,
            multi=True,
            placeholder="เลือกคำค้นหา...",
            value=keyword_options[0]['value'] if keyword_options else None
        )
    ], style={'width': '40%', 'display': 'inline-block', 'verticalAlign': 'top', 'marginBottom': '20px'}),
    
    # ตัวกรองช่วงค่าใช้จ่าย
    html.Div([
        html.Label("กรองช่วงค่าใช้จ่าย (บาท):", style={'fontWeight': 'bold'}),
        dcc.RangeSlider(
            id='filter-fee-range',
            min=df_fee['tuition_fee_num'].min() if not df_fee.empty else 0,
            max=df_fee['tuition_fee_num'].max() if not df_fee.empty else 100000,
            step=1000,
            marks={i: f'{i//1000}k' for i in range(0, int(df_fee['tuition_fee_num'].max() + 1000), 10000)},
            value=[df_fee['tuition_fee_num'].min() if not df_fee.empty else 0,
                   df_fee['tuition_fee_num'].max() if not df_fee.empty else 100000]
        ),
        html.Div(id='fee-range-display', style={'marginTop': '10px'})
    ], style={'width': '50%', 'display': 'inline-block', 'verticalAlign': 'top', 'paddingLeft': '40px', 'marginBottom': '20px'}),
    
    # กราฟจำนวนหลักสูตรตามคำค้น
    dcc.Graph(id='count-graph'),
    
    # กราฟสัดส่วนประเภทหลักสูตร
    dcc.Graph(id='program-type-pie'),
    
    # ตารางรายละเอียดหลักสูตร
    html.H2("รายละเอียดหลักสูตร"),
    dash_table.DataTable(
        id='program-table',
        columns=[{"name": i, "id": i} for i in df.columns if i != 'tuition_fee_num'],
        page_size=10,
        style_table={'overflowX': 'auto'},
        style_cell={
            'textAlign': 'left',
            'padding': '5px',
            'minWidth': '150px',
            'whiteSpace': 'normal',
            'fontSize': '14px',
        },
        filter_action="native",
        sort_action="native",
        row_selectable="single",
        selected_rows=[],
    ),
    
    html.Div(id='selected-info', style={'marginTop': 20, 'fontSize': '16px', 'color': '#222'}),
])

@app.callback(
    Output('fee-range-display', 'children'),
    Input('filter-fee-range', 'value')
)
def update_fee_range_text(range_values):
    return f"ช่วงค่าใช้จ่าย: {range_values[0]:,} บาท ถึง {range_values[1]:,} บาท"

@app.callback(
    Output('count-graph', 'figure'),
    Output('program-type-pie', 'figure'),
    Output('program-table', 'data'),
    Input('filter-keyword', 'value'),
    Input('filter-fee-range', 'value')
)
def update_dashboard(selected_keywords, fee_range):
    filtered_df = df.copy()
    
    # กรองคำค้น
    if selected_keywords and len(selected_keywords) > 0:
        filtered_df = filtered_df[filtered_df['keyword'].isin(selected_keywords)]
    
    # กรองค่าใช้จ่าย
    if fee_range and len(fee_range) == 2:
        filtered_df = filtered_df[
            (filtered_df['tuition_fee_num'] >= fee_range[0]) & 
            (filtered_df['tuition_fee_num'] <= fee_range[1])
        ]
    
    # กราฟจำนวนหลักสูตรตามคำค้น
    count_by_keyword = filtered_df['keyword'].value_counts().reset_index()
    count_by_keyword.columns = ['Keyword', 'Number of Programs']
    fig_count = px.bar(count_by_keyword, x='Keyword', y='Number of Programs',
                       title='จำนวนหลักสูตรตามคำค้น',
                       text='Number of Programs',
                       color='Keyword',
                       color_discrete_sequence=px.colors.qualitative.Pastel)
    fig_count.update_layout(xaxis_title='คำค้น', yaxis_title='จำนวนหลักสูตร',
                            plot_bgcolor='white')

    # กราฟสัดส่วนประเภทหลักสูตร
    if 'ประเภทหลักสูตร' in filtered_df.columns and filtered_df['ประเภทหลักสูตร'].notnull().any():
        type_counts = filtered_df['ประเภทหลักสูตร'].value_counts().reset_index()
        type_counts.columns = ['Program Type', 'Count']
        fig_pie = px.pie(type_counts, values='Count', names='Program Type',
                         title='สัดส่วนประเภทหลักสูตร',
                         color_discrete_sequence=px.colors.sequential.RdBu)
    else:
        fig_pie = {}
    
    # เตรียมข้อมูลสำหรับตาราง
    table_data = filtered_df.drop(columns=['tuition_fee_num']).to_dict('records')
    
    return fig_count, fig_pie, table_data

@app.callback(
    Output('selected-info', 'children'),
    Input('program-table', 'selected_rows'),
    Input('program-table', 'data')
)
def display_selected_row(selected_rows, table_data):
    if not selected_rows or len(selected_rows) == 0:
        return "เลือกหลักสูตรในตารางเพื่อดูรายละเอียดเพิ่มเติม"
    
    idx = selected_rows[0]
    program = table_data[idx]
    return html.Div([
        html.H3(f"รายละเอียด: {program.get('program_name', '')}"),
        html.P(f"มหาวิทยาลัย: {program.get('university', '')}"),
        html.P(f"คณะ: {program.get('faculty', '')}"),
        html.P(f"ประเภทหลักสูตร: {program.get('ประเภทหลักสูตร', 'ไม่ระบุ')}"),
        html.P(f"ค่าใช้จ่าย: {program.get('ค่าใช้จ่าย', 'ไม่ระบุ')}"),
        html.A("ดูหน้าเว็บ", href=program.get('url', '#'), target="_blank", style={'color': '#007acc', 'textDecoration': 'underline'}),
    ])

if __name__ == "__main__":
    app.run(debug=True)

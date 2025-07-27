import pandas as pd
import dash
from dash import dcc, html, dash_table
from dash.dependencies import Input, Output
import plotly.express as px
import re
import os

# --- 1. โหลดและเตรียมข้อมูล ---
def load_data(file_path='mytcas_scrape.xlsx'):
    if not os.path.exists(file_path):
        print(f"❌ Error: ไม่พบไฟล์ '{file_path}'. กรุณาตรวจสอบว่าไฟล์อยู่ในไดเรกทอรีเดียวกันกับสคริปต์.")
        return pd.DataFrame(columns=['program_name', 'university', 'faculty', 'tuition_fee', 'program_type', 'tuition_fee_numeric', 'keyword'])
    
    try:
        df = pd.read_excel(file_path)
        print(f"✔️ โหลดข้อมูลจาก '{file_path}' สำเร็จ. จำนวนแถว: {len(df)}")
        
        def parse_tuition_fee(fee_str):
            if isinstance(fee_str, str):
                numbers = re.findall(r'\d[\d,\.]*', fee_str)
                if numbers:
                    try:
                        return float(numbers[0].replace(',', ''))
                    except ValueError:
                        return None
            return None

        df['tuition_fee_numeric'] = df['tuition_fee'].apply(parse_tuition_fee)
        
        df.dropna(subset=['program_name', 'university', 'tuition_fee_numeric', 'keyword'], inplace=True)
        
        df['faculty'] = df['faculty'].fillna('ไม่ระบุคณะ')
        df['program_type'] = df['program_type'].fillna('ไม่ระบุประเภท')

        # *** การเปลี่ยนแปลงที่นี่: สร้าง unique_id สำหรับ value และ display_text สำหรับ label ***
        # unique_id ยังคงเป็นข้อความเต็มเพื่อให้ระบุหลักสูตรได้อย่างเฉพาะเจาะจง
        df['unique_id'] = df['program_name'] + " - " + df['university'] + " (" + df['faculty'] + ")"
        
        # สร้างคอลัมน์ใหม่สำหรับข้อความที่แสดงใน Dropdown (label)
        # โดยอาจจะตัดให้สั้นลง หรือเลือกแสดงเฉพาะชื่อหลักสูตรและมหาวิทยาลัย
        # ตัวอย่าง: แสดงเฉพาะชื่อหลักสูตร และเพิ่มชื่อมหาวิทยาลัยถ้าไม่ซ้ำกัน
        df['display_text'] = df['program_name'] + " (" + df['university'] + ")"
        
        df.sort_values(by=['program_name', 'university'], inplace=True)
        
        return df
    except Exception as e:
        print(f"❌ Error ในการโหลดหรือเตรียมข้อมูล: {e}")
        return pd.DataFrame(columns=['program_name', 'university', 'faculty', 'tuition_fee', 'program_type', 'tuition_fee_numeric', 'keyword'])

df = load_data()

if df.empty:
    print("ไม่สามารถสร้าง Dashboard ได้เนื่องจากไม่มีข้อมูลที่ใช้งานได้. โปรดตรวจสอบไฟล์ mytcas_scrape.xlsx และโครงสร้างข้อมูล.")
    app = dash.Dash(__name__)
    app.layout = html.Div([
        html.H1("ภาพรวมหลักสูตร MyTCAS", style={'textAlign': 'center', 'color': '#0A1172'}),
        html.Div("ไม่สามารถโหลดข้อมูลหลักสูตรได้. กรุณาตรวจสอบไฟล์ 'mytcas_scrape.xlsx' และโครงสร้างข้อมูลให้ถูกต้อง", 
                 style={'textAlign': 'center', 'color': 'red', 'fontSize': '18px', 'marginTop': '50px', 'padding': '20px'})
    ], style={'backgroundColor': '#F8F8F8', 'minHeight': '100vh', 'padding': '20px'})
else:
    app = dash.Dash(__name__, 
                    external_stylesheets=['https://fonts.googleapis.com/css2?family=Kanit:wght@300;400;600&display=swap'],
                    meta_tags=[{'name': 'viewport', 'content': 'width=device-width, initial-scale=1.0'}])

    app.layout = html.Div(style={
        'fontFamily': "'Kanit', sans-serif",
        'backgroundColor': '#F0F8FF',
        'padding': '20px 40px',
        'minHeight': '100vh',
        'color': '#333333'
    }, children=[
        html.H1("ภาพรวมหลักสูตร MyTCAS 🎓", style={
            'textAlign': 'center',
            'color': '#001F54',
            'marginBottom': '30px',
            'fontSize': '3em',
            'fontWeight': '600'
        }),

        html.Div([
            html.Div([
                html.H2("ตัวกรองข้อมูล", style={'color': '#034078', 'marginBottom': '15px'}),
                
                html.P("เลือกประเภทหลักสูตร:", style={'color': '#555555', 'marginBottom': '5px'}),
                dcc.Dropdown(
                    id='keyword-filter-dropdown',
                    options=[
                        {'label': 'ทั้งหมด', 'value': 'all'},
                        {'label': 'วิศวกรรมคอมพิวเตอร์', 'value': 'วิศวกรรม คอมพิวเตอร์'},
                        {'label': 'วิศวกรรมปัญญาประดิษฐ์', 'value': 'วิศวกรรม ปัญญาประดิษฐ์'}
                    ],
                    value='all',
                    clearable=False,
                    style={'backgroundColor': '#FFFFFF', 'borderColor': '#B0C4DE', 'borderRadius': '5px', 'padding': '5px', 'marginBottom': '15px'}
                ),

                html.P("เลือกมหาวิทยาลัย:", style={'color': '#555555', 'marginBottom': '5px'}),
                dcc.Dropdown(
                    id='university-filter-dropdown',
                    options=[{'label': 'ทั้งหมด', 'value': 'all'}],
                    multi=True,
                    value=['all'],
                    placeholder="เลือกมหาวิทยาลัย...",
                    style={'backgroundColor': '#FFFFFF', 'borderColor': '#B0C4DE', 'borderRadius': '5px', 'padding': '5px', 'marginBottom': '15px'}
                ),

                html.P("เลือกคณะ:", style={'color': '#555555', 'marginBottom': '5px'}),
                dcc.Dropdown(
                    id='faculty-filter-dropdown',
                    options=[{'label': 'ทั้งหมด', 'value': 'all'}],
                    multi=True,
                    value=['all'],
                    placeholder="เลือกคณะ...",
                    style={'backgroundColor': '#FFFFFF', 'borderColor': '#B0C4DE', 'borderRadius': '5px', 'padding': '5px', 'marginBottom': '25px'}
                ),

                html.P("เลือกหลักสูตรที่ต้องการเปรียบเทียบค่าเทอม:", style={'color': '#555555', 'marginBottom': '15px'}),
                dcc.Dropdown(
                    id='program-dropdown',
                    options=[],
                    multi=True,
                    placeholder="ค้นหาและเลือกหลักสูตร...",
                    style={
                        'backgroundColor': '#FFFFFF', 
                        'borderColor': '#B0C4DE',
                        'borderRadius': '5px',
                        'padding': '5px'
                    },
                    clearable=True
                ),
            ], style={
                'flex': '1',
                'marginRight': '20px',
                'padding': '25px', 
                'backgroundColor': '#FFFFFF', 
                'borderRadius': '10px', 
                'boxShadow': '0 4px 12px rgba(0, 0, 0, 0.1)',
                'border': '1px solid #E0E0E0'
            }),

            html.Div([
                html.H2("ข้อมูลสรุปค่าเทอม", style={'color': '#034078', 'marginBottom': '15px'}),
                html.Div(id='summary-output', style={'padding': '10px', 'color': '#333333', 'lineHeight': '1.8'}),
            ], style={
                'flex': '1',
                'padding': '25px', 
                'backgroundColor': '#FFFFFF', 
                'borderRadius': '10px', 
                'boxShadow': '0 4px 12px rgba(0, 0, 0, 0.1)',
                'border': '1px solid #E0E0E0'
            })
        ], style={'display': 'flex', 'justifyContent': 'space-between', 'marginBottom': '40px'}),

        html.Hr(style={'borderColor': '#B0C4DE', 'borderWidth': '1px', 'borderStyle': 'solid', 'margin': '40px 0'}),

        html.H2("ค่าเทอมเฉลี่ยภาพรวม 📈", style={
            'textAlign': 'center',
            'color': '#001F54',
            'marginBottom': '25px',
            'fontSize': '2.2em'
        }),
        html.Div([
            dcc.Graph(id='overview-tuition-chart', config={'displayModeBar': False})
        ], style={
            'padding': '25px', 
            'backgroundColor': '#FFFFFF', 
            'borderRadius': '10px', 
            'boxShadow': '0 4px 12px rgba(0, 0, 0, 0.1)',
            'border': '1px solid #E0E0E0',
            'marginBottom': '40px'
        }),

        html.Hr(style={'borderColor': '#B0C4DE', 'borderWidth': '1px', 'borderStyle': 'solid', 'margin': '40px 0'}),

        html.H2("กราฟเปรียบเทียบค่าเทอม 📊", style={
            'textAlign': 'center',
            'color': '#001F54',
            'marginBottom': '25px',
            'fontSize': '2.2em'
        }),

        html.Div([
            dcc.Graph(id='tuition-fee-comparison-chart', config={'displayModeBar': False})
        ], style={
            'padding': '25px', 
            'backgroundColor': '#FFFFFF', 
            'borderRadius': '10px', 
            'boxShadow': '0 4px 12px rgba(0, 0, 0, 0.1)',
            'border': '1px solid #E0E0E0',
            'marginBottom': '40px'
        }),

        html.Hr(style={'borderColor': '#B0C4DE', 'borderWidth': '1px', 'borderStyle': 'solid', 'margin': '40px 0'}),

        html.H2("รายละเอียดหลักสูตรทั้งหมด 📋", style={
            'textAlign': 'center',
            'color': '#001F54',
            'marginBottom': '25px',
            'fontSize': '2.2em'
        }),

        html.Div(id='all-programs-table', style={
            'padding': '25px', 
            'backgroundColor': '#FFFFFF', 
            'borderRadius': '10px', 
            'boxShadow': '0 4px 12px rgba(0, 0, 0, 0.1)',
            'border': '1px solid #E0E0E0'
        })
    ])

    # --- 4. สร้าง Callbacks สำหรับการโต้ตอบ ---

    @app.callback(
        [Output('university-filter-dropdown', 'options'),
         Output('university-filter-dropdown', 'value'),
         Output('faculty-filter-dropdown', 'options'),
         Output('faculty-filter-dropdown', 'value'),
         Output('program-dropdown', 'options'),
         Output('program-dropdown', 'value')],
        [Input('keyword-filter-dropdown', 'value'),
         Input('university-filter-dropdown', 'value'),
         Input('faculty-filter-dropdown', 'value')]
    )
    def update_filter_options(selected_keyword, selected_universities_current, selected_faculties_current):
        current_filtered_df = df.copy()

        if selected_keyword and selected_keyword != 'all':
            current_filtered_df = current_filtered_df[current_filtered_df['keyword'] == selected_keyword]
        
        university_options = [{'label': 'ทั้งหมด', 'value': 'all'}]
        university_options.extend([{'label': uni, 'value': uni} for uni in sorted(current_filtered_df['university'].unique())])

        new_uni_value = ['all']
        if selected_universities_current:
            if 'all' in selected_universities_current and len(selected_universities_current) > 1:
                new_uni_value = ['all']
            elif 'all' not in selected_universities_current:
                valid_selected_unis = [uni for uni in selected_universities_current if uni in [opt['value'] for opt in university_options]]
                if valid_selected_unis:
                    new_uni_value = valid_selected_unis
                else:
                    new_uni_value = ['all']
        
        if new_uni_value and 'all' not in new_uni_value:
             current_filtered_df = current_filtered_df[current_filtered_df['university'].isin(new_uni_value)]

        faculty_options = [{'label': 'ทั้งหมด', 'value': 'all'}]
        faculty_options.extend([{'label': fac, 'value': fac} for fac in sorted(current_filtered_df['faculty'].unique())])
        
        new_fac_value = ['all']
        if selected_faculties_current:
            if 'all' in selected_faculties_current and len(selected_faculties_current) > 1:
                new_fac_value = ['all']
            elif 'all' not in selected_faculties_current:
                valid_selected_facs = [fac for fac in selected_faculties_current if fac in [opt['value'] for opt in faculty_options]]
                if valid_selected_facs:
                    new_fac_value = valid_selected_facs
                else:
                    new_fac_value = ['all']

        if new_fac_value and 'all' not in new_fac_value:
            current_filtered_df = current_filtered_df[current_filtered_df['faculty'].isin(new_fac_value)]

        # *** การเปลี่ยนแปลงที่นี่: ใช้ 'display_text' สำหรับ label และ 'unique_id' สำหรับ value ***
        program_options = [{'label': row['display_text'], 'value': row['unique_id']} for index, row in current_filtered_df.iterrows()]
        
        return university_options, new_uni_value, faculty_options, new_fac_value, program_options, [] 

    @app.callback(
        Output('overview-tuition-chart', 'figure'),
        [Input('keyword-filter-dropdown', 'value'),
         Input('university-filter-dropdown', 'value'),
         Input('faculty-filter-dropdown', 'value')]
    )
    def update_overview_chart(selected_keyword, selected_universities, selected_faculties):
        filtered_df_for_overview = df.copy()

        if selected_keyword and selected_keyword != 'all':
            filtered_df_for_overview = filtered_df_for_overview[filtered_df_for_overview['keyword'] == selected_keyword]
        
        if selected_universities and 'all' not in selected_universities:
            filtered_df_for_overview = filtered_df_for_overview[filtered_df_for_overview['university'].isin(selected_universities)]

        if selected_faculties and 'all' not in selected_faculties:
            filtered_df_for_overview = filtered_df_for_overview[filtered_df_for_overview['faculty'].isin(selected_faculties)]
        
        if filtered_df_for_overview.empty:
            fig_overview = px.bar()
            fig_overview.update_layout(
                title='<span style="color:#001F54;"><b>ไม่พบข้อมูลสำหรับภาพรวม</b></span>',
                plot_bgcolor='#F8F8F8', paper_bgcolor='#FFFFFF',
                font=dict(color='#333333', family="'Kanit', sans-serif")
            )
            return fig_overview

        avg_tuition_by_uni = filtered_df_for_overview.groupby('university')['tuition_fee_numeric'].mean().reset_index()
        avg_tuition_by_uni = avg_tuition_by_uni.sort_values(by='tuition_fee_numeric', ascending=False)

        fig_overview = px.bar(
            avg_tuition_by_uni,
            x='university',
            y='tuition_fee_numeric',
            title='<span style="color:#001F54;"><b>ค่าเทอมเฉลี่ยตามมหาวิทยาลัย</b></span>',
            labels={
                'university': 'มหาวิทยาลัย',
                'tuition_fee_numeric': 'ค่าเทอมเฉลี่ย (บาท/ปี)'
            },
            color='university',
            color_discrete_sequence=px.colors.qualitative.Pastel
        )
        fig_overview.update_layout(
            xaxis_title='มหาวิทยาลัย',
            yaxis_title='ค่าเทอมเฉลี่ย (บาท/ปี)',
            plot_bgcolor='#F8F8F8',
            paper_bgcolor='#FFFFFF',
            font=dict(color='#333333', family="'Kanit', sans-serif"),
            title_font=dict(color='#001F54', size=20),
            margin=dict(l=40, r=40, t=80, b=40)
        )
        return fig_overview

    @app.callback(
        [Output('tuition-fee-comparison-chart', 'figure'),
         Output('summary-output', 'children'),
         Output('all-programs-table', 'children')],
        [Input('program-dropdown', 'value'),
         Input('keyword-filter-dropdown', 'value'),
         Input('university-filter-dropdown', 'value'),
         Input('faculty-filter-dropdown', 'value')]
    )
    def update_dashboard(selected_programs_unique_ids, selected_keyword, selected_universities, selected_faculties):
        display_df = df.copy()

        if selected_keyword and selected_keyword != 'all':
            display_df = display_df[display_df['keyword'] == selected_keyword]
        
        if selected_universities and 'all' not in selected_universities:
            display_df = display_df[display_df['university'].isin(selected_universities)]

        if selected_faculties and 'all' not in selected_faculties:
            display_df = display_df[display_df['faculty'].isin(selected_faculties)]

        comparison_df = pd.DataFrame()
        summary_text = html.P("โปรดเลือกหลักสูตรอย่างน้อยหนึ่งรายการจาก Dropdown ด้านบนเพื่อเปรียบเทียบค่าเทอม")
        
        if selected_programs_unique_ids:
            comparison_df = df[df['unique_id'].isin(selected_programs_unique_ids)].copy() 
            
            if not comparison_df.empty:
                fig = px.bar(
                    comparison_df,
                    x='program_name',
                    y='tuition_fee_numeric',
                    color='university',
                    barmode='group',
                    title='<span style="color:#001F54;"><b>เปรียบเทียบค่าเทอมหลักสูตรที่เลือก</b></span>',
                    labels={
                        'program_name': 'ชื่อหลักสูตร',
                        'tuition_fee_numeric': 'ค่าเทอม (บาท/ปี)',
                        'university': 'มหาวิทยาลัย'
                    },
                    hover_data={
                        'tuition_fee': True, 
                        'program_type': True, 
                        'faculty': True,
                        'tuition_fee_numeric': ':.2f'
                    },
                    color_discrete_sequence=px.colors.sequential.Blues_r
                )
                fig.update_layout(
                    xaxis={'categoryorder':'total ascending'},
                    yaxis_title='ค่าเทอม (บาท/ปี)',
                    xaxis_title='หลักสูตร',
                    plot_bgcolor='#F8F8F8',
                    paper_bgcolor='#FFFFFF',
                    font=dict(color='#333333', family="'Kanit', sans-serif"),
                    title_font=dict(color='#001F54', size=20),
                    margin=dict(l=40, r=40, t=80, b=40)
                )
                
                min_fee_program = comparison_df.loc[comparison_df['tuition_fee_numeric'].idxmin()]
                max_fee_program = comparison_df.loc[comparison_df['tuition_fee_numeric'].idxmax()]
                avg_fee = comparison_df['tuition_fee_numeric'].mean()

                summary_text = html.Div([
                    html.P(f"✨ จำนวนหลักสูตรที่เลือก: {len(comparison_df)} หลักสูตร"),
                    html.P(f"💰 ค่าเทอมเฉลี่ย: {avg_fee:,.2f} บาท/ปี"),
                    html.P([
                        html.Span("⬇️ ค่าเทอมต่ำสุด: "),
                        html.B(f"{min_fee_program['program_name']} ({min_fee_program['university']})"),
                        f" - {min_fee_program['tuition_fee']}"
                    ]),
                    html.P([
                        html.Span("⬆️ ค่าเทอมสูงสุด: "),
                        html.B(f"{max_fee_program['program_name']} ({max_fee_program['university']})"),
                        f" - {max_fee_program['tuition_fee']}"
                    ]),
                ])
            else:
                fig = px.bar()
                fig.update_layout(
                    title='<span style="color:#001F54;"><b>ไม่พบข้อมูลสำหรับหลักสูตรที่เลือก</b></span>',
                    plot_bgcolor='#F8F8F8', paper_bgcolor='#FFFFFF'
                )
                summary_text = html.P("ไม่พบข้อมูลหลักสูตรที่ตรงกับที่คุณเลือก. โปรดลองใหม่อีกครั้ง.")
        else:
            fig = px.bar()
            fig.update_layout(
                title='<span style="color:#001F54;"><b>เลือกหลักสูตรเพื่อเปรียบเทียบค่าเทอม</b></span>',
                plot_bgcolor='#F8F8F8', paper_bgcolor='#FFFFFF'
            )

        all_programs_table = dash_table.DataTable(
            id='table-container',
            columns=[
                {"name": "ชื่อหลักสูตร", "id": "program_name", "type": "text"},
                {"name": "มหาวิทยาลัย", "id": "university", "type": "text"},
                {"name": "คณะ", "id": "faculty", "type": "text"},
                {"name": "ประเภทหลักสูตร", "id": "program_type", "type": "text"},
                {"name": "ค่าใช้จ่าย", "id": "tuition_fee", "type": "text"},
                {"name": "คำค้นหลัก", "id": "keyword", "type": "text"}
            ],
            data=display_df.to_dict('records'),
            filter_action="native",
            sort_action="native",
            page_size=10,
            style_header={
                'backgroundColor': '#034078',
                'color': 'white',
                'fontWeight': 'bold',
                'textAlign': 'left',
                'fontFamily': "'Kanit', sans-serif"
            },
            style_data={
                'backgroundColor': 'white',
                'color': '#333333',
                'fontFamily': "'Kanit', sans-serif"
            },
            style_cell={
                'padding': '10px',
                'border': '1px solid #E0E0E0',
                'whiteSpace': 'normal',
                'height': 'auto',
                'minWidth': '100px', 'width': '150px', 'maxWidth': '300px',
                'textAlign': 'left'
            },
            style_data_conditional=[
                {
                    'if': {'row_index': 'odd'},
                    'backgroundColor': '#F8F8F8'
                }
            ],
            export_format="xlsx",
            export_headers="display"
        )

        return fig, summary_text, all_programs_table

if __name__ == '__main__':
    try:
        app.run(debug=True, port=8050)
    except NameError:
        print("Dash App ไม่ได้ถูกสร้างขึ้นเนื่องจากไม่มีข้อมูลหรือมีข้อผิดพลาดร้ายแรงในการเตรียมข้อมูล. กรุณาแก้ไขไฟล์ 'mytcas_scrape.xlsx' หรือตรวจสอบการประมวลผลข้อมูล.")
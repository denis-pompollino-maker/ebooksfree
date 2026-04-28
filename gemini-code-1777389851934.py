import streamlit as st
import pandas as pd
import plotly.graph_objects as go
import plotly.express as px
import calendar
from datetime import datetime, timedelta

# 1. CONFIGURAÇÃO DA PÁGINA
st.set_page_config(layout="wide", page_title="Industrial Analytics Hub", page_icon="📈")

# --- ESTILIZAÇÃO CSS CORPORATIVA (PREMIUM LIGHT MODE) ---
st.markdown("""
    <style>
    @import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700;900&display=swap');
    * { font-family: 'Inter', sans-serif; }
    
    /* Fundo da Aplicação Global */
    .stApp { background-color: #f8fafc; color: #0f172a; }
    
    /* Remover padding superior excessivo do Streamlit */
    .block-container { padding-top: 2rem !important; }

    /* Menu Lateral Refinado */
    [data-testid="stSidebar"] { background-color: #ffffff; border-right: 1px solid #e2e8f0; box-shadow: 2px 0 10px rgba(0,0,0,0.02); }
    [data-testid="stSidebar"] .stRadio div[role="radiogroup"] > label {
        background-color: #ffffff; border: 1px solid transparent;
        padding: 10px 15px !important; border-radius: 8px !important;
        margin-bottom: 8px !important; color: #475569 !important; font-weight: 500; transition: all 0.2s ease;
    }
    [data-testid="stSidebar"] .stRadio div[role="radiogroup"] > label:hover {
        background-color: #f1f5f9; border: 1px solid #e2e8f0;
    }
    [data-testid="stSidebar"] .stRadio div[role="radiogroup"] > label[data-checked="true"] {
        background: linear-gradient(135deg, #1d4ed8 0%, #1e40af 100%) !important;
        color: white !important; font-weight: 600; box-shadow: 0 4px 6px -1px rgba(29, 78, 216, 0.2);
    }

    /* Cards de Métricas Estilo Neumorphism Soft */
    .metric-container { display: flex; justify-content: space-between; gap: 15px; margin-bottom: 20px; }
    .metric-card {
        background: #ffffff; padding: 20px; border-radius: 12px;
        text-align: center; border: 1px solid #f1f5f9;
        flex: 1; display: flex; flex-direction: column; justify-content: center;
        box-shadow: 0 4px 6px -1px rgba(0, 0, 0, 0.05), 0 2px 4px -1px rgba(0, 0, 0, 0.03);
        transition: transform 0.2s ease, box-shadow 0.2s ease;
    }
    .metric-card:hover { transform: translateY(-2px); box-shadow: 0 10px 15px -3px rgba(0, 0, 0, 0.08); }
    .metric-title { color: #64748b; font-size: 0.75rem; font-weight: 600; text-transform: uppercase; letter-spacing: 0.5px; margin-bottom: 8px; }
    .metric-value { color: #0f172a; font-size: 1.8rem; font-weight: 900; line-height: 1.2; }
    
    /* Classes de GAP Refinadas */
    .gap-box { background: #f8fafc; border-radius: 6px; padding: 8px; margin-top: 10px; border: 1px solid #e2e8f0; }
    .gap-negative { color: #e11d48; font-size: 0.85rem; font-weight: 700; display: flex; align-items: center; justify-content: center; gap: 4px; }
    .gap-positive { color: #059669; font-size: 0.85rem; font-weight: 700; display: flex; align-items: center; justify-content: center; gap: 4px; }

    /* Calendário */
    .calendar-day-name { text-align: center; font-weight: 700; color: #475569; font-size: 0.85rem; padding-bottom: 10px; border-bottom: 2px solid #e2e8f0; margin-bottom: 10px;}
    .calendar-grid { display: grid; grid-template-columns: repeat(7, 1fr); gap: 10px; width: 100%; }
    .day-card { background: #ffffff; border-radius: 8px; padding: 12px; min-height: 95px; border: 1px solid #e2e8f0; box-shadow: 0 1px 3px rgba(0,0,0,0.05); }
    .day-number { font-size: 1.1rem; font-weight: 900; color: #1e293b; }
    .day-status { font-size: 0.8rem; font-weight: 700; margin-top: 5px; text-align: right; }

    /* Headers de Seção Estilo Profissional */
    .section-header {
        display: flex; align-items: center; gap: 10px;
        color: #1e293b; font-weight: 800; font-size: 1.2rem;
        margin-top: 30px; margin-bottom: 15px; padding-bottom: 8px;
        border-bottom: 2px solid #e2e8f0;
    }
    .section-header::before { content: ''; display: block; width: 6px; height: 20px; background: #1d4ed8; border-radius: 3px; }

    /* 5 Porquês e Caixas de Texto */
    .five-why-box { border: 1px solid #cbd5e1; padding: 25px; background: #ffffff; border-radius: 12px; margin-top: 20px; box-shadow: 0 4px 6px rgba(0,0,0,0.02); }
    .five-why-line { border-bottom: 1px dashed #cbd5e1; padding: 12px 0; font-size: 0.95rem; color: #334155; }
    </style>
    """, unsafe_allow_html=True)

# 2. CARREGAMENTO E LIMPEZA
@st.cache_data
def load_data(file_obj):
    df_order = pd.read_excel(file_obj, sheet_name="Result by order")
    df_stops = pd.read_excel(file_obj, sheet_name="Stop machine item")
    df_order.columns = df_order.columns.str.strip()
    df_stops.columns = df_stops.columns.str.strip()
    
    nums = ['Run Time', 'Horário Padrão', 'Machine Counter', 'Peças Estoque - Ajuste', 'Average Speed', 'Minutos', 'QTD']
    for df in [df_order, df_stops]:
        for col in nums:
            if col in df.columns:
                df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0)

    df_order['Data'] = pd.to_datetime(df_order['Data'], errors='coerce')
    df_order = df_order.dropna(subset=['Data'])
    df_order['Máquina'] = df_order['Máquina'].fillna(0).astype(int).astype(str)
    df_order['Turno'] = df_order['Turno'].fillna(0).astype(int).astype(str)
    
    df_stops['Data'] = pd.to_datetime(df_stops['Data'], errors='coerce')
    df_stops = df_stops.dropna(subset=['Data'])
    df_stops['Máquina'] = df_stops['Máquina'].fillna(0).astype(int).astype(str)
    df_stops['Turno'] = df_stops['Turno'].fillna(0).astype(int).astype(str)

    def categorize(m): return "BABY" if m in ['2', '3', '4', '5', '6'] else "ADULTO"
    df_order['Categoria'] = df_order['Máquina'].apply(categorize)
    
    return df_order, df_stops

@st.cache_data
def load_metas_completas(file, data_ref):
    try:
        xls = pd.ExcelFile(file)
        target = next((s for s in xls.sheet_names if "PEÇAS" in s.upper()), None)
        df_raw = pd.read_excel(file, sheet_name=target, header=None)
        
        row_dates = df_raw.iloc[2, :].tolist() 
        row_meta_geral = df_raw.iloc[124, :].tolist() 
        
        meta_geral_mes = 0
        meta_ate_hoje = 0
        col_idx_hoje = None
        
        mapping_maquinas = {"1": 6, "2": 28, "3": 47, "4": 58, "5": 77, "6": 96, "7": 113}
        metas_maq_hoje = {}

        for col_idx, d_val in enumerate(row_dates):
            if isinstance(d_val, (datetime, pd.Timestamp)):
                if d_val.month == data_ref.month and d_val.year == data_ref.year:
                    valor_geral = pd.to_numeric(row_meta_geral[col_idx], errors='coerce') or 0
                    meta_geral_mes += valor_geral
                    if d_val.date() <= data_ref:
                        meta_ate_hoje += valor_geral
                if d_val.date() == data_ref:
                    col_idx_hoje = col_idx

        if col_idx_hoje is not None:
            for maq, row_idx in mapping_maquinas.items():
                val_maq = pd.to_numeric(df_raw.iloc[row_idx, col_idx_hoje], errors='coerce') or 0
                metas_maq_hoje[maq] = val_maq
                
        return meta_geral_mes, meta_ate_hoje, metas_maq_hoje
    except Exception as e:
        return 0, 0, {}

# Gráfico de Velocímetro Profissional
def mini_gauge(label, value, color, target, height=180):
    fig = go.Figure(go.Indicator(
        mode="gauge+number", value=value,
        number={'suffix': "%", 'font': {'size': 24, 'color': '#0f172a', 'family': 'Inter', 'weight': 'bold'}},
        title={'text': label, 'font': {'size': 14, 'color': '#64748b', 'family': 'Inter'}},
        gauge={
            'axis': {'range': [0, 100], 'tickwidth': 1, 'tickcolor': "#cbd5e1"},
            'bar': {'color': color, 'thickness': 0.75},
            'bgcolor': "#f1f5f9",
            'borderwidth': 0,
            'threshold': {'line': {'color': "#0f172a", 'width': 3}, 'thickness': 0.75, 'value': target}
        }
    ))
    fig.update_layout(height=height, margin=dict(l=10, r=10, t=40, b=10), paper_bgcolor='rgba(0,0,0,0)', font={'family': 'Inter'})
    return fig

# Tema padrão para gráficos Plotly
def apply_corporate_layout(fig):
    fig.update_layout(
        paper_bgcolor='rgba(0,0,0,0)', plot_bgcolor='rgba(0,0,0,0)',
        font=dict(family='Inter', color='#334155'),
        xaxis=dict(showgrid=True, gridcolor='#e2e8f0', zeroline=False),
        yaxis=dict(showgrid=True, gridcolor='#e2e8f0', zeroline=False),
        margin=dict(l=0, r=0, t=30, b=0)
    )
    return fig

# --- SIDEBAR ---
with st.sidebar:
    st.markdown("<h2 style='color:#1d4ed8; font-weight:900;'>🏭 ANALYTICS HUB</h2>", unsafe_allow_html=True)
    uploaded_file = st.file_uploader("📂 Base de Produção (.xlsm)", type=["xlsm"])
    up_datas = st.file_uploader("📂 Base de Metas (.xlsx)", type=["xlsx"])
    st.markdown("---")
    if uploaded_file:
        menu = st.radio("NAVEGAÇÃO", ["📋 REPORTE DIÁRIO", "📈 PERFORMANCE", "🛑 TOP 10 PARADAS", "📅 CALENDÁRIO", "📋 ANÁLISE SEMANAL"])

if uploaded_file:
    df_order, df_stops = load_data(uploaded_file)

    # =========================================================
    # ABA 1: REPORTE DIÁRIO
    # =========================================================
    if menu == "📋 REPORTE DIÁRIO":
        st.markdown("<div class='section-header'>Filtros Rápidos</div>", unsafe_allow_html=True)
        col_f1, col_f2 = st.columns(2)
        with col_f1:
            data_ref_reporte = st.date_input("Data de Referência", df_order['Data'].max().date())
        with col_f2:
            datas_disp = sorted(df_order['Data'].dt.date.unique().tolist(), reverse=True)
            dias_sel = st.multiselect("Comparar com (Tabelas):", datas_disp, default=[data_ref_reporte])

        st.markdown(f"<h1 style='color:#0f172a; margin-top:20px; font-weight:900;'>Reporte Executivo: {data_ref_reporte.strftime('%d/%m/%Y')}</h1>", unsafe_allow_html=True)

        df_acumulado_mes = df_order[(df_order['Data'].dt.month == data_ref_reporte.month) & (df_order['Data'].dt.year == data_ref_reporte.year) & (df_order['Data'].dt.date <= data_ref_reporte)]
        estoque_acum_mes = df_acumulado_mes['Peças Estoque - Ajuste'].sum()
        total_mc_mes = df_acumulado_mes['Machine Counter'].sum()
        mov_acum_mes = (df_acumulado_mes['Run Time'].sum() / df_acumulado_mes['Horário Padrão'].sum() * 100) if df_acumulado_mes['Horário Padrão'].sum() > 0 else 0
        loss_acum_mes = ((total_mc_mes - estoque_acum_mes) / total_mc_mes * 100) if total_mc_mes > 0 else 0
        
        meta_mov, meta_loss = 90.0, 2.5
        gap_mov, gap_loss = mov_acum_mes - meta_mov, loss_acum_mes - meta_loss
        meta_geral_mes, meta_dinamica_hoje, metas_maq_hoje = load_metas_completas(up_datas, data_ref_reporte) if up_datas else (0, 0, {})

        # Visão Geral do Mês
        st.markdown(f"""
            <div class="metric-container">
                <div class="metric-card">
                    <div class="metric-title">Movimentação Mês</div>
                    <div class="metric-value" style="color:{'#059669' if mov_acum_mes>=meta_mov else '#e11d48'}">{mov_acum_mes:.1f}%</div>
                    <div class="gap-box">Meta 90% | {'✅' if gap_mov>=0 else '🚨'} {gap_mov:+.1f}%</div>
                </div>
                <div class="metric-card">
                    <div class="metric-title">Perda (Loss) Mês</div>
                    <div class="metric-value" style="color:{'#059669' if loss_acum_mes<=meta_loss else '#e11d48'}">{loss_acum_mes:.1f}%</div>
                    <div class="gap-box">Meta 2.5% | {'✅' if gap_loss<=0 else '🚨'} {gap_loss:+.1f}%</div>
                </div>
                <div class="metric-card">
                    <div class="metric-title">Estoque Realizado (Mês)</div>
                    <div class="metric-value" style="color:#1d4ed8">{estoque_acum_mes:,.0f} <span style="font-size:1rem; color:#64748b">pçs</span></div>
                    <div class="gap-box">Meta Dinâmica Hoje: {meta_dinamica_hoje:,.0f}</div>
                </div>
            </div>
        """, unsafe_allow_html=True)

        # Análise de Gap Diário
        if up_datas:
            st.markdown(f"<div class='section-header'>Acompanhamento de Metas por Máquina (Gap)</div>", unsafe_allow_html=True)
            df_dia_gap = df_order[df_order['Data'].dt.date == data_ref_reporte]
            res_gap = df_dia_gap.groupby('Máquina').agg({'Peças Estoque - Ajuste':'sum'}).reset_index()
            maquinas_alvo = ["1", "2", "3", "4", "5", "6", "7"]
            
            cols_gap = st.columns(len(maquinas_alvo))
            for i, maq in enumerate(maquinas_alvo):
                realizado = res_gap.loc[res_gap['Máquina'] == maq, 'Peças Estoque - Ajuste'].sum() if maq in res_gap['Máquina'].values else 0
                meta_maq = metas_maq_hoje.get(maq, 0)
                gap = realizado - meta_maq
                
                color_class = "gap-positive" if gap >= 0 else "gap-negative"
                icon = "🔥" if gap >= 0 else "🔻"
                
                with cols_gap[i]:
                    st.markdown(f"""
                        <div class="metric-card" style="padding: 15px 10px;">
                            <div class="metric-title" style="color:#1e40af;">MÁQUINA {maq}</div>
                            <div class="metric-value" style="font-size:1.4rem;">{realizado:,.0f}</div>
                            <div style="font-size:0.75rem; color:#64748b; font-weight:600; margin-top:5px;">Meta: {meta_maq:,.0f}</div>
                            <div class="{color_class}" style="margin-top:8px;">{icon} {gap:,.0f}</div>
                        </div>
                    """, unsafe_allow_html=True)
        else:
            st.warning("Carregue o arquivo de DATAS (Metas) para visualizar o acompanhamento de Gap por Máquina.")

        # Detalhamento Histórico
        for dia in dias_sel:
            st.markdown(f"<div class='section-header'>Resultados Operacionais - {dia.strftime('%d/%m/%Y')}</div>", unsafe_allow_html=True)
            df_dia = df_order[df_order['Data'].dt.date == dia]
            res = df_dia.groupby(['Categoria', 'Máquina']).agg({'Run Time':'sum','Horário Padrão':'sum','Machine Counter':'sum','Peças Estoque - Ajuste':'sum'}).reset_index()
            res['Movimentação %'] = (res['Run Time'] / res['Horário Padrão'].replace(0,1) * 100).round(1)
            res['Perda %'] = ((res['Machine Counter'] - res['Peças Estoque - Ajuste']) / res['Machine Counter'].replace(0,1) * 100).round(1)
            st.dataframe(res[['Categoria','Máquina','Movimentação %','Perda %','Peças Estoque - Ajuste']].rename(columns={'Peças Estoque - Ajuste':'Estoque (pçs)'}), use_container_width=True, hide_index=True)

    # =========================================================
    # ABA 2: PERFORMANCE
    # =========================================================
    elif menu == "📈 PERFORMANCE":
        st.sidebar.subheader("Filtros do Dashboard")
        f_data = st.sidebar.date_input("Período Analisado", [df_order['Data'].min(), df_order['Data'].max()], key='p1')
        f_maq = st.sidebar.multiselect("Selecionar Máquinas", sorted(df_order['Máquina'].unique()), default=sorted(df_order['Máquina'].unique()))
        df_f = df_order[(df_order['Data'].dt.date >= f_data[0]) & (df_order['Data'].dt.date <= f_data[1]) & (df_order['Máquina'].isin(f_maq))]
        
        st.markdown("<div class='section-header'>Visão Geral de Performance</div>", unsafe_allow_html=True)
        st.markdown(f"""
            <div class="metric-container">
                <div class="metric-card"><div class="metric-title">Machine Counter Real</div><div class="metric-value">{df_f["Machine Counter"].sum():,.0f}</div></div>
                <div class="metric-card"><div class="metric-title">Peças para Estoque</div><div class="metric-value" style="color:#1d4ed8;">{df_f["Peças Estoque - Ajuste"].sum():,.0f}</div></div>
                <div class="metric-card"><div class="metric-title">Run Time Total (min)</div><div class="metric-value">{df_f["Run Time"].sum():,.0f}</div></div>
            </div>
        """, unsafe_allow_html=True)

        col1, col2 = st.columns(2)
        hp_sum = df_f['Horário Padrão'].sum()
        with col1: 
            fig1 = mini_gauge("Movimentação Total (%)", (df_f['Run Time'].sum()/hp_sum*100 if hp_sum>0 else 0), "#059669", 90)
            st.plotly_chart(fig1, use_container_width=True)
        with col2: 
            fig2 = mini_gauge("Loss Total (%)", ((df_f['Machine Counter'].sum()-df_f['Peças Estoque - Ajuste'].sum())/df_f['Machine Counter'].sum()*100 if df_f['Machine Counter'].sum()>0 else 0), "#e11d48", 2.5)
            st.plotly_chart(fig2, use_container_width=True)

    # =========================================================
    # ABA 3: TOP 10 PARADAS
    # =========================================================
    elif menu == "🛑 TOP 10 PARADAS":
        f_d_s = st.sidebar.date_input("Período das Paradas", [df_stops['Data'].min(), df_stops['Data'].max()], key='p2')
        df_s_f = df_stops[(df_stops['Data'].dt.date >= f_d_s[0]) & (df_stops['Data'].dt.date <= f_d_s[1])]
        
        st.markdown("<div class='section-header'>Pareto de Paradas Industriais</div>", unsafe_allow_html=True)
        
        col_g1, col_g2 = st.columns(2)
        with col_g1:
            st.markdown("<h4 style='color:#334155; text-align:center;'>Top 10: Impacto em Minutos</h4>", unsafe_allow_html=True)
            df_plot1 = df_s_f.groupby('Problema')['Minutos'].sum().sort_values().tail(10).reset_index()
            fig1 = px.bar(df_plot1, x='Minutos', y='Problema', orientation='h', text='Minutos', color_discrete_sequence=['#e11d48'])
            fig1.update_traces(textposition='outside', marker_border_radius=4)
            st.plotly_chart(apply_corporate_layout(fig1), use_container_width=True)

        with col_g2:
            st.markdown("<h4 style='color:#334155; text-align:center;'>Top 10: Frequência (Nº Ocorrências)</h4>", unsafe_allow_html=True)
            df_plot2 = df_s_f.groupby('Problema')['QTD'].sum().sort_values().tail(10).reset_index()
            fig2 = px.bar(df_plot2, x='QTD', y='Problema', orientation='h', text='QTD', color_discrete_sequence=['#1d4ed8'])
            fig2.update_traces(textposition='outside', marker_border_radius=4)
            st.plotly_chart(apply_corporate_layout(fig2), use_container_width=True)

    # =========================================================
    # ABA 4: CALENDÁRIO
    # =========================================================
    elif menu == "📅 CALENDÁRIO":
        mes_sel = st.sidebar.selectbox("Mês de Visualização", list(calendar.month_name)[1:], index=datetime.now().month-1)
        m_idx = list(calendar.month_name).index(mes_sel) + 1
        df_c = df_order[(df_order['Data'].dt.month == m_idx)]
        cal_data = df_c.groupby(df_c['Data'].dt.day).agg({'Run Time':'sum','Horário Padrão':'sum'}).reset_index()
        
        st.markdown(f"<div class='section-header'>Heatmap de Produção: {mes_sel.upper()}</div>", unsafe_allow_html=True)
        cols = st.columns(7)
        for i, d in enumerate(['SEG','TER','QUA','QUI','SEX','SÁB','DOM']): 
            cols[i].markdown(f"<div class='calendar-day-name'>{d}</div>", unsafe_allow_html=True)
        
        days = list(calendar.Calendar(0).itermonthdays(datetime.now().year, m_idx))
        html_grid = '<div class="calendar-grid">'
        for d in days:
            if d == 0: 
                html_grid += '<div style="background:transparent; border:none;"></div>'
            else:
                row = cal_data[cal_data['Data']==d]
                mov = (row['Run Time'].values[0]/row['Horário Padrão'].values[0]*100) if not row.empty and row['Horário Padrão'].values[0]>0 else 0
                
                # Gradiente de cor mais suave para o calendário
                if mov > 85:
                    cor_bg = "#dcfce7" # Verde bem claro
                    cor_txt = "#166534"
                elif mov > 0:
                    cor_bg = "#fee2e2" # Vermelho bem claro
                    cor_txt = "#991b1b"
                else:
                    cor_bg = "#ffffff"
                    cor_txt = "#94a3b8"
                    
                html_grid += f'<div class="day-card" style="background:{cor_bg}; border-color:{cor_txt}40"><span class="day-number" style="color:{cor_txt}">{d}</span><div class="day-status" style="color:{cor_txt}">{mov:.1f}%</div></div>'
        st.markdown(html_grid + '</div>', unsafe_allow_html=True)

    # =========================================================
    # ABA 5: ANÁLISE SEMANAL
    # =========================================================
    elif menu == "📋 ANÁLISE SEMANAL":
        st.sidebar.subheader("Filtros do Relatório")
        maq_b = st.sidebar.selectbox("Foco na Máquina", sorted(df_order['Máquina'].unique()))
        turno_b = st.sidebar.multiselect("Turnos", sorted(df_order['Turno'].unique()), default=sorted(df_order['Turno'].unique()))
        periodo_b = st.sidebar.date_input("Janela de Análise", [df_order['Data'].max() - timedelta(days=7), df_order['Data'].max()])
        
        df_b_all = df_order[(df_order['Data'].dt.date >= periodo_b[0]) & (df_order['Data'].dt.date <= periodo_b[1]) & (df_order['Turno'].isin(turno_b))]
        df_b = df_b_all[df_b_all['Máquina'] == maq_b]
        df_sb = df_stops[(df_stops['Data'].dt.date >= periodo_b[0]) & (df_stops['Data'].dt.date <= periodo_b[1]) & 
                         (df_stops['Máquina'] == maq_b) & (df_stops['Turno'].isin(turno_b))]

        str_turnos = ", ".join(turno_b) if turno_b else "Todos"
        st.markdown(f"""
            <div style="background: linear-gradient(135deg, #1e293b 0%, #0f172a 100%); border-radius: 12px; padding: 25px; margin-bottom: 20px; color: white; box-shadow: 0 4px 15px rgba(0,0,0,0.1);">
                <h2 style="margin:0; font-weight:900; color:white;">RELATÓRIO SEMANAL: MÁQUINA {maq_b}</h2>
                <div style="margin-top:10px; font-size:1rem; color:#cbd5e1;"><b>Turnos Analisados:</b> {str_turnos} &nbsp;|&nbsp; <b>Período:</b> {periodo_b[0].strftime('%d/%m')} a {periodo_b[1].strftime('%d/%m/%Y')}</div>
            </div>
        """, unsafe_allow_html=True)

        m_v = (df_b["Run Time"].sum()/df_b["Horário Padrão"].replace(0,1).sum()*100)
        l_v = ((df_b["Machine Counter"].sum()-df_b["Peças Estoque - Ajuste"].sum())/df_b["Machine Counter"].replace(0,1).sum()*100)

        v1, v2, v3 = st.columns([1, 1, 1])
        with v1: st.plotly_chart(mini_gauge("Eficiência (Mov)", m_v, "#059669", 85, 180), use_container_width=True)
        with v2: st.plotly_chart(mini_gauge("Perda (Loss)", l_v, "#e11d48", 5, 180), use_container_width=True)
        with v3: st.markdown(f'<div class="metric-card" style="height:180px; justify-content:center;"><div class="metric-title">Total Entregue</div><div class="metric-value" style="font-size:2.5rem; color:#1d4ed8;">{df_b["Peças Estoque - Ajuste"].sum():,.0f}</div></div>', unsafe_allow_html=True)

        st.markdown("<div class='section-header'>Benchmarking e Análise de Falhas</div>", unsafe_allow_html=True)
        col_rank, col_stops = st.columns([1, 2])
        
        with col_rank:
            rank_df = df_b_all.groupby('Máquina').agg({'Run Time':'sum','Horário Padrão':'sum'}).reset_index()
            rank_df['Mov %'] = (rank_df['Run Time'] / rank_df['Horário Padrão'].replace(0,1) * 100).round(1)
            rank_df = rank_df.sort_values('Mov %', ascending=False).reset_index(drop=True)
            rank_df.index += 1
            st.dataframe(rank_df[['Máquina', 'Mov %']].style.highlight_between(subset='Máquina', left=maq_b, right=maq_b, color='#dbeafe'), use_container_width=True)

        with col_stops:
            stop_imp = df_sb.groupby('Problema')['Minutos'].sum().sort_values(ascending=True).tail(5)
            if not stop_imp.empty:
                pior_parada = stop_imp.index[-1]
                df_p_plot = stop_imp.reset_index()
                df_p_plot['Label'] = df_p_plot.apply(lambda r: f"{r['Minutos']} min", axis=1)
                fig_b = px.bar(df_p_plot, x='Minutos', y='Problema', orientation='h', text='Label', color_discrete_sequence=['#fb923c']) # Cor Laranja Alerta
                fig_b.update_traces(textposition='outside', marker_border_radius=4)
                st.plotly_chart(apply_corporate_layout(fig_b), use_container_width=True)
            else: pior_parada = "Sem dados de parada"

        st.markdown(f"""
            <div class="five-why-box">
                <h4 style="color:#0f172a; margin-top:0;">Investigação de Causa Raiz (5 Porquês): <span style="color:#e11d48;">{pior_parada}</span></h4>
                <div class="five-why-line"><b>1º Por quê?</b> </div>
                <div class="five-why-line"><b>2º Por quê?</b> </div>
                <div class="five-why-line"><b>3º Por quê?</b> </div>
                <div class="five-why-line"><b>4º Por quê?</b> </div>
                <div class="five-why-line"><b>5º Por quê?</b> </div>
                <div style="display:flex; gap:15px; margin-top:20px;">
                    <div style="flex:1; background:#f8fafc; border:1px solid #cbd5e1; padding:15px; border-radius:8px;"><b>CAUSA RAIZ ENCONTRADA:</b></div>
                    <div style="flex:2; background:#f0fdf4; border:1px solid #bbf7d0; padding:15px; border-radius:8px;"><b>AÇÃO CORRETIVA / PREVENTIVA:</b></div>
                </div>
            </div>
        """, unsafe_allow_html=True)
else:
    st.info("💡 Carregue a Base de Produção (.xlsm) e a Base de Metas (.xlsx) no menu lateral para iniciar.")
import streamlit as st
import pandas as pd
import plotly.graph_objects as go
import plotly.express as px
import calendar
from datetime import datetime, timedelta

# 1. CONFIGURAÇÃO DA PÁGINA
st.set_page_config(layout="wide", page_title="Industrial Analytics Hub", page_icon="⚙️")

# --- ESTILIZAÇÃO CSS CORPORATIVA ---
st.markdown("""
    <style>
    @import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;600;800;900&display=swap');
    * { font-family: 'Inter', sans-serif; }
    .stApp { background-color: #f8fafc; color: #0f172a; }
    [data-testid="stSidebar"] { background-color: #ffffff; border-right: 1px solid #e2e8f0; }
    [data-testid="stSidebar"] .stRadio div[role="radiogroup"] > label {
        background-color: #ffffff; border: 1px solid transparent;
        padding: 10px 15px !important; border-radius: 8px !important;
        margin-bottom: 5px !important; color: #475569 !important; font-weight: 600; cursor: pointer;
    }
    [data-testid="stSidebar"] .stRadio div[role="radiogroup"] > label[data-checked="true"] {
        background: linear-gradient(135deg, #1d4ed8 0%, #1e40af 100%) !important;
        color: white !important; box-shadow: 0 4px 6px -1px rgba(29, 78, 216, 0.2);
    }
    .metric-container { display: flex; justify-content: space-between; gap: 12px; margin-bottom: 20px; }
    .metric-card {
        background: #ffffff; padding: 15px; border-radius: 10px;
        text-align: center; border: 1px solid #e2e8f0;
        flex: 1; min-height: 80px; display: flex; flex-direction: column; justify-content: center;
        box-shadow: 0 2px 4px rgba(0,0,0,0.02);
    }
    .metric-title { color: #64748b; font-size: 0.7rem; font-weight: 800; text-transform: uppercase; margin-bottom: 5px; }
    .metric-value { color: #0f172a; font-size: 1.5rem; font-weight: 900; line-height: 1.1; }
    .gap-text { font-size: 0.75rem; font-weight: 700; margin-top: 5px; }
    .color-red { color: #e11d48; }
    .color-green { color: #059669; }
    .section-header {
        background: #f1f5f9; padding: 10px 15px; border-radius: 6px;
        color: #1e40af; font-weight: 900; text-transform: uppercase;
        margin-top: 25px; margin-bottom: 15px; border-left: 5px solid #1d4ed8; font-size: 1rem;
    }
    .calendar-day-name { text-align: center; font-weight: 800; color: #475569; font-size: 0.85rem; padding-bottom: 5px; }
    .calendar-grid { display: grid; grid-template-columns: repeat(7, 1fr); gap: 8px; width: 100%; }
    .day-card { background: #ffffff; border-radius: 6px; padding: 10px; min-height: 85px; border: 1px solid #e2e8f0; }
    .day-number { font-size: 1rem; font-weight: 900; color: #1e293b; }
    .day-status { font-size: 0.75rem; font-weight: 700; margin-top: 5px; text-align: right; }
    .five-why-box { border: 2px solid #cbd5e1; padding: 20px; background: #ffffff; border-radius: 10px; margin-top: 15px; }
    .five-why-line { border-bottom: 1px dashed #cbd5e1; padding: 10px 0; font-size: 0.9rem; color: #000; }
    h1, h2, h3, p, span, label { color: #000000 !important; }
    </style>
    """, unsafe_allow_html=True)

# 2. CARREGAMENTO E LIMPEZA COM DIAGNÓSTICO
@st.cache_data
def load_data(file_obj):
    try:
        df_order = pd.read_excel(file_obj, sheet_name="Result by order")
        df_stops = pd.read_excel(file_obj, sheet_name="Stop machine item")
    except Exception as e:
        st.error(f"❌ Erro ao ler abas do Excel. Certifique-se de que as abas 'Result by order' e 'Stop machine item' existem. Erro: {e}")
        st.stop()

    # Limpeza de espaços em branco nos nomes das colunas (MUITO IMPORTANTE)
    df_order.columns = df_order.columns.str.strip()
    df_stops.columns = df_stops.columns.str.strip()
    
    # DIAGNÓSTICO DE COLUNAS FALTANTES
    colunas_necessarias_order = ['Data', 'Máquina', 'Turno', 'Run Time', 'Horário Padrão', 'Machine Counter', 'Peças em Estoque']
    faltam_order = [col for col in colunas_necessarias_order if col not in df_order.columns]
    
    if faltam_order:
        st.error(f"❌ ERRO: Faltam colunas na aba 'Result by order'.")
        st.warning(f"O código procurou por: {faltam_order}")
        st.info(f"Colunas que o Pandas encontrou no seu arquivo: {list(df_order.columns)}")
        st.stop()

    colunas_necessarias_stops = ['Data', 'Máquina', 'Turno', 'Minutos', 'QTD', 'Problema']
    faltam_stops = [col for col in colunas_necessarias_stops if col not in df_stops.columns]
    
    if faltam_stops:
        st.error(f"❌ ERRO: Faltam colunas na aba 'Stop machine item'.")
        st.warning(f"O código procurou por: {faltam_stops}")
        st.info(f"Colunas que o Pandas encontrou no seu arquivo: {list(df_stops.columns)}")
        st.stop()

    # Conversão de numéricos
    for col in ['Run Time', 'Horário Padrão', 'Machine Counter', 'Peças em Estoque']:
        df_order[col] = pd.to_numeric(df_order[col], errors='coerce').fillna(0)
        
    for col in ['Minutos', 'QTD']:
        df_stops[col] = pd.to_numeric(df_stops[col], errors='coerce').fillna(0)

    # Tratamento de Datas e Textos
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
        if not target: return 0, 0, {}

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
                    valor_geral = pd.to_numeric(row_meta_geral[col_idx], errors='coerce')
                    if pd.notna(valor_geral):
                        meta_geral_mes += valor_geral
                        if d_val.date() <= data_ref:
                            meta_ate_hoje += valor_geral
                
                if d_val.date() == data_ref:
                    col_idx_hoje = col_idx

        if col_idx_hoje is not None:
            for maq, row_idx in mapping_maquinas.items():
                val_maq = pd.to_numeric(df_raw.iloc[row_idx, col_idx_hoje], errors='coerce')
                metas_maq_hoje[maq] = val_maq if pd.notna(val_maq) else 0
                
        return meta_geral_mes, meta_ate_hoje, metas_maq_hoje
    except Exception as e:
        return 0, 0, {}

def mini_gauge(label, value, color, target, height=180):
    fig = go.Figure(go.Indicator(
        mode="gauge+number", value=value,
        number={'suffix': "%", 'font': {'size': 20, 'color': '#0f172a'}},
        title={'text': label, 'font': {'size': 14, 'color': '#64748b'}},
        gauge={
            'axis': {'range': [0, 100], 'tickwidth': 1, 'tickcolor': "black"},
            'bar': {'color': color},
            'bgcolor': "#e2e8f0",
            'threshold': {'line': {'color': "#000000", 'width': 3}, 'value': target}
        }
    ))
    fig.update_layout(height=height, margin=dict(l=10, r=10, t=30, b=10), paper_bgcolor='rgba(0,0,0,0)')
    return fig

# --- SIDEBAR ---
with st.sidebar:
    st.markdown("<h1 style='color:#1d4ed8; font-weight:900;'>🏭 ANALYTICS HUB</h1>", unsafe_allow_html=True)
    uploaded_file = st.file_uploader("📂 Excel Produção (.xlsm)", type=["xlsm", "xlsx"])
    up_datas = st.file_uploader("📂 Excel DATAS Metas (.xlsx)", type=["xlsx"])
    st.markdown("---")
    if uploaded_file:
        menu = st.radio("NAVEGAÇÃO", ["📋 REPORTE DIÁRIO", "📈 PERFORMANCE", "🛑 TOP 10 PARADAS", "📅 CALENDÁRIO", "📋 ANÁLISE SEMANAL"])

if uploaded_file:
    df_order, df_stops = load_data(uploaded_file)

    if df_order.empty:
        st.error("Nenhuma data válida encontrada no arquivo de produção. Verifique o Excel.")
        st.stop()

    # =========================================================
    # ABA 1: REPORTE DIÁRIO
    # =========================================================
    if menu == "📋 REPORTE DIÁRIO":
        st.subheader("⚙️ Filtros da Página")
        col_f1, col_f2 = st.columns(2)
        with col_f1:
            data_ref_reporte = st.date_input("Data de Referência", df_order['Data'].max().date())
        with col_f2:
            datas_disp = sorted(df_order['Data'].dt.date.unique().tolist(), reverse=True)
            dias_sel = st.multiselect("Filtrar histórico (Tabelas):", datas_disp, default=datas_disp[:3] if len(datas_disp) >= 3 else datas_disp)

        st.markdown(f"## 📋 Reporte Diário de Produção - {data_ref_reporte.strftime('%d/%m/%Y')}")

        df_acumulado_mes = df_order[(df_order['Data'].dt.month == data_ref_reporte.month) & (df_order['Data'].dt.year == data_ref_reporte.year) & (df_order['Data'].dt.date <= data_ref_reporte)]
        
        estoque_acum_mes = df_acumulado_mes['Peças em Estoque'].sum()
        total_mc_mes = df_acumulado_mes['Machine Counter'].sum()
        mov_acum_mes = (df_acumulado_mes['Run Time'].sum() / df_acumulado_mes['Horário Padrão'].sum() * 100) if df_acumulado_mes['Horário Padrão'].sum() > 0 else 0
        loss_acum_mes = ((total_mc_mes - estoque_acum_mes) / total_mc_mes * 100) if total_mc_mes > 0 else 0
        
        meta_mov, meta_loss = 90.0, 2.5
        gap_mov = mov_acum_mes - meta_mov
        gap_loss = loss_acum_mes - meta_loss
        
        meta_geral_mes, meta_dinamica_hoje, metas_maq_hoje = load_metas_completas(up_datas, data_ref_reporte) if up_datas else (0, 0, {})

        st.markdown(f"""
            <div class="metric-container">
                <div class="metric-card">
                    <div class="metric-title">Movimentação Mês (Meta 90%)</div>
                    <div class="metric-value">{mov_acum_mes:.1f}%</div>
                    <div class="gap-text {'color-green' if gap_mov>=0 else 'color-red'}">{gap_mov:+.1f}% vs meta</div>
                </div>
                <div class="metric-card">
                    <div class="metric-title">Loss Mês (Meta 2,5%)</div>
                    <div class="metric-value color-red">{loss_acum_mes:.1f}%</div>
                    <div class="gap-text {'color-green' if gap_loss<=0 else 'color-red'}">{gap_loss:+.1f}% vs meta</div>
                </div>
                <div class="metric-card">
                    <div class="metric-title">Estoque Realizado Mês</div>
                    <div class="metric-value">{estoque_acum_mes:,.0f}</div>
                </div>
            </div>
            <div class="metric-container">
                <div class="metric-card"><div class="metric-title">Meta Acumulada (Até {data_ref_reporte.strftime('%d/%m')})</div><div class="metric-value" style="color:#1d4ed8">{meta_dinamica_hoje:,.0f}</div></div>
                <div class="metric-card"><div class="metric-title">Meta Geral do Mês (Datas)</div><div class="metric-value" style="color:#059669">{meta_geral_mes:,.0f}</div></div>
            </div>
        """, unsafe_allow_html=True)

        if up_datas:
            st.markdown(f"<div class='section-header'>🎯 ANÁLISE DE PEÇAS FALTANTES (GAP) - {data_ref_reporte.strftime('%d/%m/%Y')}</div>", unsafe_allow_html=True)
            df_dia_gap = df_order[df_order['Data'].dt.date == data_ref_reporte]
            res_gap = df_dia_gap.groupby('Máquina').agg({'Peças em Estoque':'sum'}).reset_index()
            
            gap_data = []
            maquinas_alvo = ["1", "2", "3", "4", "5", "6", "7"]
            for maq in maquinas_alvo:
                realizado = res_gap.loc[res_gap['Máquina'] == maq, 'Peças em Estoque'].sum() if maq in res_gap['Máquina'].values else 0
                meta_maq = metas_maq_hoje.get(maq, 0)
                faltante = meta_maq - realizado
                status = "✅ Bateu Meta" if faltante <= 0 else f"🚨 Faltam {faltante:,.0f} peças"
                gap_data.append({"Máquina": f"MQ{maq}", "Meta": meta_maq, "Realizado": realizado, "Status (Faltante)": status})
            
            st.table(pd.DataFrame(gap_data))
        else:
            st.warning("⚠️ Carregue o arquivo Excel DATAS na barra lateral para calcular as peças faltantes por máquina.")

        for dia in dias_sel:
            st.markdown(f"<div class='section-header'>DETALHAMENTO POR MÁQUINA - {dia.strftime('%d/%m/%Y')}</div>", unsafe_allow_html=True)
            df_dia = df_order[df_order['Data'].dt.date == dia]
            res = df_dia.groupby(['Categoria', 'Máquina']).agg({'Run Time':'sum','Horário Padrão':'sum','Machine Counter':'sum','Peças em Estoque':'sum'}).reset_index()
            res['Movimentação %'] = (res['Run Time'] / res['Horário Padrão'].replace(0,1) * 100).round(1)
            res['Perda %'] = ((res['Machine Counter'] - res['Peças em Estoque']) / res['Machine Counter'].replace(0,1) * 100).round(1)
            st.table(res[['Categoria','Máquina','Movimentação %','Perda %','Peças em Estoque']].rename(columns={'Peças em Estoque':'Qtd Estoque'}))

    # =========================================================
    # ABA 2: PERFORMANCE
    # =========================================================
    elif menu == "📈 PERFORMANCE":
        st.sidebar.subheader("Filtros")
        f_data = st.sidebar.date_input("Período", [df_order['Data'].min(), df_order['Data'].max()], key='p1')
        
        if len(f_data) == 2:
            data_ini, data_fim = f_data[0], f_data[1]
        else:
            data_ini = data_fim = f_data[0]

        f_maq = st.sidebar.multiselect("Máquinas", sorted(df_order['Máquina'].unique()), default=sorted(df_order['Máquina'].unique()), key='m1')
        f_turno = st.sidebar.multiselect("Turnos", sorted(df_order['Turno'].unique()), default=sorted(df_order['Turno'].unique()), key='t1')
        
        df_f = df_order[(df_order['Data'].dt.date >= data_ini) & (df_order['Data'].dt.date <= data_fim) & (df_order['Máquina'].isin(f_maq)) & (df_order['Turno'].isin(f_turno))]
        
        st.markdown(f"""
            <div class="metric-container">
                <div class="metric-card"><div class="metric-title">Machine Counter</div><div class="metric-value">{df_f["Machine Counter"].sum():,.0f}</div></div>
                <div class="metric-card"><div class="metric-title">Peças Estoque</div><div class="metric-value" style="color:#1d4ed8;">{df_f["Peças em Estoque"].sum():,.0f}</div></div>
                <div class="metric-card"><div class="metric-title">Run Time Total</div><div class="metric-value">{df_f["Run Time"].sum():,.0f}m</div></div>
            </div>
        """, unsafe_allow_html=True)

        col1, col2 = st.columns(2)
        hp_sum = df_f['Horário Padrão'].sum()
        with col1: st.plotly_chart(mini_gauge("Movimentação (%)", (df_f['Run Time'].sum()/hp_sum*100 if hp_sum>0 else 0), "#059669", 90), use_container_width=True)
        with col2: st.plotly_chart(mini_gauge("Loss (%)", ((df_f['Machine Counter'].sum()-df_f['Peças em Estoque'].sum())/df_f['Machine Counter'].sum()*100 if df_f['Machine Counter'].sum()>0 else 0), "#e11d48", 2.5), use_container_width=True)

    # =========================================================
    # ABA 3: TOP 10 PARADAS
    # =========================================================
    elif menu == "🛑 TOP 10 PARADAS":
        f_d_s = st.sidebar.date_input("Período", [df_stops['Data'].min(), df_stops['Data'].max()], key='p2')
        
        if len(f_d_s) == 2:
            d_ini, d_fim = f_d_s[0], f_d_s[1]
        else:
            d_ini = d_fim = f_d_s[0]

        df_s_f = df_stops[(df_stops['Data'].dt.date >= d_ini) & (df_stops['Data'].dt.date <= d_fim)]
        
        fig1 = px.bar(df_s_f.groupby('Problema')['Minutos'].sum().sort_values().tail(10), orientation='h', title="Minutos Totais", color_discrete_sequence=['#e11d48'])
        fig1.update_layout(paper_bgcolor='rgba(0,0,0,0)', plot_bgcolor='rgba(0,0,0,0)', font={'color':'black'})
        st.plotly_chart(fig1, use_container_width=True)

        fig2 = px.bar(df_s_f.groupby('Problema')['QTD'].sum().sort_values().tail(10), orientation='h', title="Frequência (Qtd)", color_discrete_sequence=['#1d4ed8'])
        fig2.update_layout(paper_bgcolor='rgba(0,0,0,0)', plot_bgcolor='rgba(0,0,0,0)', font={'color':'black'})
        st.plotly_chart(fig2, use_container_width=True)

    # =========================================================
    # ABA 4: CALENDÁRIO
    # =========================================================
    elif menu == "📅 CALENDÁRIO":
        meses_pt = ["Janeiro", "Fevereiro", "Março", "Abril", "Maio", "Junho", "Julho", "Agosto", "Setembro", "Outubro", "Novembro", "Dezembro"]
        
        anos_disponiveis = sorted(df_order['Data'].dt.year.dropna().unique().tolist(), reverse=True)
        ano_analise = st.sidebar.selectbox("Ano", anos_disponiveis)
        
        df_ano = df_order[df_order['Data'].dt.year == ano_analise]
        mes_padrao = df_ano['Data'].dt.month.max() - 1 if not df_ano.empty else 0
        
        mes_sel = st.sidebar.selectbox("Mês", meses_pt, index=int(mes_padrao))
        m_idx = meses_pt.index(mes_sel) + 1
        
        df_c = df_ano[df_ano['Data'].dt.month == m_idx]
        
        df_c_copia = df_c.copy()
        df_c_copia['Dia'] = df_c_copia['Data'].dt.day
        cal_data = df_c_copia.groupby('Dia').agg({'Run Time':'sum','Horário Padrão':'sum'}).reset_index()
        
        st.markdown(f"### 📅 Cronograma {mes_sel} {ano_analise}")
        cols = st.columns(7)
        for i, d in enumerate(['Segunda','Terça','Quarta','Quinta','Sexta','Sábado','Domingo']): cols[i].markdown(f"<div class='calendar-day-name'>{d}</div>", unsafe_allow_html=True)
        
        days = list(calendar.Calendar(0).itermonthdays(ano_analise, m_idx))
        html_grid = '<div class="calendar-grid">'
        for d in days:
            if d == 0: 
                html_grid += '<div></div>'
            else:
                row = cal_data[cal_data['Dia'] == d]
                mov = 0.0
                if not row.empty:
                    hp = row['Horário Padrão'].values[0]
                    rt = row['Run Time'].values[0]
                    if hp > 0:
                        mov = (rt / hp) * 100

                cor = "#dcfce7" if mov >= 85 else "#fee2e2" if mov > 0 else "#f1f5f9"
                txt_cor = "#166534" if mov >= 85 else "#991b1b" if mov > 0 else "#94a3b8"
                html_grid += f'<div class="day-card" style="background:{cor};"><span class="day-number" style="color:{txt_cor}">{d}</span><div class="day-status" style="color:{txt_cor}">{mov:.1f}%</div></div>'
        
        st.markdown(html_grid + '</div>', unsafe_allow_html=True)

    # =========================================================
    # ABA 5: ANÁLISE SEMANAL
    # =========================================================
    elif menu == "📋 ANÁLISE SEMANAL":
        st.sidebar.subheader("Filtros Board")
        maq_b = st.sidebar.selectbox("Máquina", sorted(df_order['Máquina'].unique()))
        turno_b = st.sidebar.multiselect("Turnos", sorted(df_order['Turno'].unique()), default=sorted(df_order['Turno'].unique()), key='tb')
        
        periodo_b = st.sidebar.date_input("Período", [df_order['Data'].max() - timedelta(days=7), df_order['Data'].max()])
        if len(periodo_b) == 2:
            p_ini, p_fim = periodo_b[0], periodo_b[1]
        else:
            p_ini = p_fim = periodo_b[0]
        
        df_b_all = df_order[(df_order['Data'].dt.date >= p_ini) & (df_order['Data'].dt.date <= p_fim) & (df_order['Turno'].isin(turno_b))]
        df_b = df_b_all[df_b_all['Máquina'] == maq_b]
        df_sb = df_stops[(df_stops['Data'].dt.date >= p_ini) & (df_stops['Data'].dt.date <= p_fim) & 
                         (df_stops['Máquina'] == maq_b) & (df_stops['Turno'].isin(turno_b))]

        str_turnos = ", ".join(turno_b) if turno_b else "Nenhum"
        st.markdown(f"""<div style="text-align:center; border-bottom:3px solid #1d4ed8; padding-bottom:10px; margin-bottom:15px;">
            <h1 style="color:#0f172a; margin:0;">RELATÓRIO SEMANAL DE PERFORMANCE - MÁQUINA {maq_b}</h1>
            <h3 style="color:#1d4ed8; margin:0;">TURNO(S): {str_turnos}</h3>
            <p style="color:#64748b; font-size:1rem;">Período: {p_ini.strftime('%d/%m')} a {p_fim.strftime('%d/%m/%Y')}</p></div>""", unsafe_allow_html=True)

        m_v = (df_b["Run Time"].sum()/df_b["Horário Padrão"].replace(0,1).sum()*100)
        l_v = ((df_b["Machine Counter"].sum()-df_b["Peças em Estoque"].sum())/df_b["Machine Counter"].replace(0,1).sum()*100)
        pecas_v = df_b["Peças em Estoque"].sum()

        v1, v2, v3 = st.columns([1, 1, 1])
        with v1: st.plotly_chart(mini_gauge("Movimentação", m_v, "#059669", 85, 180), use_container_width=True)
        with v2: st.plotly_chart(mini_gauge("Loss", l_v, "#e11d48", 5, 180), use_container_width=True)
        with v3: st.markdown(f'<div class="metric-card" style="height:150px;"><div class="metric-title">Peças Enviadas</div><div class="metric-value" style="font-size:2.2rem; color:#1d4ed8;">{pecas_v:,.0f}</div></div>', unsafe_allow_html=True)

        rank_df = df_b_all.groupby('Máquina').agg({'Run Time':'sum','Horário Padrão':'sum'}).reset_index()
        rank_df['Mov %'] = (rank_df['Run Time'] / rank_df['Horário Padrão'].replace(0,1) * 100).round(1)
        rank_df = rank_df.sort_values('Mov %', ascending=False).reset_index(drop=True)
        rank_df.index += 1
        
        check_maq = rank_df[rank_df['Máquina'] == maq_b]
        if not check_maq.empty:
            posicao = check_maq.index[0]
            total_maqs = len(rank_df)
            if posicao <= 2:
                msg, cor_msg, bg_msg = f"🌟 EXCELENTE! Máquina {maq_b} no TOP 2 ({posicao}º).", "#064e3b", "#dcfce7"
            elif posicao > (total_maqs - 2):
                msg, cor_msg, bg_msg = f"🚨 ATENÇÃO! Máquina {maq_b} na posição {posicao}º. Foco total!", "#9f1239", "#ffe4e6"
            else:
                msg, cor_msg, bg_msg = f"📈 BOM TRABALHO! Máquina {maq_b} na posição {posicao}º.", "#1e40af", "#dbeafe"
            st.markdown(f'<div style="background:{bg_msg}; color:{cor_msg}; padding: 12px; border-radius: 8px; margin-bottom: 10px; text-align: center; font-weight: 700;">{msg}</div>', unsafe_allow_html=True)

        col_rank, col_stops = st.columns([1, 2])
        with col_rank:
            st.markdown("🏆 **Ranking Mov. (%)**")
            st.table(rank_df[['Máquina', 'Mov %']])

        with col_stops:
            st.markdown("🛑 **Impacto das Paradas (%)**")
            stop_imp = df_sb.groupby('Problema')['Minutos'].sum().sort_values(ascending=True).tail(5)
            if not stop_imp.empty:
                pior_parada = stop_imp.index[-1]
                total_min_p = df_sb['Minutos'].sum()
                df_p_plot = stop_imp.reset_index()
                df_p_plot['%'] = (df_p_plot['Minutos'] / total_min_p * 100).round(1)
                df_p_plot['Label'] = df_p_plot.apply(lambda r: f"{r['Minutos']} min ({r['%']}%)", axis=1)
                fig_b = px.bar(df_p_plot, x='Minutos', y='Problema', orientation='h', text='Label', color_discrete_sequence=['#e11d48'])
                fig_b.update_layout(height=250, margin=dict(l=0,r=0,t=0,b=0), paper_bgcolor='rgba(0,0,0,0)', font={'color':'black'})
                st.plotly_chart(fig_b, use_container_width=True)
            else: pior_parada = "Nenhuma parada registrada"

        st.markdown(f"""
            <div class="five-why-box">
                <h4 style="color:#0f172a; margin-top:0;">ANÁLISE DE CAUSA RAIZ - 5 PORQUÊS: <span style="color:#e11d48;">{pior_parada}</span></h4>
                <div class="five-why-line"><b>1º Por que?</b> _________________________________________________________________</div>
                <div class="five-why-line"><b>2º Por que?</b> _________________________________________________________________</div>
                <div class="five-why-line"><b>3º Por que?</b> _________________________________________________________________</div>
                <div class="five-why-line"><b>4º Por que?</b> _________________________________________________________________</div>
                <div class="five-why-line"><b>5º Por que?</b> _________________________________________________________________</div>
                <br>
                <div style="display:flex; gap:10px;">
                    <div style="flex:1; border:1px solid #cbd5e1; background:#f8fafc; padding:10px; min-height:80px;"><b>CAUSA RAIZ:</b></div>
                    <div style="flex:2; border:1px solid #cbd5e1; background:#f8fafc; padding:10px; min-height:80px;"><b>AÇÃO CORRETIVA:</b></div>
                </div>
            </div>
        """, unsafe_allow_html=True)
else:
    st.info("💡 Carregue o arquivo Excel para iniciar.")

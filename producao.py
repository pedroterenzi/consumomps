import streamlit as st
import pandas as pd
import calendar
from datetime import datetime
import re

# 1. CONFIGURAÇÃO DA PÁGINA
st.set_page_config(layout="wide", page_title="Betting Analytics Pro", page_icon="💹")

# --- ESTILIZAÇÃO CSS PREMIUM (ADAPTADA PARA MOBILE) ---
st.markdown("""
    <style>
    @import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;600;800&display=swap');
    * { font-family: 'Inter', sans-serif; }
    .stApp { background-color: #0f172a; }

    /* Grid do Calendário Responsivo */
    .calendar-grid { 
        display: grid; 
        grid-template-columns: repeat(7, 1fr); 
        gap: 10px; 
        margin-top: 20px; 
    }
    
    @media (max-width: 768px) {
        .calendar-grid { grid-template-columns: repeat(1, 1fr); }
        .metric-card { margin-bottom: 10px; }
    }

    .day-name { text-align: center; color: #94a3b8; font-weight: 800; font-size: 0.8rem; text-transform: uppercase; }
    .day-card { background: #1e293b; border-radius: 16px; padding: 15px; min-height: 100px; display: flex; flex-direction: column; justify-content: space-between; border: 1px solid rgba(255, 255, 255, 0.05); }
    .day-number { font-size: 1.1rem; font-weight: 800; color: #f8fafc; }
    .day-off { color: #64748b; font-size: 0.7rem; font-weight: 600; text-align: right; }
    .day-value { font-size: 0.9rem; font-weight: 700; margin-top: 5px; }
    
    .green-card { background: linear-gradient(135deg, #059669 0%, #064e3b 100%); border: none; }
    .red-card { background: linear-gradient(135deg, #dc2626 0%, #7f1d1d 100%); border: none; }
    .empty-card { display: none; }

    /* Performance Cards */
    .perf-container { display: flex; flex-direction: column; gap: 12px; margin-top: 20px; }
    .perf-card { background: #1e293b; border-radius: 12px; padding: 15px 25px; display: flex; align-items: center; justify-content: space-between; border: 1px solid rgba(255, 255, 255, 0.05); }
    .perf-name { color: #f8fafc; font-weight: 700; font-size: 1rem; }
    .perf-meta { color: #94a3b8; font-size: 0.8rem; }
    .val-green { color: #10b981; font-weight: 800; }
    .val-red { color: #ef4444; font-weight: 800; }
    .roi-bar-bg { background: #334155; border-radius: 10px; height: 8px; width: 100%; overflow: hidden; margin-top: 5px; }
    .roi-bar-fill { height: 100%; border-radius: 10px; }

    /* Métricas Topo */
    .metric-card { padding: 20px; border-radius: 20px; text-align: center; color: white; font-weight: 800; box-shadow: 0 10px 15px -3px rgba(0, 0, 0, 0.3); }
    </style>
    """, unsafe_allow_html=True)

st.title("💹 Betting Analytics Executive")

# --- 2. BARRA LATERAL (AJUSTADA PARA NUVEM) ---
st.sidebar.header("🕹️ Painel Online")

# Nome do arquivo que você vai subir no GitHub
NOME_ARQUIVO_PADRAO = "AccountStatement_ (1).csv"
arquivo_path = st.sidebar.text_input("Nome do Arquivo CSV", NOME_ARQUIVO_PADRAO)
stake_padrao = st.sidebar.number_input("Stake (R$)", value=600.0)

def clean_money(val):
    if val == '--' or pd.isna(val): return 0.0
    return float(str(val).replace(',', ''))

try:
    # Lê o arquivo da pasta atual
    df_raw = pd.read_csv(arquivo_path)
    
    # Processamento de colunas
    if 'Descrição' in df_raw.columns:
        df_raw = df_raw.rename(columns={'Descrição': 'Evento'})
        df_raw['Valor (R$)'] = df_raw['Entrada de Dinheiro (R$)'].apply(clean_money) + df_raw['Saída de Dinheiro (R$)'].apply(clean_money)
    
    df = df_raw[~df_raw['Evento'].str.contains('Depósito|Deposit|Withdraw|Saque|Transferência', case=False, na=False)].copy()
    
    meses_pt = {'jan': 'Jan', 'fev': 'Feb', 'mar': 'Mar', 'abr': 'Apr', 'mai': 'May', 'jun': 'Jun', 
                'jul': 'Jul', 'ago': 'Aug', 'set': 'Sep', 'out': 'Oct', 'nov': 'Nov', 'dez': 'Dec'}
    for pt, en in meses_pt.items():
        df['Data'] = df['Data'].str.replace(pt, en, case=False)
    
    df['Data'] = pd.to_datetime(df['Data'])
    df['Data_Apenas'] = df['Data'].dt.date

    # Filtro de Data
    data_sel = st.sidebar.date_input("Período", [df['Data_Apenas'].min(), df['Data_Apenas'].max()])

    if len(data_sel) == 2:
        start, end = data_sel
        df_f = df[(df['Data_Apenas'] >= start) & (df['Data_Apenas'] <= end)].copy()
        
        def extract_id(row):
            match = re.search(r'Ref: (\d+)', str(row['Evento']))
            return match.group(1) if match else row.name
            
        df_f['ID_Ref'] = df_f.apply(extract_id, axis=1)
        df_clean = df_f.groupby(['ID_Ref', 'Data_Apenas', 'Evento']).agg({'Valor (R$)': 'sum'}).reset_index()

        # --- 3. MÉTRICAS TOPO ---
        total_l = df_clean['Valor (R$)'].sum()
        total_s = total_l / stake_padrao
        
        c1, c2, c3 = st.columns(3)
        with c1:
            bg = "#10b981" if total_l >= 0 else "#ef4444"
            st.markdown(f'<div class="metric-card" style="background:{bg}">Lucro Líquido<br><span style="font-size:26px">R$ {total_l:,.2f}</span></div>', unsafe_allow_html=True)
        with c2:
            bg = "#10b981" if total_s >= 0 else "#ef4444"
            st.markdown(f'<div class="metric-card" style="background:{bg}">Saldo Stakes<br><span style="font-size:26px">{total_s:,.2f} STK</span></div>', unsafe_allow_html=True)
        with c3:
            st.markdown(f'<div class="metric-card" style="background:#1e293b; border: 1px solid #334155">Entradas Total<br><span style="font-size:26px">{len(df_clean)}</span></div>', unsafe_allow_html=True)

        # --- 4. CALENDÁRIO ---
        st.subheader("📅 Diário de Bordo")
        ano, mes = start.year, start.month
        cal = calendar.Calendar(firstweekday=0)
        dias_mes = list(cal.itermonthdays(ano, mes))
        lucro_dia = df_clean.groupby(pd.to_datetime(df_clean['Data_Apenas']).dt.day)['Valor (R$)'].sum()

        cols_names = ['Seg', 'Ter', 'Qua', 'Qui', 'Sex', 'Sáb', 'Dom']
        html_cal = '<div class="calendar-grid">'
        for name in cols_names: html_cal += f'<div class="day-name">{name}</div>'
        for dia in dias_mes:
            if dia == 0: html_cal += '<div class="empty-card"></div>'
            else:
                valor = lucro_dia.get(dia, None)
                card_class = "day-card"
                if valor is None: content = f'<div class="day-number">{dia}</div><div class="day-off">OFF</div>'
                elif valor > 0.05:
                    card_class += " green-card"
                    content = f'<div class="day-number">{dia}</div><div class="day-value">+R$ {valor:,.2f}</div>'
                elif valor < -0.05:
                    card_class += " red-card"
                    content = f'<div class="day-number">{dia}</div><div class="day-value">-R$ {abs(valor):,.2f}</div>'
                else: content = f'<div class="day-number">{dia}</div><div class="day-value">0.00</div>'
                html_cal += f'<div class="{card_class}">{content}</div>'
        html_cal += '</div>'
        st.markdown(html_cal, unsafe_allow_html=True)

        st.markdown("<br><br>", unsafe_allow_html=True)

        # --- 5. PERFORMANCE POR ESTRATÉGIA ---
        st.subheader("🎯 Performance por Estratégia")
        def extrair_est(txt):
            txt = str(txt).split('Ref:')[0]
            return txt.split('/')[-1].strip() if '/' in txt else "Match Odds / Outros"

        df_clean['Estrategia'] = df_clean['Evento'].apply(extrair_est)
        resumo = df_clean.groupby('Estrategia').agg({'Valor (R$)': 'sum', 'ID_Ref': 'count'}).rename(columns={'ID_Ref': 'Entradas', 'Valor (R$)': 'Lucro'})
        resumo['ROI'] = (resumo['Lucro'] / (resumo['Entradas'] * stake_padrao)) * 100
        resumo = resumo.sort_values('Lucro', ascending=False)

        html_perf = '<div class="perf-container">'
        for est, row in resumo.iterrows():
            c_val = "val-green" if row['Lucro'] >= 0 else "val-red"
            c_roi = "#10b981" if row['ROI'] >= 0 else "#ef4444"
            w_roi = max(min(abs(row['ROI']), 100), 2)
            html_perf += f'''
            <div class="perf-card">
                <div class="perf-info">
                    <div class="perf-name">{est}</div>
                    <div class="perf-meta">{int(row['Entradas'])} Entr. • {row['Lucro']/stake_padrao:,.2f} STK</div>
                </div>
                <div class="perf-value {c_val}">R$ {row['Lucro']:,.2f}</div>
                <div style="flex:1">
                    <div style="font-size:0.7rem; color:#94a3b8; text-align:right">{row['ROI']:.1f}% ROI</div>
                    <div class="roi-bar-bg"><div class="roi-bar-fill" style="width:{w_roi}%; background:{c_roi}"></div></div>
                </div>
            </div>'''
        html_perf += '</div>'
        st.markdown(html_perf, unsafe_allow_html=True)

except Exception as e:
    st.info("📌 Para carregar: Suba o arquivo CSV para o GitHub com o nome 'AccountStatement_ (1).csv'")
    st.error(f"Detalhe: {e}")
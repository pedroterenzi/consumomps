import streamlit as st
import pandas as pd
import io
import gspread
from google.oauth2.service_account import Credentials

# Configuração da Página
st.set_page_config(page_title="PCP Elite - Automação Drive", layout="wide")

# --- CONFIGURAÇÕES GOOGLE SHEETS ---
# O ID da sua planilha extraído do link que você enviou
SPREADSHEET_ID = '1OAYZ66D4XxE0B4IIlrAtGqOFEx3qVP2s9jqzCbE_lWw'

def enviar_para_google_sheets(df):
    try:
        # Define os escopos de acesso necessários
        scopes = ['https://www.googleapis.com/auth/spreadsheets', 'https://www.googleapis.com/auth/drive']
        
        # Busca as credenciais configuradas nos Secrets do Streamlit
        if "google_credentials" in st.secrets:
            creds = Credentials.from_service_account_info(st.secrets["google_credentials"], scopes=scopes)
            client = gspread.authorize(creds)
            sh = client.open_by_key(SPREADSHEET_ID)
            worksheet = sh.get_worksheet(0) # Acessa a primeira aba da planilha
            
            # Converte o DataFrame para strings (evita erros de formato no Google Sheets)
            df_export = df.astype(str)
            valores = df_export.values.tolist()
            
            # Adiciona os dados na próxima linha disponível (Append)
            worksheet.append_rows(valores)
            return True
        else:
            st.error("Erro: Credenciais 'google_credentials' não encontradas nos Secrets do Streamlit.")
            return False
    except Exception as e:
        st.error(f"Erro na integração com o Drive: {e}")
        return False

# --- FUNÇÕES DE TRATAMENTO DE DADOS ---

def clean_id(val):
    if pd.isna(val) or val == "": return ""
    return str(val).strip().split('.')[0].lstrip('0')

def parse_num(val):
    if pd.isna(val) or val == "" or str(val).strip() == "-": return 0.0
    s = str(val).strip().replace(',', '.')
    try:
        if s.count('.') > 1:
            parts = s.split('.')
            s = "".join(parts[:-1]) + "." + parts[-1]
        return float(s)
    except: return 0.0

def get_consolidated_specs(df_spec, parent_code):
    parent_clean = clean_id(parent_code)
    first_level = df_spec[df_spec['G1_COD'].apply(clean_id) == parent_clean]
    materiais_unitarios = {}
    scale_factor = 1.0
    
    for _, row in first_level.iterrows():
        comp_code = clean_id(row['G1_COMP'])
        qty = parse_num(row['G1_QUANT'])
        if (not comp_code.startswith(('5', '6'))) and qty > 1:
            scale_factor = qty
            sub_items = df_spec[df_spec['G1_COD'].apply(clean_id) == comp_code]
            for _, sub_row in sub_items.iterrows():
                s_sku = clean_id(sub_row['G1_COMP'])
                if s_sku.startswith(('5', '6')):
                    s_qty = parse_num(sub_row['G1_QUANT'])
                    s_desc = sub_row['DESC_COMP'] if 'DESC_COMP' in df_spec.columns else "MP"
                    materiais_unitarios[s_sku] = {'ratio': s_qty, 'desc': s_desc}
            break
            
    for _, row in first_level.iterrows():
        comp_code = clean_id(row['G1_COMP'])
        if comp_code.startswith(('5', '6')):
            qty = parse_num(row['G1_QUANT'])
            desc = row['DESC_COMP'] if 'DESC_COMP' in df_spec.columns else "MP"
            ratio = qty / scale_factor if scale_factor > 0 else qty
            if comp_code not in materiais_unitarios:
                materiais_unitarios[comp_code] = {'ratio': ratio, 'desc': desc}
                
    return materiais_unitarios, scale_factor

# --- INTERFACE STREAMLIT ---

st.title("🚀 PCP Elite - Automação Direta Google Drive")
st.info("O resultado será enviado automaticamente para a Planilha Mestre no Drive.")

# Sidebar para Uploads Manuais
st.sidebar.header("📂 Upload de Arquivos")
f_oficial = st.sidebar.file_uploader("1. Relatório Oficial", type=["xlsm", "xlsx"])
f_spec = st.sidebar.file_uploader("2. SPEC.xlsx", type=["xlsx", "csv"])
f_perdas = st.sidebar.file_uploader("3. Real x Stand (Planilha1)", type=["xlsx"])
f_reg = st.sidebar.file_uploader("4. Controle Requisição", type=["xlsx"])

if f_oficial and f_spec and f_perdas and f_reg:
    with st.spinner('Lendo bases...'):
        df_oficial = pd.read_excel(f_oficial, sheet_name='Result by order')
        df_skus = pd.read_excel(f_oficial, sheet_name='Dados SKUs')
        df_spec = pd.read_excel(f_spec) if f_spec.name.endswith('.xlsx') else pd.read_csv(f_spec)
        df_perdas = pd.read_excel(f_perdas, sheet_name='Planilha1')
        df_reg = pd.read_excel(f_reg, sheet_name='REGISTROS', skiprows=2)

    df_oficial['OP_REF'] = df_oficial['Nº Ordem'].apply(clean_id)
    df_reg['OP_REF'] = df_reg['OP'].apply(clean_id)
    df_reg['SKU_REF'] = df_reg['SKU'].apply(clean_id)
    
    op_alvo = st.selectbox("Selecione a OP Atual", sorted(df_oficial['OP_REF'].unique(), reverse=True))
    op_anterior = st.text_input("Informe a OP Anterior (Opcional)")

    if st.button("📊 Processar e Salvar no Drive"):
        # Lógica de Cálculo
        dados_op = df_oficial[df_oficial['OP_REF'] == op_alvo]
        prod_bruta = dados_op['Machine Counter'].sum()
        prod_estoque = dados_op['Peças Estoque - Ajuste'].sum()
        cod_pai = clean_id(dados_op['Código'].iloc[0])

        materiais, escala_spec = get_consolidated_specs(df_spec, cod_pai)
        sku_info = df_skus[df_skus.iloc[:, 0].apply(clean_id) == cod_pai]
        fardo_estoque = parse_num(sku_info.iloc[0, 3]) if not sku_info.empty else escala_spec

        pre_report = []
        for sku, info in materiais.items():
            if "MOD" in sku or not sku.isdigit(): continue
            
            # Regra Polybag
            if sku.startswith('5905'):
                meta = (prod_estoque / fardo_estoque) * 1.04
                if op_anterior:
                    d_ant = df_oficial[df_oficial['OP_REF'] == clean_id(op_anterior)]
                    c_prev = (d_ant['Peças Estoque - Ajuste'].sum() / fardo_estoque) * 1.04
                else: c_prev = 0
            else:
                # Regra Normal
                f_row = df_perdas[df_perdas['Código'].apply(clean_id) == sku]
                fator = parse_num(f_row['% da espec'].values[0]) if not f_row.empty else 1.0
                meta = (info['ratio'] * prod_bruta) * fator
                if op_anterior:
                    p_ant = df_oficial[df_oficial['OP_REF'] == clean_id(op_anterior)]['Machine Counter'].sum()
                    c_prev = (info['ratio'] * p_ant) * fator
                else: c_prev = 0
            
            # FIFO de Lotes
            ops_busca = [clean_id(op_anterior), op_alvo] if op_anterior else [op_alvo]
            lotes_disp = df_reg[(df_reg['OP_REF'].isin(ops_busca)) & (df_reg['SKU_REF'] == sku)].copy()
            saldo = meta
            reserva = c_prev
            
            if lotes_disp.empty:
                pre_report.append({"OP": op_alvo, "Código": sku, "Descrição": info['desc'], "Quantidade": round(meta, 3), "Lote": "S/ REGISTRO"})
            else:
                for _, l_row in lotes_disp.iterrows():
                    if saldo <= 0: break
                    qtd_ent = parse_num(l_row['QUANTIDADE'])
                    if reserva > 0:
                        gasto = min(reserva, qtd_ent); reserva -= gasto; qtd_ent -= gasto
                    if qtd_ent > 0 and saldo > 0:
                        uso = min(saldo, qtd_ent); saldo -= uso
                        pre_report.append({"OP": clean_id(l_row['OP']), "Código": sku, "Descrição": info['desc'], "Quantidade": uso, "Lote": str(l_row['LOTE']).strip()})
                
                if saldo > 1.0:
                    pre_report.append({"OP": "VERIFICAR", "Código": sku, "Descrição": info['desc'], "Quantidade": round(saldo, 3), "Lote": "FALTA REGISTRO"})

        df_pre = pd.DataFrame(pre_report)
        if not df_pre.empty:
            # Agrupamento final
            df_final = df_pre.groupby(['OP', 'Código', 'Descrição', 'Lote'], as_index=False).agg({'Quantidade': 'sum'})
            df_final = df_final[['OP', 'Código', 'Descrição', 'Quantidade', 'Lote']]
            
            st.write("### Resultado da OP Processada:")
            st.table(df_final.round(3))
            
            # Envio para o Google Drive
         def enviar_para_google_sheets(df):
    try:
        scopes = ['https://www.googleapis.com/auth/spreadsheets', 'https://www.googleapis.com/auth/drive']
        
        if "google_credentials" in st.secrets:
            # Pega as credenciais dos Secrets
            creds_info = dict(st.secrets["google_credentials"])
            
            # CORREÇÃO CRÍTICA: Trata as quebras de linha da private_key
            if "private_key" in creds_info:
                creds_info["private_key"] = creds_info["private_key"].replace("\\n", "\n")
            
            creds = Credentials.from_service_account_info(creds_info, scopes=scopes)
            client = gspread.authorize(creds)
            sh = client.open_by_key(SPREADSHEET_ID)
            worksheet = sh.get_worksheet(0)
            
            df_export = df.astype(str)
            valores = df_export.values.tolist()
            
            worksheet.append_rows(valores)
            return True
        else:
            st.error("Erro: Credenciais 'google_credentials' não encontradas.")
            return False
    except Exception as e:
        st.error(f"Erro na integração com o Drive: {e}")
        return False
else:
    st.info("Por favor, faça o upload dos 4 arquivos na barra lateral para começar.")

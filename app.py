import streamlit as st
import pandas as pd
import io

# Configuração da Página
st.set_page_config(page_title="PCP Elite - Gestão de Saldo", layout="wide")

# --- FUNÇÕES DE TRATAMENTO ---
def clean_id(val):
    if pd.isna(val) or val == "": return ""
    return str(val).strip().split('.')[0].lstrip('0')

def parse_num(val):
    if pd.isna(val) or val == "" or str(val).strip() == "-": return 0.0
    # Se já for número, retorna ele mesmo
    if isinstance(val, (int, float)): return float(val)
    # Trata strings
    s = str(val).strip().replace('.', '').replace(',', '.')
    try:
        return float(s)
    except:
        return 0.0

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

st.title("📊 PCP - Consumo com Gestão de Saldo")

# --- SIDEBAR: UPLOAD ---
st.sidebar.header("📂 Envio de Arquivos")
f_oficial = st.sidebar.file_uploader("1. Relatório Oficial", type=["xlsm", "xlsx"])
f_spec = st.sidebar.file_uploader("2. SPEC.xlsx", type=["xlsx", "csv"])
f_perdas = st.sidebar.file_uploader("3. Real x Stand (Planilha1)", type=["xlsx"])
f_reg = st.sidebar.file_uploader("4. Controle Requisição (Entradas)", type=["xlsx"])
f_historico = st.sidebar.file_uploader("5. Consumo MP (Histórico)", type=["xlsx", "csv"])

if f_oficial and f_spec and f_perdas and f_reg and f_historico:
    with st.spinner('Processando saldos...'):
        df_oficial = pd.read_excel(f_oficial, sheet_name='Result by order')
        df_skus = pd.read_excel(f_oficial, sheet_name='Dados SKUs')
        df_spec = pd.read_excel(f_spec) if f_spec.name.endswith('.xlsx') else pd.read_csv(f_spec)
        df_perdas = pd.read_excel(f_perdas, sheet_name='Planilha1')
        
        # Lendo Controle de Requisição
        df_reg = pd.read_excel(f_reg, sheet_name='REGISTROS', skiprows=2)
        
        # Lendo Histórico de Consumo
        if f_historico.name.endswith('.csv'):
            df_hist = pd.read_csv(f_historico)
        else:
            df_hist = pd.read_excel(f_historico)

    # --- HIGIENIZAÇÃO E CONVERSÃO NUMÉRICA (CORREÇÃO DO ERRO) ---
    df_oficial['OP_REF'] = df_oficial['Nº Ordem'].apply(clean_id)
    
    # Garantir que quantidades sejam números nos dois arquivos
    df_reg['QUANTIDADE_NUM'] = df_reg['QUANTIDADE'].apply(parse_num)
    df_hist['QUANTIDADE_NUM'] = df_hist['Quantidade'].apply(parse_num)
    
    # SKU e Lotes como Strings limpas
    df_reg['SKU_REF'] = df_reg['SKU'].apply(clean_id)
    df_reg['LOTE_REF'] = df_reg['LOTE'].astype(str).str.strip()
    
    df_hist['SKU_REF'] = df_hist['Código'].apply(clean_id)
    df_hist['LOTE_REF'] = df_hist['Lote'].astype(str).str.strip()

    # --- CÁLCULO DE SALDO DISPONÍVEL ---
    # Soma entradas por Lote
    entradas = df_reg.groupby(['SKU_REF', 'LOTE_REF'])['QUANTIDADE_NUM'].sum().reset_index()
    # Soma o que já foi consumido
    consumidos = df_hist.groupby(['SKU_REF', 'LOTE_REF'])['QUANTIDADE_NUM'].sum().reset_index()
    
    # Mescla para calcular saldo
    saldos_lotes = pd.merge(entradas, consumidos, on=['SKU_REF', 'LOTE_REF'], how='left', suffixes=('_ENT', '_CONS'))
    saldos_lotes['QUANTIDADE_NUM_CONS'] = saldos_lotes['QUANTIDADE_NUM_CONS'].fillna(0)
    saldos_lotes['SALDO_REAL'] = saldos_lotes['QUANTIDADE_NUM_ENT'] - saldos_lotes['QUANTIDADE_NUM_CONS']
    
    op_alvo = st.selectbox("Selecione a OP Atual", sorted(df_oficial['OP_REF'].unique(), reverse=True))
    op_anterior = st.text_input("Informe a OP Anterior")

    if st.button("🚀 Gerar Relatório com Saldo Real"):
        dados_op = df_oficial[df_oficial['OP_REF'] == op_alvo]
        prod_bruta = dados_op['Machine Counter'].sum()
        prod_estoque = dados_op['Peças Estoque - Ajuste'].sum()
        cod_pai = clean_id(dados_op['Código'].iloc[0])

        st.info(f"Produção: {prod_bruta} peças | SKU: {cod_pai}")

        materiais, escala_spec = get_consolidated_specs(df_spec, cod_pai)
        sku_info = df_skus[df_skus.iloc[:, 0].apply(clean_id) == cod_pai]
        fardo_estoque = parse_num(sku_info.iloc[0, 3]) if not sku_info.empty else escala_spec

        pre_report = []

        for sku, info in materiais.items():
            if "MOD" in sku or not sku.isdigit(): continue
            
            # Necessidade
            if sku.startswith('5905'):
                meta = (prod_estoque / fardo_estoque) * 1.04
            else:
                f_row = df_perdas[df_perdas['Código'].apply(clean_id) == sku]
                fator = parse_num(f_row['% da espec'].values[0]) if not f_row.empty else 1.0
                meta = (info['ratio'] * prod_bruta) * fator
            
            # Busca lotes que tenham saldo > 0
            lotes_com_estoque = saldos_lotes[(saldos_lotes['SKU_REF'] == sku) & (saldos_lotes['SALDO_REAL'] > 0.01)].copy()
            
            saldo_a_abater = meta
            
            if lotes_com_estoque.empty:
                pre_report.append({"OP": op_alvo, "Código": sku, "Descrição": info['desc'], "Quantidade": round(meta, 3), "Lote": "S/ SALDO"})
            else:
                # Consome os lotes disponíveis
                for _, l_row in lotes_com_estoque.iterrows():
                    if saldo_a_abater <= 0: break
                    
                    disp = l_row['SALDO_REAL']
                    uso = min(saldo_a_abater, disp)
                    
                    pre_report.append({
                        "OP": op_alvo, 
                        "Código": sku, 
                        "Descrição": info['desc'], 
                        "Quantidade": uso, 
                        "Lote": l_row['LOTE_REF']
                    })
                    saldo_a_abater -= uso
                
                if saldo_a_abater > 0.5:
                    pre_report.append({"OP": "VERIFICAR", "Código": sku, "Descrição": info['desc'], "Quantidade": round(saldo_a_abater, 3), "Lote": "FALTA SALDO"})

        df_final = pd.DataFrame(pre_report)
        if not df_final.empty:
            df_final = df_final.groupby(['OP', 'Código', 'Descrição', 'Lote'], as_index=False).agg({'Quantidade': 'sum'})
            df_final['Quantidade'] = df_final['Quantidade'].round(3)
            
            st.table(df_final[['OP', 'Código', 'Descrição', 'Quantidade', 'Lote']])

            # Bloco de cópia
            df_copia = df_final[['OP', 'Código', 'Descrição', 'Quantidade', 'Lote']].copy()
            df_copia['Quantidade'] = df_copia['Quantidade'].apply(lambda x: "{:.3f}".format(x).replace('.', ','))
            dados_brutos = df_copia.to_csv(index=False, header=False, sep='\t')
            st.text_area("Copiar para o Sheets", value=dados_brutos, height=250)
        else:
            st.warning("Nenhum material encontrado.")
else:
    st.info("Aguardando upload dos 5 arquivos.")

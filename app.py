import streamlit as st
import pandas as pd
import io

st.set_page_config(page_title="PCP Elite - Final Gold", layout="wide")

# --- FUNÇÕES DE HIGIENE ---

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
            ratio_unitario = qty / scale_factor if scale_factor > 0 else qty
            if comp_code not in materiais_unitarios:
                materiais_unitarios[comp_code] = {'ratio': ratio_unitario, 'desc': desc}
                
    return materiais_unitarios, scale_factor

st.title("📊 Relatório PCP - Consumo por Lote (Final Gold)")

# --- SIDEBAR ---
f_oficial = st.sidebar.file_uploader("1. Relatório Oficial", type=["xlsm", "xlsx"])
f_spec = st.sidebar.file_uploader("2. SPEC.xlsx", type=["xlsx", "csv"])
f_perdas = st.sidebar.file_uploader("3. Real x Stand (Planilha1)", type=["xlsx"])
f_reg = st.sidebar.file_uploader("4. Controle Requisição", type=["xlsx"])

if f_oficial and f_spec and f_perdas and f_reg:
    with st.spinner('Consolidando dados...'):
        # Carregando abas extras para a regra da Polybag
        df_oficial = pd.read_excel(f_oficial, sheet_name='Result by order')
        df_skus = pd.read_excel(f_oficial, sheet_name='Dados SKUs')
        df_spec = pd.read_excel(f_spec) if f_spec.name.endswith('.xlsx') else pd.read_csv(f_spec)
        df_perdas = pd.read_excel(f_perdas, sheet_name='Planilha1')
        df_reg = pd.read_excel(f_reg, sheet_name='REGISTROS', skiprows=2)

    df_oficial['OP_REF'] = df_oficial['Nº Ordem'].apply(clean_id)
    df_reg['OP_REF'] = df_reg['OP'].apply(clean_id)
    df_reg['SKU_REF'] = df_reg['SKU'].apply(clean_id)
    
    op_alvo = st.selectbox("Selecione a OP Atual", sorted(df_oficial['OP_REF'].unique(), reverse=True))
    op_anterior = st.text_input("Informe a OP Anterior")

    if st.button("🚀 Gerar Relatório Final"):
        dados_op = df_oficial[df_oficial['OP_REF'] == op_alvo]
        prod_real_bruta = dados_op['Machine Counter'].sum()
        # COLUNA AN: Peças Estoque - Ajuste
        prod_estoque_real = dados_op['Peças Estoque - Ajuste'].sum()
        
        cod_pai = clean_id(dados_op['Código'].iloc[0])
        materiais, escala_spec = get_consolidated_specs(df_spec, cod_pai)
        
        # Busca fraldas por pacote na aba Dados SKUs (Coluna D)
        sku_info = df_skus[df_skus.iloc[:, 0].apply(clean_id) == cod_pai]
        fardo_estoque = parse_num(sku_info.iloc[0, 3]) if not sku_info.empty else escala_spec

        pre_report = []

        for sku, info in materiais.items():
            if "MOD" in sku or not sku.isdigit(): continue
            
            # --- LÓGICA ESPECIAL PARA POLYBAG (5905...) ---
            if sku.startswith('5905'):
                # (Total de peças no estoque / Fraldas por pacote) * 1.04
                consumo_meta = (prod_estoque_real / fardo_estoque) * 1.04
                # Para a OP anterior (FIFO), repetimos a lógica se informada
                if op_anterior:
                    dados_ant = df_oficial[df_oficial['OP_REF'] == clean_id(op_anterior)]
                    prod_estoque_ant = dados_ant['Peças Estoque - Ajuste'].sum()
                    consumo_prev_meta = (prod_estoque_ant / fardo_estoque) * 1.04
                else:
                    consumo_prev_meta = 0
            else:
                # --- LÓGICA NORMAL PARA OUTRAS MPs ---
                f_row = df_perdas[df_perdas['Código'].apply(clean_id) == sku]
                fator = parse_num(f_row['% da espec'].values[0]) if not f_row.empty else 1.0
                consumo_meta = (info['ratio'] * prod_real_bruta) * fator
                
                if op_anterior:
                    prod_ant_bruta = df_oficial[df_oficial['OP_REF'] == clean_id(op_anterior)]['Machine Counter'].sum()
                    consumo_prev_meta = (info['ratio'] * prod_ant_bruta) * fator
                else:
                    consumo_prev_meta = 0
            
            # --- RATEIO DE LOTES ---
            ops_busca = [clean_id(op_anterior), op_alvo] if op_anterior else [op_alvo]
            lotes_disp = df_reg[(df_reg['OP_REF'].isin(ops_busca)) & (df_reg['SKU_REF'] == sku)].copy()
            
            saldo_a_preencher = consumo_meta
            reserva_ant = consumo_prev_meta
            
            if lotes_disp.empty:
                pre_report.append({"Cód MP": sku, "Descrição": info['desc'], "Kg": round(consumo_meta, 3), "Lote": "S/ REGISTRO", "OP": op_alvo})
            else:
                for _, l_row in lotes_disp.iterrows():
                    if saldo_a_preencher <= 0: break
                    qtd_entrada = parse_num(l_row['QUANTIDADE'])
                    if reserva_ant > 0:
                        gasto_ant = min(reserva_ant, qtd_entrada)
                        reserva_ant -= gasto_ant
                        qtd_entrada -= gasto_ant
                    if qtd_entrada > 0 and saldo_a_preencher > 0:
                        uso = min(saldo_a_preencher, qtd_entrada)
                        saldo_a_preencher -= uso
                        pre_report.append({"Cód MP": sku, "Descrição": info['desc'], "Kg": uso, "Lote": str(l_row['LOTE']).strip(), "OP": clean_id(l_row['OP'])})
                
                if saldo_a_preencher > 1.0:
                    pre_report.append({"Cód MP": sku, "Descrição": info['desc'], "Kg": saldo_a_preencher, "Lote": "FALTA REGISTRO", "OP": "VERIFICAR"})

        # Agrupamento e Exibição
        df_final = pd.DataFrame(pre_report)
        if not df_final.empty:
            df_final = df_final.groupby(['Cód MP', 'Descrição', 'Lote', 'OP'], as_index=False).agg({'Kg': 'sum'})
            df_final['Kg'] = df_final['Kg'].round(3)
            st.subheader(f"Relatório Consolidado - OP {op_alvo}")
            st.table(df_final)
            
            buffer = io.BytesIO()
            df_final.to_excel(buffer, index=False)
            st.download_button("📥 Baixar Excel", buffer.getvalue(), f"PCP_OP_{op_alvo}.xlsx")

else:
    st.info("Carregue os arquivos para começar.")

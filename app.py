import streamlit as st
import pandas as pd
import io

st.set_page_config(page_title="PCP - Consumo Consolidado", layout="wide")

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
    """ Busca todos os componentes da SPEC e consolida as quantidades unitárias """
    parent_clean = clean_id(parent_code)
    # Busca primeiro nível
    first_level = df_spec[df_spec['G1_COD'].apply(clean_id) == parent_clean]
    
    materiais_unitarios = {}
    scale_factor = 1.0
    
    # Identifica se há um código de fralda (intermediário)
    for _, row in first_level.iterrows():
        comp_code = clean_id(row['G1_COMP'])
        qty = parse_num(row['G1_QUANT'])
        
        if comp_code.startswith('3') and qty > 1:
            scale_factor = qty
            sub_items = df_spec[df_spec['G1_COD'].apply(clean_id) == comp_code]
            for _, sub_row in sub_items.iterrows():
                s_sku = clean_id(sub_row['G1_COMP'])
                s_qty = parse_num(sub_row['G1_QUANT'])
                s_desc = sub_row['DESC_COMP'] if 'DESC_COMP' in df_spec.columns else "MP"
                materiais_unitarios[s_sku] = {'ratio': s_qty, 'desc': s_desc}
        
        elif comp_code.startswith('5') or comp_code.startswith('6'):
            desc = row['DESC_COMP'] if 'DESC_COMP' in df_spec.columns else "MP"
            # Ajusta a espec do pacote para unidade
            ratio = qty / scale_factor if scale_factor > 0 else qty
            if comp_code not in materiais_unitarios:
                materiais_unitarios[comp_code] = {'ratio': ratio, 'desc': desc}
                
    return materiais_unitarios, scale_factor

st.title("📊 Relatório PCP - Consumo por Lote (Final)")

# --- SIDEBAR ---
f_oficial = st.sidebar.file_uploader("1. Relatório Oficial", type=["xlsm", "xlsx"])
f_spec = st.sidebar.file_uploader("2. SPEC.xlsx", type=["xlsx", "csv"])
f_perdas = st.sidebar.file_uploader("3. Real x Stand (Planilha1)", type=["xlsx"])
f_reg = st.sidebar.file_uploader("4. Controle Requisição", type=["xlsx"])

if f_oficial and f_spec and f_perdas and f_reg:
    with st.spinner('Consolidando dados...'):
        df_oficial = pd.read_excel(f_oficial, sheet_name='Result by order')
        df_spec = pd.read_excel(f_spec) if f_spec.name.endswith('.xlsx') else pd.read_csv(f_spec)
        df_perdas = pd.read_excel(f_perdas, sheet_name='Planilha1')
        df_reg = pd.read_excel(f_reg, sheet_name='REGISTROS', skiprows=2)

    df_oficial['OP_REF'] = df_oficial['Nº Ordem'].apply(clean_id)
    df_reg['OP_REF'] = df_reg['OP'].apply(clean_id)
    df_reg['SKU_REF'] = df_reg['SKU'].apply(clean_id)
    
    op_alvo = st.selectbox("Selecione a OP Atual", sorted(df_oficial['OP_REF'].unique(), reverse=True))
    op_anterior = st.text_input("Informe a OP Anterior (para Saldo)")

    if st.button("🚀 Gerar Relatório Consolidado"):
        dados_op = df_oficial[df_oficial['OP_REF'] == op_alvo]
        prod_real = dados_op['Machine Counter'].sum()
        prod_prev = df_oficial[df_oficial['OP_REF'] == clean_id(op_anterior)]['Machine Counter'].sum() if op_anterior else 0
        cod_pai = clean_id(dados_op['Código'].iloc[0])

        materiais, escala = get_consolidated_specs(df_spec, cod_pai)
        
        final_list = []

        for sku, info in materiais.items():
            if "MOD" in sku or not sku.isdigit(): continue
            
            # Fator de Perda
            f_row = df_perdas[df_perdas['Código'].apply(clean_id) == sku]
            fator = parse_num(f_row['% da espec'].values[0]) if not f_row.empty else 1.0
            
            # Consumo total calculado pela ESPEC
            consumo_meta = (info['ratio'] * prod_real) * fator
            consumo_prev_meta = (info['ratio'] * prod_prev) * fator
            
            # Busca lotes (Anterior + Atual)
            ops_busca = [clean_id(op_anterior), op_alvo] if op_anterior else [op_alvo]
            lotes_disp = df_reg[(df_reg['OP_REF'].isin(ops_busca)) & (df_reg['SKU_REF'] == sku)].copy()
            
            saldo_a_preencher = consumo_meta
            reserva_ant = consumo_prev_meta
            
            if lotes_disp.empty:
                final_list.append({"Cód MP": sku, "Descrição": info['desc'], "Kg": round(consumo_meta, 3), "Lote": "S/ REGISTRO", "Status": "Verificar Manual"})
            else:
                for _, l_row in lotes_disp.iterrows():
                    if saldo_a_preencher <= 0: break
                    
                    qtd_lote = parse_num(l_row['QUANTIDADE'])
                    
                    # 1. Abate o que a OP anterior "comeu"
                    if reserva_ant > 0:
                        gasto_ant = min(reserva_ant, qtd_lote)
                        reserva_ant -= gasto_ant
                        qtd_lote -= gasto_ant
                    
                    # 2. O que sobrou preenche a OP atual
                    if qtd_lote > 0 and saldo_a_preencher > 0:
                        uso_op_atual = min(saldo_a_preencher, qtd_lote)
                        saldo_a_preencher -= uso_op_atual
                        final_list.append({
                            "Cód MP": sku, "Descrição": info['desc'], 
                            "Kg": round(uso_op_atual, 3), "Lote": l_row['LOTE'], "Status": f"Entrada {l_row['OP']}"
                        })

                # 3. Se após todos os lotes ainda faltar quilos para bater a meta
                if saldo_a_preencher > 0.1:
                    final_list.append({
                        "Cód MP": sku, "Descrição": info['desc'], 
                        "Kg": round(saldo_a_preencher, 3), "Lote": "SALDO MÁQUINA", "Status": "Pé de Máquina"
                    })

        df_res = pd.DataFrame(final_list)
        st.table(df_res)
        
        buffer = io.BytesIO()
        df_res.to_excel(buffer, index=False)
        st.download_button("📥 Baixar Excel", buffer.getvalue(), f"PCP_Consumo_OP_{op_alvo}.xlsx")

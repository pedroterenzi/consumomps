import streamlit as st
import pandas as pd
import io

st.set_page_config(page_title="PCP Universal - Rastreabilidade", layout="wide")

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

def get_all_components(df_spec, target_code):
    """ Função recursiva ou de varredura total para explodir todos os níveis da SPEC """
    target_clean = clean_id(target_code)
    
    # Lista para armazenar apenas matérias-primas finais (MPs)
    final_mps = {}
    
    def walk_tree(code, multiplier=1.0):
        components = df_spec[df_spec['G1_COD'].apply(clean_id) == clean_id(code)]
        
        for _, row in components.iterrows():
            comp_code = clean_id(row['G1_COMP'])
            # Tira a espec base (G1_QUANT) e divide pelo lote base (B1_QB) se existir
            base_qty = parse_num(row['G1_QUANT'])
            lote_base = parse_num(row['B1_QB']) if 'B1_QB' in row and parse_num(row['B1_QB']) > 0 else 1.0
            
            actual_qty = (base_qty / lote_base) * multiplier
            
            # Se o componente começa com 5 ou 6, é Matéria Prima final
            if comp_code.startswith('5') or comp_code.startswith('6'):
                if comp_code not in final_mps:
                    final_mps[comp_code] = {
                        'qty': 0.0, 
                        'desc': row['DESC_COMP'] if 'DESC_COMP' in row else "MP"
                    }
                final_mps[comp_code]['qty'] += actual_qty
            # Se for outro código (ex: começa com 1 ou 3), continua descendo na árvore
            elif comp_code.isdigit() and not comp_code.startswith('5'):
                walk_tree(comp_code, actual_qty)

    walk_tree(target_clean)
    return final_mps

st.title("📊 PCP Universal - Consumo Real por Lote")

# --- SIDEBAR ---
f_oficial = st.sidebar.file_uploader("1. Relatório Oficial", type=["xlsm", "xlsx"])
f_spec = st.sidebar.file_uploader("2. SPEC.xlsx", type=["xlsx", "csv"])
f_perdas = st.sidebar.file_uploader("3. Real x Stand (Opcional - Aba Planilha1)", type=["xlsx"])
f_reg = st.sidebar.file_uploader("4. Controle Requisição", type=["xlsx"])

# Dicionário de backup para perdas (caso não queira subir o Real x Stand)
PERDAS_PADRAO = {"5905": 1.04} # Polybags 104%

if f_oficial and f_spec and f_reg:
    with st.spinner('Varrendo estrutura de produtos...'):
        df_oficial = pd.read_excel(f_oficial, sheet_name='Result by order')
        df_spec = pd.read_excel(f_spec) if f_spec.name.endswith('.xlsx') else pd.read_csv(f_spec)
        df_reg = pd.read_excel(f_reg, sheet_name='REGISTROS', skiprows=2)
        
        df_perdas = None
        if f_perdas:
            df_perdas = pd.read_excel(f_perdas, sheet_name='Planilha1')

    df_oficial['OP_REF'] = df_oficial['Nº Ordem'].apply(clean_id)
    df_reg['OP_REF'] = df_reg['OP'].apply(clean_id)
    df_reg['SKU_REF'] = df_reg['SKU'].apply(clean_id)
    
    op_alvo = st.selectbox("Selecione a OP", sorted(df_oficial['OP_REF'].unique(), reverse=True))
    op_anterior = st.text_input("OP Anterior (para saldo)")

    if st.button("🚀 Gerar Relatório"):
        # 1. Identifica Produto e Produção
        dados_op = df_oficial[df_oficial['OP_REF'] == op_alvo]
        prod_real = dados_op['Machine Counter'].sum()
        prod_prev = df_oficial[df_oficial['OP_REF'] == clean_id(op_anterior)]['Machine Counter'].sum() if op_anterior else 0
        cod_pai = clean_id(dados_op['Código'].iloc[0])

        # 2. Explode TODA a árvore de materiais
        materiais = get_all_components(df_spec, cod_pai)
        
        pre_report = []

        for sku, info in materiais.items():
            # Define Fator de Perda
            fator = 1.0
            if sku.startswith('5905'):
                fator = 1.04
            elif df_perdas is not None:
                f_row = df_perdas[df_perdas['Código'].apply(clean_id) == sku]
                fator = parse_num(f_row['% da espec'].values[0]) if not f_row.empty else 1.0
            
            # Meta de consumo (Kg)
            consumo_meta = (info['qty'] * prod_real) * fator
            consumo_prev = (info['qty'] * prod_prev) * fator
            
            # Busca Lotes
            ops_busca = [clean_id(op_anterior), op_alvo] if op_anterior else [op_alvo]
            lotes_disp = df_reg[(df_reg['OP_REF'].isin(ops_busca)) & (df_reg['SKU_REF'] == sku)].copy()
            
            saldo_a_preencher = consumo_meta
            reserva_ant = consumo_prev
            
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
            st.table(df_final)
            
            buffer = io.BytesIO()
            df_final.to_excel(buffer, index=False)
            st.download_button("📥 Baixar Excel", buffer.getvalue(), f"PCP_OP_{op_alvo}.xlsx")
        else:
            st.error("Nenhum material encontrado.")

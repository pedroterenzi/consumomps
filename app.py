import streamlit as st
import pandas as pd
import io

st.set_page_config(page_title="PCP - Consumo Real por SPEC", layout="wide")

# --- FUNÇÕES DE LIMPEZA E CONVERSÃO ---

def clean_id(val):
    if pd.isna(val) or val == "": return ""
    return str(val).strip().split('.')[0].lstrip('0')

def parse_num_spec(val):
    """Converte a G1_QUANT garantindo que seja um decimal (evita os bilhões)"""
    if pd.isna(val) or val == "": return 0.0
    s = str(val).strip().replace(',', '.')
    try:
        num = float(s)
        # Se o número for absurdamente alto (ex > 100), provavelmente é erro de decimal do Excel
        if num > 1000: 
            return num / 1000000
        return num
    except:
        return 0.0

def parse_general_num(val):
    """Converte Machine Counter e Quantidades de Lote"""
    if pd.isna(val) or val == "" or str(val).strip() == "-": return 0.0
    s = str(val).strip().replace('.', '').replace(',', '.')
    try:
        return float(s)
    except:
        return 0.0

st.title("📊 Relatório de Consumo Real (Base SPEC)")

# --- SIDEBAR ---
st.sidebar.header("Upload de Arquivos")
f_oficial = st.sidebar.file_uploader("1. Relatório Oficial (Peças)", type=["xlsm", "xlsx"])
f_spec = st.sidebar.file_uploader("2. SPEC.xlsx (Especs)", type=["xlsx", "csv"])
f_perdas = st.sidebar.file_uploader("3. Real x Stand (Aba Planilha1)", type=["xlsx"])
f_reg = st.sidebar.file_uploader("4. Controle Requisição (Lotes)", type=["xlsx"])

if f_oficial and f_spec and f_perdas and f_reg:
    with st.spinner('Sincronizando dados...'):
        df_oficial = pd.read_excel(f_oficial, sheet_name='Result by order')
        df_spec = pd.read_excel(f_spec) if f_spec.name.endswith('.xlsx') else pd.read_csv(f_spec)
        df_perdas = pd.read_excel(f_perdas, sheet_name='Planilha1')
        df_reg = pd.read_excel(f_reg, sheet_name='REGISTROS', skiprows=2)

    # Preparar IDs
    df_oficial['OP_REF'] = df_oficial['Nº Ordem'].apply(clean_id)
    df_reg['OP_REF'] = df_reg['OP'].apply(clean_id)
    df_reg['SKU_REF'] = df_reg['SKU'].apply(clean_id)
    
    op_alvo = st.selectbox("Selecione a OP", sorted(df_oficial['OP_REF'].unique(), reverse=True))
    op_anterior = st.text_input("Informe a OP Anterior (para herança de lote)")

    if st.button("Gerar Consumo Real"):
        # 1. Produção Real (Soma da coluna Machine Counter)
        dados_op = df_oficial[df_oficial['OP_REF'] == op_alvo]
        prod_real_total = dados_op['Machine Counter'].sum()
        cod_pai = clean_id(dados_op['Código'].iloc[0]) # Código do produto final
        
        # Produção da anterior (para o cálculo de FIFO)
        prod_prev = df_oficial[df_oficial['OP_REF'] == clean_id(op_anterior)]['Machine Counter'].sum() if op_anterior else 0

        # 2. Buscar Componentes na SPEC
        specs_do_produto = df_spec[df_spec['G1_COD'].apply(clean_id) == cod_pai]
        
        if specs_do_produto.empty:
            st.error(f"O produto {cod_pai} não foi encontrado no arquivo SPEC.xlsx")
        else:
            st.success(f"OP: {op_alvo} | Produto: {cod_pai} | Produção Real: {prod_real_total:,.0f} peças")
            
            resultado_final = []

            for _, row in specs_do_produto.iterrows():
                mp_codigo = clean_id(row['G1_COMP'])
                espec_unitaria = parse_num_spec(row['G1_QUANT'])
                desc_mp = row['DESC_COMP'] if 'DESC_COMP' in df_spec.columns else "Materia Prima"
                
                # Fator de Perda
                f_row = df_perdas[df_perdas['Código'].apply(clean_id) == mp_codigo]
                fator_perda = parse_general_num(f_row['% da espec'].values[0]) if not f_row.empty else 1.0
                
                # CÁLCULO DO CONSUMO (Kg)
                kg_necessario_atual = (espec_unitaria * prod_real_total) * fator_perda
                kg_necessario_prev = (espec_unitaria * prod_prev) * fator_perda
                
                # 3. Cruzamento com REGISTROS (Lotes)
                ops_busca = [clean_id(op_anterior), op_alvo] if op_anterior else [op_alvo]
                lotes_disponiveis = df_reg[(df_reg['OP_REF'].isin(ops_busca)) & (df_reg['SKU_REF'] == mp_codigo)].copy()
                
                restante_prev = kg_necessario_prev
                restante_atual = kg_necessario_atual
                
                if lotes_disponiveis.empty:
                    resultado_final.append({
                        "Código MP": mp_codigo, "Descrição": desc_mp, "Consumo (Kg)": round(restante_atual, 3),
                        "Lote": "S/ REGISTRO", "Origem": "Verificar Manual"
                    })
                else:
                    for _, l_row in lotes_disponiveis.iterrows():
                        if restante_atual <= 0 and restante_prev <= 0: break
                        
                        qtd_pallet = parse_general_num(l_row['QUANTIDADE'])
                        lote_id = l_row['LOTE']
                        
                        # Abate o consumo da OP anterior primeiro
                        if restante_prev > 0:
                            gasto_prev = min(restante_prev, qtd_pallet)
                            restante_prev -= gasto_prev
                            qtd_pallet -= gasto_prev
                        
                        # O que sobrou vai para a OP atual
                        if qtd_pallet > 0 and restante_atual > 0:
                            gasto_atual = min(restante_atual, qtd_pallet)
                            restante_atual -= gasto_atual
                            resultado_final.append({
                                "Código MP": mp_codigo, "Descrição": desc_mp, 
                                "Consumo (Kg)": round(gasto_atual, 3), "Lote": lote_id, "Origem": f"Entrada {l_row['OP']}"
                            })

                    if restante_atual > 0.1:
                        resultado_final.append({
                            "Código MP": mp_codigo, "Descrição": desc_mp, 
                            "Consumo (Kg)": round(restante_atual, 3), "Lote": "SALDO ANTIGO", "Origem": "Estoque Máquina"
                        })

            df_res = pd.DataFrame(resultado_final)
            st.table(df_res)
            
            # Exportar Excel
            buffer = io.BytesIO()
            df_res.to_excel(buffer, index=False)
            st.download_button("📥 Baixar Relatório", buffer.getvalue(), f"Consumo_Real_{op_alvo}.xlsx")

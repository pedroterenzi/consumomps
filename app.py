import streamlit as st
import pandas as pd
import io

st.set_page_config(page_title="PCP - Sistema de Consumo Integrado", layout="wide")

# --- FUNÇÕES DE HIGIENE ---
def clean_id(val):
    if pd.isna(val) or val == "": return ""
    return str(val).strip().split('.')[0].lstrip('0')

def parse_num(val):
    if pd.isna(val) or val == "" or str(val).strip() == "-": return 0.0
    s = str(val).replace('.', '').replace(',', '.')
    try: return float(s)
    except: return 0.0

st.title("🚀 Sistema PCP - Consumo por Lote (V3 - SPEC Integration)")
st.markdown("---")

# --- SIDEBAR: UPLOADS ---
st.sidebar.header("📂 Arquivos Necessários")
f_oficial = st.sidebar.file_uploader("1. Relatório Oficial (Produção)", type=["xlsm", "xlsx"])
f_stand = st.sidebar.file_uploader("2. Real x Stand (Perdas)", type=["xlsx"])
f_spec = st.sidebar.file_uploader("3. SPEC.xlsx (Engenharia)", type=["xlsx", "csv"])
f_reg = st.sidebar.file_uploader("4. Controle Requisição (Lotes)", type=["xlsx"])

if f_oficial and f_stand and f_spec and f_reg:
    try:
        with st.spinner('Sincronizando bases de dados...'):
            # Leitura das bases
            df_oficial = pd.read_excel(f_oficial, sheet_name='Result by order')
            df_stand_perdas = pd.read_excel(f_stand, sheet_name='Planilha1')
            df_reg = pd.read_excel(f_reg, sheet_name='REGISTROS', skiprows=2)
            
            # Leitura da SPEC (Engenharia)
            if f_spec.name.endswith('.csv'):
                df_spec = pd.read_csv(f_spec)
            else:
                df_spec = pd.read_excel(f_spec)

            # Padronização de Chaves
            df_oficial['OP_REF'] = df_oficial['Nº Ordem'].apply(clean_id)
            df_reg['OP_REF'] = df_reg['OP'].apply(clean_id)
            df_reg['SKU_REF'] = df_reg['SKU'].apply(clean_id)
            
            # Pegando os códigos de produto final do Relatório Oficial (coluna Código)
            ops_disponiveis = sorted(df_oficial['OP_REF'].unique(), reverse=True)

        # --- INTERFACE ---
        col1, col2 = st.columns(2)
        with col1:
            op_selecionada = st.selectbox("Selecione a OP Atual", ops_disponiveis)
        with col2:
            op_anterior = st.text_input("OP Anterior (Herança de Saldo)")

        if st.button("📊 Gerar Relatório de Consumo"):
            # 1. Identificar o Código do Produto Final e Produção Real
            dados_op = df_oficial[df_oficial['OP_REF'] == op_selecionada]
            cod_produto_final = clean_id(dados_op['Código'].iloc[0])
            prod_real_atual = dados_op['Machine Counter'].sum()
            
            prod_real_prev = 0
            if op_anterior:
                prod_real_prev = df_oficial[df_oficial['OP_REF'] == clean_id(op_anterior)]['Machine Counter'].sum()

            # 2. Buscar Especs na Engenharia (SPEC.xlsx)
            specs_produto = df_spec[df_spec['G1_COD'].apply(clean_id) == cod_produto_final]

            if specs_produto.empty:
                st.error(f"Código de produto {cod_produto_final} não encontrado no arquivo SPEC.xlsx")
            else:
                st.info(f"Produto: {cod_produto_final} | Produção: {prod_real_atual:,.0f} peças")
                
                relatorio_final = []

                for _, row_spec in specs_produto.iterrows():
                    mp_sku = clean_id(row_spec['G1_COMP'])
                    espec_unitaria = parse_num(row_spec['G1_QUANT'])
                    desc_mp = row_spec['DESC_COMP'] if 'DESC_COMP' in df_spec.columns else "MP"
                    
                    # Fator de Perda (Planilha1)
                    f_row = df_stand_perdas[df_stand_perdas['Código'].apply(clean_id) == mp_sku]
                    fator_perda = parse_num(f_row['% da espec'].values[0]) if not f_row.empty else 1.0
                    
                    # Cálculo de Demanda (Kg)
                    kg_necessario_atual = (espec_unitaria * prod_real_atual) * fator_perda
                    kg_necessario_prev = (espec_unitaria * prod_real_prev) * fator_perda
                    
                    # 3. Cruzamento com Lotes (Lógica FIFO)
                    ops_busca = [clean_id(op_anterior), op_selecionada] if op_anterior else [op_selecionada]
                    lotes_db = df_reg[(df_reg['OP_REF'].isin(ops_busca)) & (df_reg['SKU_REF'] == mp_sku)].copy()
                    
                    reserva_prev = kg_necessario_prev
                    reserva_atual = kg_necessario_atual
                    
                    if lotes_db.empty:
                        relatorio_final.append({
                            "Código MP": mp_sku, "Descrição": desc_mp, "Consumo (Kg)": round(reserva_atual, 3),
                            "Lote": "S/ REGISTRO", "Origem": "Verificar Manual"
                        })
                    else:
                        for _, l_row in lotes_db.iterrows():
                            qtd_lote = parse_num(l_row['QUANTIDADE'])
                            lote_id = l_row['LOTE']
                            
                            # Consome primeiro para a anterior
                            if reserva_prev > 0:
                                gasto_prev = min(reserva_prev, qtd_lote)
                                reserva_prev -= gasto_prev
                                qtd_lote -= gasto_prev
                            
                            # O que sobrou vai para a atual
                            if qtd_lote > 0 and reserva_atual > 0:
                                gasto_atual = min(reserva_atual, qtd_lote)
                                reserva_atual -= gasto_atual
                                relatorio_final.append({
                                    "Código MP": mp_sku, "Descrição": desc_mp, 
                                    "Consumo (Kg)": round(gasto_atual, 3), "Lote": lote_id, "Origem": f"Entrada {l_row['OP']}"
                                })

                        if reserva_atual > 0.5:
                            relatorio_final.append({
                                "Código MP": mp_sku, "Descrição": desc_mp, 
                                "Consumo (Kg)": round(reserva_atual, 3), "Lote": "SALDO MÁQUINA", "Origem": "Estoque Antigo"
                            })

                df_res = pd.DataFrame(relatorio_final)
                st.table(df_res)
                
                # Download
                buffer = io.BytesIO()
                with pd.ExcelWriter(buffer, engine='xlsxwriter') as writer:
                    df_res.to_excel(writer, index=False)
                st.download_button("📥 Baixar Relatório PCP", buffer.getvalue(), f"Consumo_Lote_OP_{op_selecionada}.xlsx")

    except Exception as e:
        st.error(f"Erro no processamento: {e}")
else:
    st.info("Aguardando o upload dos 4 arquivos necessários.")

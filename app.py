import streamlit as st
import pandas as pd
import io

st.set_page_config(page_title="PCP - Consumo Real Blindado", layout="wide")

# --- FUNÇÕES DE LIMPEZA PESADA (PARA EVITAR BILHÕES) ---

def clean_id(val):
    if pd.isna(val) or val == "": return ""
    return str(val).strip().split('.')[0].lstrip('0')

def parse_espec_clean(val):
    """ Impede que 0.007 vire 7 milhões por erro de formatação do Excel """
    if pd.isna(val) or val == "" or str(val).strip() == "-": return 0.0
    # Remove qualquer ponto de milhar e troca vírgula por ponto
    s = str(val).strip().replace('.', '').replace(',', '.')
    try:
        num = float(s)
        # Lógica: uma espec de MP por fralda NUNCA é maior que 1kg (1.0)
        # Se for maior, é porque o Excel leu errado os decimais (ex: 7000 ao invés de 0.007)
        while num > 1.0:
            num = num / 1000
        return num
    except:
        return 0.0

def parse_prod_real(val):
    """ Para Machine Counter e Quantidades de Pallet """
    if pd.isna(val) or val == "" or str(val).strip() == "-": return 0.0
    s = str(val).strip().replace('.', '').replace(',', '.')
    try:
        return float(s)
    except:
        return 0.0

st.title("📊 Relatório de Consumo Real PCP")
st.markdown("---")

# --- SIDEBAR ---
st.sidebar.header("📁 Upload de Arquivos")
f_oficial = st.sidebar.file_uploader("1. Relatório Oficial (Excel)", type=["xlsm", "xlsx"])
f_spec = st.sidebar.file_uploader("2. SPEC.xlsx (CSV ou Excel)", type=["xlsx", "csv"])
f_perdas = st.sidebar.file_uploader("3. Real x Stand (Aba Planilha1)", type=["xlsx"])
f_reg = st.sidebar.file_uploader("4. Controle Requisição (Lotes)", type=["xlsx"])

if f_oficial and f_spec and f_perdas and f_reg:
    try:
        with st.spinner('Limpando e processando bases...'):
            # Carregamento
            df_oficial = pd.read_excel(f_oficial, sheet_name='Result by order')
            
            if f_spec.name.endswith('.csv'):
                df_spec = pd.read_csv(f_spec)
            else:
                df_spec = pd.read_excel(f_spec)
                
            df_perdas = pd.read_excel(f_perdas, sheet_name='Planilha1')
            df_reg = pd.read_excel(f_reg, sheet_name='REGISTROS', skiprows=2)

            # Higienização de IDs
            df_oficial['OP_REF'] = df_oficial['Nº Ordem'].apply(clean_id)
            df_reg['OP_REF'] = df_reg['OP'].apply(clean_id)
            df_reg['SKU_REF'] = df_reg['SKU'].apply(clean_id)
            df_spec['G1_COD_CLEAN'] = df_spec['G1_COD'].apply(clean_id)
            
            ops_disponiveis = sorted(df_oficial['OP_REF'].unique(), reverse=True)

        # --- SELEÇÃO ---
        col1, col2 = st.columns(2)
        with col1:
            op_alvo = st.selectbox("Selecione a OP Atual", ops_disponiveis)
        with col2:
            op_anterior = st.text_input("Informe a OP Anterior (para herança de lote)")

        if st.button("🚀 Gerar Relatório de Consumo"):
            # 1. Produção Real
            dados_atual = df_oficial[df_oficial['OP_REF'] == op_alvo]
            prod_real_atual = dados_atual['Machine Counter'].sum()
            cod_pai = clean_id(dados_atual['Código'].iloc[0])
            
            prod_real_prev = 0
            if op_anterior:
                prod_real_prev = df_oficial[df_oficial['OP_REF'] == clean_id(op_anterior)]['Machine Counter'].sum()

            # 2. Filtrar Componentes na SPEC (ignorando Mão de Obra e itens sem código numérico)
            specs_op = df_spec[df_spec['G1_COD_CLEAN'] == cod_pai].copy()
            
            if specs_op.empty:
                st.error(f"Produto {cod_pai} não encontrado no SPEC.xlsx")
            else:
                st.success(f"OP: {op_alvo} | Produção: {prod_real_atual:,.0f} peças")
                
                final_list = []

                for _, row in specs_op.iterrows():
                    mp_sku = clean_id(row['G1_COMP'])
                    # Pular se for Mão de Obra ou código inválido
                    if "MOD" in mp_sku or not mp_sku.isdigit(): continue
                    
                    espec_corrigida = parse_espec_clean(row['G1_QUANT'])
                    desc_mp = row['DESC_COMP'] if 'DESC_COMP' in df_spec.columns else "Materia Prima"
                    
                    # Fator de Perda
                    f_row = df_perdas[df_perdas['Código'].apply(clean_id) == mp_sku]
                    fator = parse_prod_real(f_row['% da espec'].values[0]) if not f_row.empty else 1.0
                    
                    # CÁLCULO FINAL (KG)
                    demanda_atual = (espec_corrigida * prod_real_atual) * fator
                    demanda_prev = (espec_corrigida * prod_real_prev) * fator
                    
                    # 3. Cruzamento de Lotes (FIFO)
                    ops_busca = [clean_id(op_anterior), op_alvo] if op_anterior else [op_alvo]
                    lotes_disp = df_reg[(df_reg['OP_REF'].isin(ops_busca)) & (df_reg['SKU_REF'] == mp_sku)].copy()
                    
                    reserva_anterior = demanda_prev
                    balde_atual = demanda_atual
                    
                    if lotes_disp.empty:
                        final_list.append({
                            "Cód MP": mp_sku, "Descrição": desc_mp, "Qtd (Kg)": round(balde_atual, 3),
                            "Lote": "S/ REGISTRO", "Status": "Verificar Manual"
                        })
                    else:
                        for _, l_row in lotes_disp.iterrows():
                            if balde_atual <= 0: break
                            
                            qtd_lote = parse_prod_real(l_row['QUANTIDADE'])
                            lote_id = l_row['LOTE']
                            
                            # Consome primeiro o que a anterior usou
                            if reserva_anterior > 0:
                                gasto_ant = min(reserva_anterior, qtd_lote)
                                reserva_anterior -= gasto_ant
                                qtd_lote -= gasto_ant
                            
                            # O que sobrou vai para a atual
                            if qtd_lote > 0 and balde_atual > 0:
                                uso_atual = min(balde_atual, qtd_lote)
                                balde_atual -= uso_atual
                                final_list.append({
                                    "Cód MP": mp_sku, "Descrição": desc_mp, 
                                    "Qtd (Kg)": round(uso_atual, 3), "Lote": lote_id, "Status": f"Entrada {l_row['OP']}"
                                })
                        
                        if balde_atual > 0.1:
                            final_list.append({
                                "Cód MP": mp_sku, "Descrição": desc_mp, 
                                "Qtd (Kg)": round(balde_atual, 3), "Lote": "SALDO MÁQUINA", "Status": "Estoque Antigo"
                            })

                df_res = pd.DataFrame(final_list)
                st.table(df_res)
                
                # Download Excel
                output = io.BytesIO()
                df_res.to_excel(output, index=False)
                st.download_button("📥 Baixar Relatório", output.getvalue(), f"Consumo_Real_OP_{op_alvo}.xlsx")

    except Exception as e:
        st.error(f"Erro no processamento: {e}")

else:
    st.info("Aguardando upload dos arquivos na barra lateral...")

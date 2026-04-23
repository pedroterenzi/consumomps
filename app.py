import streamlit as st
import pandas as pd
import io

st.set_page_config(page_title="PCP - Rastreabilidade de Lotes", layout="wide")

# --- FUNÇÕES DE LIMPEZA E CONVERSÃO ---

def clean_id(val):
    if pd.isna(val) or val == "": return ""
    return str(val).strip().split('.')[0].lstrip('0')

def parse_num(val):
    """ Converte valores de forma segura, tratando pontos de milhar e vírgulas """
    if pd.isna(val) or val == "": return 0.0
    s = str(val).strip()
    # Se houver mais de um ponto e uma vírgula (ex: 1.234.567,89)
    if s.count('.') >= 1 and ',' in s:
        s = s.replace('.', '').replace(',', '.')
    # Se houver apenas vírgula (ex: 1234,56)
    elif ',' in s:
        s = s.replace(',', '.')
    try:
        res = float(s)
        # Proteção contra valores astronômicos por erro de leitura de decimal
        return res if res < 1000000 else res / 1000 
    except:
        return 0.0

st.title("📊 Relatório de Consumo Real por Lote")

# --- SIDEBAR ---
st.sidebar.header("Arquivos")
file_oficial = st.sidebar.file_uploader("Relatório Oficial (Peças Real)", type=["xlsm", "xlsx"])
file_stand = st.sidebar.file_uploader("Real x Stand (Specs e Perdas)", type=["xlsx"])
file_registros = st.sidebar.file_uploader("Controle Requisição (Lotes)", type=["xlsx"])

if file_oficial and file_stand and file_registros:
    with st.spinner('Processando dados...'):
        df_oficial = pd.read_excel(file_oficial, sheet_name='Result by order')
        df_stand = pd.read_excel(file_stand, sheet_name='2-Totais por OP   Produto')
        df_perdas = pd.read_excel(file_stand, sheet_name='Planilha1')
        df_reg = pd.read_excel(file_registros, sheet_name='REGISTROS', skiprows=2)

    # Limpeza de colunas
    df_oficial['OP_REF'] = df_oficial['Nº Ordem'].apply(clean_id)
    df_reg['OP_REF'] = df_reg['OP'].apply(clean_id)
    df_reg['SKU_REF'] = df_reg['SKU'].apply(clean_id)
    df_stand['OP_REF'] = df_stand.iloc[:, 0].apply(clean_id)
    
    col1, col2 = st.columns(2)
    with col1:
        op_alvo = st.selectbox("Selecione a OP Atual (ex: 18940)", sorted(df_oficial['OP_REF'].unique(), reverse=True))
    with col2:
        op_anterior = st.text_input("Informe a OP Anterior (ex: 18938)")

    if st.button("Gerar Relatório PCP"):
        # 1. Produção Real
        prod_real = df_oficial[df_oficial['OP_REF'] == op_alvo]['Machine Counter'].sum()
        
        # 2. Localizar Materiais na Real x Stand
        materiais = df_stand[df_stand.iloc[:, 0].astype(str).str.contains(op_alvo)].copy()
        
        final_report = []

        for _, row in materiais.iterrows():
            sku = clean_id(row.iloc[4]) # M A T E R I A L CODIGO
            if not sku or not str(row.iloc[4]).isdigit(): continue
            
            # Cálculo da ESPEC manual conforme seu relato
            qtd_programada_op = parse_num(row.iloc[3])
            qtd_std_mp = parse_num(row.iloc[11])
            espec = qtd_std_mp / qtd_programada_op if qtd_programada_op > 0 else 0
            
            # Fator de Perda
            fator_row = df_perdas[df_perdas['Código'].apply(clean_id) == sku]
            fator = fator_row['% da espec'].values[0] if not fator_row.empty else 1.0
            
            # TOTAL NECESSÁRIO EM KG PARA A OP ATUAL
            consumo_necessario = (espec * prod_real) * fator
            
            # 3. LÓGICA DE HERANÇA DE LOTE (OP Anterior + OP Atual)
            # Pegamos tudo que entrou para esse SKU nas duas OPs
            ops_para_busca = [op_alvo]
            if op_anterior: ops_para_busca.insert(0, clean_id(op_anterior))
            
            lotes_disponiveis = df_reg[(df_reg['OP_REF'].isin(ops_para_busca)) & (df_reg['SKU_REF'] == sku)].copy()
            
            if lotes_disponiveis.empty:
                final_report.append({
                    "Código": sku, "Descrição": row.iloc[5], "Quantidade (Kg)": round(consumo_necessario, 2),
                    "Lote": "LOTE NÃO ENCONTRADO", "Origem": "Verificar Físico"
                })
            else:
                # Aqui simulamos o consumo: o saldo que sobra de uma vai para a outra
                saldo_a_abater = consumo_necessario
                for _, lote_row in lotes_disponiveis.iterrows():
                    if saldo_necessario <= 0: break
                    
                    qtd_lote = parse_num(lote_row['QUANTIDADE'])
                    lote_id = lote_row['LOTE']
                    
                    consumo_deste_lote = min(saldo_a_abater, qtd_lote)
                    saldo_a_abater -= consumo_deste_lote
                    
                    final_report.append({
                        "Código": sku, "Descrição": row.iloc[5], 
                        "Quantidade (Kg)": round(consumo_deste_lote, 2),
                        "Lote": lote_id, "Origem": f"OP {lote_row['OP']}"
                    })
                
                # Se ainda faltar, indica que veio de mais atrás
                if saldo_a_abater > 0.5:
                    final_report.append({
                        "Código": sku, "Descrição": row.iloc[5], 
                        "Quantidade (Kg)": round(saldo_a_abater, 2),
                        "Lote": "SALDO REMANESCENTE", "Origem": "Estoque Máquina"
                    })

        df_final = pd.DataFrame(final_report)
        st.write(f"### Relatório de Consumo - OP {op_alvo}")
        st.table(df_final)

        # Download
        towrite = io.BytesIO()
        df_final.to_excel(towrite, index=False, engine='xlsxwriter')
        st.download_button("📥 Baixar Excel para o PCP", towrite.getvalue(), f"Consumo_Lote_OP_{op_alvo}.xlsx")

else:
    st.info("Aguardando upload das planilhas para calcular os consumos.")

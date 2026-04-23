import streamlit as st
import pandas as pd
import io

st.set_page_config(page_title="PCP - Rastreabilidade Total", layout="wide")

# --- FUNÇÕES DE LIMPEZA ---
def clean_id(val):
    if pd.isna(val) or val == "": return ""
    return str(val).strip().split('.')[0].lstrip('0')

def parse_num(val):
    """ Converte valores tratando o problema de pontos/vírgulas do Excel BR """
    if pd.isna(val) or val == "": return 0.0
    s = str(val).strip()
    
    # Se tiver vírgula, é formato brasileiro (ex: 1.100,50 ou 1100,50)
    if "," in s:
        s = s.replace(".", "").replace(",", ".")
    # Se não tiver vírgula mas tiver mais de um ponto (ex: 1.100.500)
    elif s.count(".") > 1:
        s = s.replace(".", "")
    
    try:
        return float(s)
    except:
        return 0.0

st.title("📊 Relatório de Consumo PCP (Com Herança de Lote)")
st.markdown("Este sistema abate o consumo da OP anterior dos lotes registrados e traz o saldo para a OP atual.")

# --- SIDEBAR ---
st.sidebar.header("Upload das Planilhas")
file_oficial = st.sidebar.file_uploader("1. Relatório Oficial (Produção Real)", type=["xlsm", "xlsx"])
file_stand = st.sidebar.file_uploader("2. Real x Stand (Specs e Perdas)", type=["xlsx"])
file_registros = st.sidebar.file_uploader("3. Controle Requisição (Lotes)", type=["xlsx"])

if file_oficial and file_stand and file_registros:
    with st.spinner('Lendo arquivos...'):
        df_oficial = pd.read_excel(file_oficial, sheet_name='Result by order')
        df_stand = pd.read_excel(file_stand, sheet_name='2-Totais por OP   Produto')
        df_perdas = pd.read_excel(file_stand, sheet_name='Planilha1')
        df_reg = pd.read_excel(file_registros, sheet_name='REGISTROS', skiprows=2)

    # Preparar referências
    df_oficial['OP_REF'] = df_oficial['Nº Ordem'].apply(clean_id)
    df_reg['OP_REF'] = df_reg['OP'].apply(clean_id)
    df_reg['SKU_REF'] = df_reg['SKU'].apply(clean_id)
    df_stand['OP_REF'] = df_stand.iloc[:, 0].apply(clean_id)
    
    ops_list = sorted(df_oficial['OP_REF'].unique(), reverse=True)
    
    col1, col2 = st.columns(2)
    with col1:
        op_alvo = st.selectbox("Selecione a OP Atual", ops_list)
    with col2:
        op_anterior = st.text_input("Informe a OP Anterior (para herança de lote)")

    if st.button("🚀 Gerar Relatório"):
        # 1. Produções Reais
        prod_atual = df_oficial[df_oficial['OP_REF'] == op_alvo]['Machine Counter'].sum()
        prod_prev = 0
        if op_anterior:
            prod_prev = df_oficial[df_oficial['OP_REF'] == clean_id(op_anterior)]['Machine Counter'].sum()
        
        # 2. Materiais da OP Atual
        materiais = df_stand[df_stand.iloc[:, 0].astype(str).str.contains(op_alvo)].copy()
        
        relatorio_final = []

        for _, row in materiais.iterrows():
            sku = clean_id(row.iloc[4])
            if not sku or not str(row.iloc[4]).isdigit(): continue
            
            desc = row.iloc[5]
            
            # Cálculo de Espec e Perda
            q_prog = parse_num(row.iloc[3])
            q_std = parse_num(row.iloc[11])
            espec = q_std / q_prog if q_prog > 0 else 0
            
            fator_row = df_perdas[df_perdas['Código'].apply(clean_id) == sku]
            fator = fator_row['% da espec'].values[0] if not fator_row.empty else 1.0
            
            # Demanda das duas OPs
            demanda_atual = (espec * prod_atual) * fator
            demanda_prev = (espec * prod_prev) * fator
            
            # 3. Lógica de Lotes (Unindo Anterior e Atual)
            buscas = []
            if op_anterior: buscas.append(clean_id(op_anterior))
            buscas.append(op_alvo)
            
            lotes_total = df_reg[(df_reg['OP_REF'].isin(buscas)) & (df_reg['SKU_REF'] == sku)].copy()
            
            if lotes_total.empty:
                relatorio_final.append({
                    "Código": sku, "Descrição": desc, "Quantidade (Kg)": round(demanda_atual, 2),
                    "Lote": "S/ REGISTRO", "Origem": "Verificar Manual"
                })
            else:
                # Primeiro, "gastamos" a demanda da OP anterior nos lotes
                saldo_prev = demanda_prev
                saldo_atual = demanda_atual
                
                for _, l_row in lotes_total.iterrows():
                    qtd_lote_orig = parse_num(l_row['QUANTIDADE'])
                    lote_id = l_row['LOTE']
                    
                    # Consumindo para a OP anterior
                    if saldo_prev > 0:
                        gasto_prev = min(saldo_prev, qtd_lote_orig)
                        saldo_prev -= gasto_prev
                        qtd_restante_no_lote = qtd_lote_orig - gasto_prev
                    else:
                        qtd_restante_no_lote = qtd_lote_orig
                    
                    # O que sobrou do lote vai para a OP Atual
                    if qtd_restante_no_lote > 0 and saldo_atual > 0:
                        uso_atual = min(saldo_atual, qtd_restante_no_lote)
                        saldo_atual -= uso_atual
                        
                        relatorio_final.append({
                            "Código": sku, "Descrição": desc, 
                            "Quantidade (Kg)": round(uso_atual, 2),
                            "Lote": lote_id, "Origem": f"Entrada {l_row['OP']}"
                        })
                
                # Se ainda faltar quilo para a OP atual após todos os lotes
                if saldo_atual > 0.5:
                    relatorio_final.append({
                        "Código": sku, "Descrição": desc, 
                        "Quantidade (Kg)": round(saldo_atual, 2),
                        "Lote": "PENDENTE", "Origem": "Saldo pendente"
                    })

        df_final = pd.DataFrame(relatorio_final)
        st.subheader(f"Relatório de Consumo Final - OP {op_alvo}")
        st.table(df_final)

        # Download
        output = io.BytesIO()
        df_final.to_excel(output, index=False, engine='xlsxwriter')
        st.download_button("📥 Baixar Excel PCP", output.getvalue(), f"Consumo_Lotes_OP_{op_alvo}.xlsx")

else:
    st.info("Carregue as planilhas para gerar o relatório.")

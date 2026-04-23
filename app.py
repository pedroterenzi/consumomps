import streamlit as st
import pandas as pd
import io

st.set_page_config(page_title="PCP - Consumo por Lote", layout="wide")

# --- FUNÇÕES DE LIMPEZA ---
def clean_id(val):
    if pd.isna(val) or val == "": return ""
    return str(val).strip().split('.')[0].lstrip('0')

def parse_num(val):
    if pd.isna(val) or val == "": return 0.0
    s = str(val).replace('.', '').replace(',', '.')
    try: return float(s)
    except: return 0.0

st.title("🚀 Sistema de Consumo MP por Lote")

# --- SIDEBAR ---
st.sidebar.header("Configurações")
file_oficial = st.sidebar.file_uploader("1. Relatório Oficial (XLSM)", type=["xlsm", "xlsx"])
file_stand = st.sidebar.file_uploader("2. Real x Stand (XLSX)", type=["xlsx"])
file_registros = st.sidebar.file_uploader("3. Controle Requisição (XLSX)", type=["xlsx"])

if file_oficial and file_stand and file_registros:
    # Leitura dos dados
    df_oficial = pd.read_excel(file_oficial, sheet_name='Result by order')
    df_stand = pd.read_excel(file_stand, sheet_name='2-Totais por OP   Produto')
    df_perdas = pd.read_excel(file_stand, sheet_name='Planilha1')
    df_reg = pd.read_excel(file_registros, sheet_name='REGISTROS', skiprows=2)

    # Preparação das chaves
    df_oficial['OP_REF'] = df_oficial['Nº Ordem'].apply(clean_id)
    df_reg['OP_REF'] = df_reg['OP'].apply(clean_id)
    df_reg['SKU_REF'] = df_reg['SKU'].apply(clean_id)
    
    ops_list = sorted(df_oficial['OP_REF'].unique(), reverse=True)
    
    col1, col2 = st.columns(2)
    with col1:
        op_alvo = st.selectbox("Selecione a OP Atual", ops_list)
    with col2:
        op_anterior = st.text_input("OP Anterior (para buscar lotes remanescentes)")

    if st.button("Gerar Relatório de Consumo"):
        # 1. Produção Real (Soma de todos os turnos da OP)
        total_produzido = df_oficial[df_oficial['OP_REF'] == op_alvo]['Machine Counter'].sum()
        
        # 2. Materiais da OP (Busca por contém na Coluna A do Real x Stand)
        materiais = df_stand[df_stand.iloc[:, 0].astype(str).str.contains(op_alvo)].copy()
        
        res_final = []

        for _, row in materiais.iterrows():
            sku = clean_id(row.iloc[4]) # M A T E R I A L CODIGO
            if not sku or not str(row.iloc[4]).isdigit(): continue
            
            # Cálculo da ESPEC (Fórmula: K753 / K750)
            qtd_programada_op = parse_num(row.iloc[3]) # QUANTIDADE da OP (cabeçalho)
            qtd_std_mp = parse_num(row.iloc[11])      # QUANTIDADE S T A N D A R D da MP
            
            espec = qtd_std_mp / qtd_programada_op if qtd_programada_op > 0 else 0
            
            # Fator de Perda (Planilha1)
            fator_row = df_perdas[df_perdas['Código'].apply(clean_id) == sku]
            fator = fator_row['% da espec'].values[0] if not fator_row.empty else 1.0
            
            # CONSUMO REAL CALCULADO (Sua lógica manual)
            consumo_total_kg = (espec * total_produzido) * fator
            
            # 3. BUSCA DE LOTES (OP Atual + OP Anterior)
            # Pegamos entradas da OP atual
            lotes_fiscais = df_reg[(df_reg['OP_REF'] == op_alvo) & (df_reg['SKU_REF'] == sku)].copy()
            
            # Se informou OP anterior, buscamos entradas nela também para compor o saldo
            if op_anterior:
                lotes_ant = df_reg[(df_reg['OP_REF'] == clean_id(op_anterior)) & (df_reg['SKU_REF'] == sku)].copy()
                lotes_fiscais = pd.concat([lotes_ant, lotes_fiscais])

            if lotes_fiscais.empty:
                res_final.append({
                    "Código": sku, "Descrição": row.iloc[5], "Qtd (Kg)": round(consumo_total_kg, 2),
                    "Lote": "NÃO INFORMADO", "Info": "Sem registro de entrada"
                })
            else:
                saldo_necessario = consumo_total_kg
                # Itera sobre os lotes para distribuir a quantidade
                for _, l_row in lotes_fiscais.iterrows():
                    if saldo_necessario <= 0: break
                    
                    qtd_no_lote = parse_num(l_row['QUANTIDADE'])
                    lote_id = l_row['LOTE']
                    
                    # Usa o que for menor: o que tem no lote ou o que falta consumir
                    consumo_neste_lote = min(saldo_necessario, qtd_no_lote)
                    saldo_necessario -= consumo_neste_lote
                    
                    res_final.append({
                        "Código": sku, "Descrição": row.iloc[5], "Qtd (Kg)": round(consumo_neste_lote, 2),
                        "Lote": lote_id, "Info": f"OP {l_row['OP']}"
                    })
                
                # Se ainda sobrar consumo sem lote
                if saldo_necessario > 0.5:
                    res_final.append({
                        "Código": sku, "Descrição": row.iloc[5], "Qtd (Kg)": round(saldo_necessario, 2),
                        "Lote": "SALDO PÉ DE MÁQUINA", "Info": "Residual de ordens passadas"
                    })

        df_res = pd.DataFrame(res_final)
        st.subheader(f"Relatório de Consumo - OP {op_alvo}")
        st.write(f"**Produção Real Total:** {total_produzido:,.0f} peças")
        st.table(df_res)

        # Download Excel
        towrite = io.BytesIO()
        df_res.to_excel(towrite, index=False, engine='xlsxwriter')
        st.download_button("📥 Baixar Relatório", towrite.getvalue(), f"Consumo_OP_{op_alvo}.xlsx")

else:
    st.info("Carregue as planilhas para processar os lotes.")

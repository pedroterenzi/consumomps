import streamlit as st
import pandas as pd
import io

# Configuração da Página
st.set_page_config(page_title="PCP - Rastreabilidade de Consumo", layout="wide")

st.title("🚀 Sistema de Consumo de MP por Lote")
st.markdown("---")

# --- FUNÇÕES DE UTILIDADE ---

def clean_id(val):
    """ Limpa o ID removendo zeros à esquerda para comparação de SKUs """
    if pd.isna(val) or val == "": return ""
    return str(val).strip().split('.')[0].lstrip('0')

def parse_number(val):
    """ Converte formatos brasileiros '1.100,50' para float """
    if pd.isna(val) or val == "": return 0.0
    s = str(val).strip()
    if "," in s and "." in s:
        s = s.replace(".", "").replace(",", ".")
    elif "," in s:
        s = s.replace(",", ".")
    try:
        return float(s)
    except:
        return 0.0

# --- SIDEBAR ---
st.sidebar.header("1. Carregar Planilhas")
file_oficial = st.sidebar.file_uploader("Relatório Oficial (Excel)", type=["xlsm", "xlsx"])
file_stand = st.sidebar.file_uploader("Real x Stand (Excel)", type=["xlsx"])
file_registros = st.sidebar.file_uploader("Controle de Requisição (Excel)", type=["xlsx"])

if file_oficial and file_stand and file_registros:
    try:
        with st.spinner('Lendo planilhas...'):
            df_oficial = pd.read_excel(file_oficial, sheet_name='Result by order')
            df_stand = pd.read_excel(file_stand, sheet_name='2-Totais por OP   Produto')
            df_perdas = pd.read_excel(file_stand, sheet_name='Planilha1')
            df_reg = pd.read_excel(file_registros, sheet_name='REGISTROS', skiprows=2)

        # Criamos uma coluna de referência simplificada no Oficial e Registros
        df_oficial['OP_REF'] = df_oficial['Nº Ordem'].apply(clean_id)
        df_reg['OP_REF'] = df_reg['OP'].apply(clean_id)
        df_reg['SKU_REF'] = df_reg['SKU'].apply(clean_id)
        
        # Lista de OPs para o usuário escolher (vinda do Relatório Oficial)
        ops_disponiveis = sorted(df_oficial['OP_REF'].unique(), reverse=True)
        
        st.header("2. Seleção de Ordem de Produção")
        op_alvo = st.selectbox("Selecione a OP Atual", ops_disponiveis)

        if st.button("📊 Gerar Relatório Detalhado"):
            # 1. Produção Real
            prod_real = df_oficial[df_oficial['OP_REF'] == op_alvo]['Machine Counter'].sum()
            
            # 2. Localizar Materiais na Real x Stand (Busca por texto contido)
            # Procuramos na Coluna A da Real x Stand se o número da OP aparece lá
            materiais_op = df_stand[df_stand.iloc[:, 0].astype(str).str.contains(op_alvo)].copy()
            
            if materiais_op.empty:
                st.error(f"Não encontrei a OP {op_alvo} na planilha Real x Stand. Verifique se o número está correto.")
            else:
                st.success(f"Produção Real: {prod_real:,.0f} peças")
                relatorio_dados = []

                for _, row in materiais_op.iterrows():
                    sku_raw = row.iloc[4] # M A T E R I A L CODIGO
                    sku_clean = clean_id(sku_raw)
                    if not sku_clean or not str(sku_raw).isdigit(): continue
                    
                    desc = row.iloc[5]
                    
                    # Cálculo da Spec
                    qtd_prog_pecas = parse_number(row.iloc[3]) # Coluna D
                    kg_std_total = parse_number(row.iloc[11]) # Coluna L (C O N S U M O QUANTIDADE)
                    
                    spec_base = kg_std_total / qtd_prog_pecas if qtd_prog_pecas > 0 else 0
                    
                    # Fator de Perda
                    fator_row = df_perdas[df_perdas['Código'].apply(clean_id) == sku_clean]
                    fator = fator_row['% da espec'].values[0] if not fator_row.empty else 1.0
                    
                    consumo_alvo = (spec_base * prod_real) * fator
                    
                    # 3. Cruzamento com Lotes
                    lotes_op = df_reg[(df_reg['OP_REF'] == op_alvo) & (df_reg['SKU_REF'] == sku_clean)].copy()
                    
                    if lotes_op.empty:
                        relatorio_dados.append({
                            "Código": sku_clean, "Descrição": desc, "Qtd (Kg)": round(consumo_alvo, 2),
                            "Lote": "N/A", "Origem": "Verificar OP Anterior"
                        })
                    else:
                        restante = consumo_alvo
                        for _, lote_row in lotes_op.iterrows():
                            if restante <= 0: break
                            qtd_lote = parse_number(lote_row['QUANTIDADE'])
                            lote_id = lote_row['LOTE']
                            
                            usado = min(restante, qtd_lote)
                            restante -= usado
                            
                            relatorio_dados.append({
                                "Código": sku_clean, "Descrição": desc, "Qtd (Kg)": round(usado, 2),
                                "Lote": lote_id, "Origem": "Entrada na OP"
                            })
                        
                        if restante > 0.1:
                            relatorio_dados.append({
                                "Código": sku_clean, "Descrição": desc, "Qtd (Kg)": round(restante, 2),
                                "Lote": "SALDO ANTERIOR", "Origem": "Pé de Máquina"
                            })

                df_final = pd.DataFrame(relatorio_dados)
                st.subheader("📋 Relatório de Consumo")
                st.table(df_final)

                # Download
                output = io.BytesIO()
                with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                    df_final.to_excel(writer, index=False, sheet_name='Consumo')
                st.download_button("📥 Baixar Excel", output.getvalue(), f"Consumo_OP_{op_alvo}.xlsx")

    except Exception as e:
        st.error(f"Erro: {e}")
else:
    st.info("Aguardando planilhas...")

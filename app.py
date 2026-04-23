import streamlit as st
import pandas as pd
import io

# Configuração da Página
st.set_page_config(page_title="PCP - Rastreabilidade de Consumo", layout="wide")

st.title("🚀 Sistema de Consumo de MP por Lote")
st.markdown("---")

# --- FUNÇÕES DE UTILIDADE (LIMPEZA DE DADOS) ---

def clean_id(val):
    """ Remove zeros à esquerda, espaços e converte 18534.0 para '18534' """
    if pd.isna(val) or val == "": return ""
    return str(val).strip().split('.')[0].lstrip('0')

def parse_number(val):
    """ Converte strings BR (1.200,50) ou formatos mistos para float """
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

# --- SIDEBAR: UPLOADS ---
st.sidebar.header("1. Carregar Planilhas")
file_oficial = st.sidebar.file_uploader("Relatório Oficial (Excel)", type=["xlsm", "xlsx"])
file_stand = st.sidebar.file_uploader("Real x Stand (Excel)", type=["xlsx"])
file_registros = st.sidebar.file_uploader("Controle de Requisição (Excel)", type=["xlsx"])

# --- PROCESSAMENTO PRINCIPAL ---

if file_oficial and file_stand and file_registros:
    try:
        # Carregando os dados
        with st.spinner('Lendo planilhas...'):
            df_oficial = pd.read_excel(file_oficial, sheet_name='Result by order')
            df_stand = pd.read_excel(file_stand, sheet_name='2-Totais por OP   Produto')
            df_perdas = pd.read_excel(file_stand, sheet_name='Planilha1')
            df_reg = pd.read_excel(file_registros, sheet_name='REGISTROS', skiprows=2)

        # Pré-processamento de Colunas de Ligação
        df_oficial['OP_REF'] = df_oficial['Nº Ordem'].apply(clean_id)
        df_reg['OP_REF'] = df_reg['OP'].apply(clean_id)
        df_reg['SKU_REF'] = df_reg['SKU'].apply(clean_id)
        df_stand['OP_REF'] = df_stand.iloc[:, 0].apply(clean_id) # Coluna A
        
        # Filtros de Seleção
        ops_disponiveis = sorted(df_oficial['OP_REF'].unique(), reverse=True)
        
        st.header("2. Seleção de Ordem de Produção")
        col_op, col_ant = st.columns(2)
        with col_op:
            op_alvo = st.selectbox("Selecione a OP Atual", ops_disponiveis)
        with col_ant:
            st.info("A lógica de 'Saldo Anterior' será aplicada se os lotes da OP atual não suprirem o consumo.")

        if st.button("📊 Gerar Relatório Detalhado"):
            # 1. Obter Produção Real (Machine Counter)
            prod_real = df_oficial[df_oficial['OP_REF'] == op_alvo]['Machine Counter'].sum()
            
            if prod_real == 0:
                st.error(f"Produção não encontrada para a OP {op_alvo}. Verifique o Relatório Oficial.")
            else:
                st.success(f"Produção Real: {prod_real:,.0f} peças")
                
                # 2. Materiais da OP no Real x Stand
                materiais_op = df_stand[df_stand['OP_REF'] == op_alvo].copy()
                relatorio_dados = []

                for _, row in materiais_op.iterrows():
                    sku = clean_id(row.iloc[4]) # Coluna M A T E R I A L CODIGO
                    if not sku: continue
                    
                    desc = row.iloc[5]
                    
                    # Cálculo da Spec Base
                    qtd_prog_pecas = parse_number(row.iloc[3])
                    kg_std_total = parse_number(row.iloc[11])
                    spec_base = kg_std_total / qtd_prog_pecas if qtd_prog_pecas > 0 else 0
                    
                    # Fator de Perda (Planilha1)
                    fator_row = df_perdas[df_perdas['Código'].apply(clean_id) == sku]
                    fator = fator_row['% da espec'].values[0] if not fator_row.empty else 1.0
                    
                    # Meta de Consumo para esta OP
                    consumo_alvo = (spec_base * prod_real) * fator
                    
                    # 3. Cruzamento com Lotes (Lógica de Abatimento)
                    lotes_op = df_reg[(df_reg['OP_REF'] == op_alvo) & (df_reg['SKU_REF'] == sku)].copy()
                    
                    if lotes_op.empty:
                        relatorio_dados.append({
                            "Código": sku, "Descrição": desc, "Qtd (Kg)": round(consumo_alvo, 2),
                            "Lote": "N/A", "Origem": "Verificar OP Anterior / Pé de Máquina"
                        })
                    else:
                        restante_consumo = consumo_alvo
                        for _, lote_row in lotes_op.iterrows():
                            if restante_consumo <= 0: break
                            
                            qtd_lote = parse_number(lote_row['QUANTIDADE'])
                            lote_id = lote_row['LOTE']
                            
                            if restante_consumo >= qtd_lote:
                                usado = qtd_lote
                                restante_consumo -= qtd_lote
                            else:
                                usado = restante_consumo
                                restante_consumo = 0
                            
                            relatorio_dados.append({
                                "Código": sku, "Descrição": desc, "Qtd (Kg)": round(usado, 2),
                                "Lote": lote_id, "Origem": "Entrada na OP Atual"
                            })
                        
                        # Se após todos os lotes registrados ainda faltar consumo
                        if restante_consumo > 0.1:
                            relatorio_dados.append({
                                "Código": sku, "Descrição": desc, "Qtd (Kg)": round(restante_consumo, 2),
                                "Lote": "SALDO ANTERIOR", "Origem": "Consumido de sobra da máquina"
                            })

                # Exibição do Resultado
                df_final = pd.DataFrame(relatorio_dados)
                st.subheader("📋 Consumo de Matéria-Prima por Lote")
                st.dataframe(df_final, use_container_width=True)

                # Opção de Download
                output = io.BytesIO()
                with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                    df_final.to_excel(writer, index=False, sheet_name='Consumo_PCP')
                
                st.download_button(
                    label="📥 Baixar Relatório PCP (Excel)",
                    data=output.getvalue(),
                    file_name=f"Relatorio_Consumo_OP_{op_alvo}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )

    except Exception as e:
        st.error(f"Erro crítico no processamento: {e}")
        st.info("Dica: Verifique se as abas 'Result by order', '2-Totais por OP Produto', 'Planilha1' e 'REGISTROS' existem com esses nomes exatos.")

else:
    st.info("👆 Por favor, carregue as três planilhas na barra lateral para começar.")

import streamlit as st
import pandas as pd
import io

# Configuração da página para usar a largura total
st.set_page_config(page_title="PCP - Rastreabilidade de Consumo", layout="wide")

st.title("🚀 Sistema de Consumo de MP por Lote (Lógica FIFO)")
st.markdown("""
Este sistema calcula o consumo real baseado na produção da máquina e distribui esse consumo entre os lotes registrados, 
considerando o saldo que sobrou da OP anterior.
""")

# --- FUNÇÕES DE TRATAMENTO DE DADOS ---

def clean_id(val):
    """Limpa OPs e SKUs: remove zeros à esquerda, espaços e .0"""
    if pd.isna(val) or val == "": return ""
    return str(val).strip().split('.')[0].lstrip('0')

def parse_num(val):
    """Converte '1.100,50' ou '1,100.50' ou 1100.5 para float de forma segura"""
    if pd.isna(val) or val == "" or str(val).strip() == "-": return 0.0
    s = str(val).strip()
    
    # Se houver vírgula e ponto, assumimos formato BR (1.200,50) -> remove ponto, troca vírgula
    if "," in s and "." in s:
        s = s.replace(".", "").replace(",", ".")
    # Se houver apenas vírgula (1200,50) -> troca por ponto
    elif "," in s:
        s = s.replace(",", ".")
    # Se houver múltiplos pontos e nenhuma vírgula (1.100.500) -> remove pontos
    elif s.count(".") > 1:
        s = s.replace(".", "")
        
    try:
        return float(s)
    except:
        return 0.0

# --- SIDEBAR - UPLOAD ---
st.sidebar.header("📂 Upload das Planilhas")
f_oficial = st.sidebar.file_uploader("1. Relatório Oficial (Peças Reais)", type=["xlsm", "xlsx"])
f_stand = st.sidebar.file_uploader("2. Real x Stand (Specs e Perdas)", type=["xlsx"])
f_reg = st.sidebar.file_uploader("3. Controle Requisição (Entrada de Lotes)", type=["xlsx"])

if f_oficial and f_stand and f_reg:
    try:
        with st.spinner('Lendo e higienizando dados...'):
            # Leitura
            df_oficial = pd.read_excel(f_oficial, sheet_name='Result by order')
            df_stand = pd.read_excel(f_stand, sheet_name='2-Totais por OP   Produto')
            df_perdas = pd.read_excel(f_stand, sheet_name='Planilha1')
            df_reg = pd.read_excel(f_reg, sheet_name='REGISTROS', skiprows=2)

            # Padronização de IDs
            df_oficial['OP_REF'] = df_oficial['Nº Ordem'].apply(clean_id)
            df_reg['OP_REF'] = df_reg['OP'].apply(clean_id)
            df_reg['SKU_REF'] = df_reg['SKU'].apply(clean_id)
            df_stand['OP_REF'] = df_stand.iloc[:, 0].apply(clean_id) # Coluna A da aba Totais

        # --- SELEÇÃO DE OP ---
        st.header("⚙️ Parâmetros de Processamento")
        col_a, col_b = st.columns(2)
        
        with col_a:
            ops_disponiveis = sorted(df_oficial['OP_REF'].unique(), reverse=True)
            op_atual = st.selectbox("Selecione a OP Alvo (ex: 18940)", ops_disponiveis)
        
        with col_b:
            op_anterior = st.text_input("Informe a OP Anterior da Máquina (ex: 18938)", help="Essencial para calcular o que sobrou na máquina")

        if st.button("📊 Gerar Relatório de Consumo por Lote"):
            # 1. Obter Produções Reais
            prod_real_atual = df_oficial[df_oficial['OP_REF'] == op_atual]['Machine Counter'].sum()
            prod_real_prev = 0
            if op_anterior:
                prod_real_prev = df_oficial[df_oficial['OP_REF'] == clean_id(op_anterior)]['Machine Counter'].sum()

            # 2. Filtrar Materiais da OP Alvo
            materiais_f = df_stand[df_stand['OP_REF'] == op_atual].copy()
            
            if materiais_f.empty:
                st.error(f"OP {op_atual} não encontrada na aba '2-Totais por OP Produto'.")
            else:
                st.success(f"Produção Real: {prod_real_atual:,.0f} peças.")
                
                lista_pcp = []

                for _, row in materiais_f.iterrows():
                    sku = clean_id(row.iloc[4]) # M A T E R I A L CODIGO
                    if not sku or not str(row.iloc[4]).isdigit(): continue
                    
                    desc = row.iloc[5]
                    
                    # CÁLCULO DA ESPEC (Fórmula: MP_Std_Qty / OP_Prog_Qty)
                    prog_pcs_op = parse_num(row.iloc[3])
                    std_kg_mp = parse_num(row.iloc[11])
                    espec = std_kg_mp / prog_pcs_op if prog_pcs_op > 0 else 0
                    
                    # FATOR DE PERDA
                    f_row = df_perdas[df_perdas['Código'].apply(clean_id) == sku]
                    fator = parse_num(f_row['% da espec'].values[0]) if not f_row.empty else 1.0
                    
                    # DEMANDA (Kg) das duas OPs
                    necessario_atual = (espec * prod_real_atual) * fator
                    necessario_prev = (espec * prod_real_prev) * fator
                    
                    # BUSCA DE LOTES (Anterior + Atual)
                    ops_busca = [clean_id(op_anterior), op_atual] if op_anterior else [op_atual]
                    lotes_db = df_reg[(df_reg['OP_REF'].isin(ops_busca)) & (df_reg['SKU_REF'] == sku)].copy()
                    
                    if lotes_db.empty:
                        lista_pcp.append({
                            "Código": sku, "Descrição": desc, "Quantidade (Kg)": round(necessario_atual, 3),
                            "Lote": "S/ REGISTRO", "Origem": "Verificar Manual"
                        })
                    else:
                        # LÓGICA FIFO: Abatendo consumo da anterior primeiro
                        reserva_prev = necessario_prev
                        reserva_atual = necessario_atual
                        
                        for _, l_row in lotes_db.iterrows():
                            qtd_pallet = parse_num(l_row['QUANTIDADE'])
                            lote_id = l_row['LOTE']
                            
                            # Consome primeiro para a anterior (Herança)
                            if reserva_prev > 0:
                                gasto_prev = min(reserva_prev, qtd_pallet)
                                reserva_prev -= gasto_prev
                                qtd_pallet -= gasto_prev
                            
                            # O que sobrou do pallet vai para a atual
                            if qtd_pallet > 0 and reserva_atual > 0:
                                gasto_atual = min(reserva_atual, qtd_pallet)
                                reserva_atual -= gasto_atual
                                
                                lista_pcp.append({
                                    "Código": sku, "Descrição": desc, 
                                    "Quantidade (Kg)": round(gasto_atual, 3),
                                    "Lote": lote_id, "Origem": f"OP {l_row['OP']}"
                                })

                        # Se ainda faltar consumo (Material já estava lá de antes da anterior)
                        if reserva_atual > 0.5:
                            lista_pcp.append({
                                "Código": sku, "Descrição": desc, 
                                "Quantidade (Kg)": round(reserva_atual, 3),
                                "Lote": "SALDO MÁQUINA", "Origem": "Ordens Antigas"
                            })

                # Exibição Final
                df_final = pd.DataFrame(lista_pcp)
                st.subheader(f"📋 Relatório de Consumo Final - OP {op_atual}")
                st.table(df_final)

                # Exportação
                buffer = io.BytesIO()
                with pd.ExcelWriter(buffer, engine='xlsxwriter') as writer:
                    df_final.to_excel(writer, index=False, sheet_name='PCP')
                st.download_button(
                    label="📥 Baixar Relatório Excel",
                    data=buffer.getvalue(),
                    file_name=f"Consumo_PCP_OP_{op_atual}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )

    except Exception as e:
        st.error(f"Erro ao processar: {e}")
        st.info("Dica: Verifique se os nomes das colunas e abas estão corretos.")

else:
    st.info("Aguardando upload das 3 planilhas para iniciar.")

import streamlit as st
import pandas as pd
import io

st.set_page_config(page_title="PCP - Rastreabilidade de Consumo", layout="wide")

st.title("🚀 Sistema de Consumo de MP por Lote")
st.markdown("Cruzamento automático de Produção Real, Especificações e Lotes.")

# --- SIDEBAR: UPLOAD DE ARQUIVOS ---
st.sidebar.header("1. Upload de Planilhas")
file_oficial = st.sidebar.file_uploader("Relatório Oficial (Excel/XLSM)", type=["xlsm", "xlsx"])
file_stand = st.sidebar.file_uploader("Real x Stand (Excel/CSV)", type=["xlsx", "csv"])
file_registros = st.sidebar.file_uploader("Controle de Requisição (Excel/CSV)", type=["xlsx", "csv"])

# --- FUNÇÕES DE PROCESSAMENTO ---
def load_data():
    try:
        # Carregar Produção Real
        df_oficial = pd.read_excel(file_oficial, sheet_name='Result by order')
        
        # Carregar Specs (ajustando para CSV ou Excel)
        if file_stand.name.endswith('.csv'):
            df_stand = pd.read_csv(file_stand)
        else:
            df_stand = pd.read_excel(file_stand, sheet_name='2-Totais por OP   Produto')
            
        # Carregar Fatores de Perda (Planilha1)
        if file_stand.name.endswith('.csv'):
             df_perdas = pd.read_csv('real x stand novo.xlsx - Planilha1.csv') # Fallback para o arquivo enviado
        else:
            df_perdas = pd.read_excel(file_stand, sheet_name='Planilha1')

        # Carregar Registros de Lotes
        if file_registros.name.endswith('.csv'):
            df_reg = pd.read_csv(file_registros, skiprows=2)
            df_reg.columns = ['DATA', 'TURNO', 'OP', 'DOC', 'SKU', 'LOTE', 'QUANTIDADE', 'CHAVE']
        else:
            df_reg = pd.read_excel(file_registros, sheet_name='REGISTROS', skiprows=2)
            
        return df_oficial, df_stand, df_perdas, df_reg
    except Exception as e:
        st.error(f"Erro ao carregar arquivos: {e}")
        return None, None, None, None

if file_oficial and file_stand and file_registros:
    df_oficial, df_stand, df_perdas, df_reg = load_data()
    
    if df_oficial is not None:
        # --- INPUTS DE OP ---
        st.header("2. Parâmetros da OP")
        col1, col2 = st.columns(2)
        with col1:
            op_alvo = st.selectbox("Selecione a OP Atual", sorted(df_oficial['Nº Ordem'].unique(), reverse=True))
        with col2:
            op_anterior = st.number_input("Informe a OP Anterior da Máquina (para saldo de lote)", value=0)

        if st.button("Gerar Relatório de Consumo"):
            # 1. Produção Total da OP
            prod_real = df_oficial[df_oficial['Nº Ordem'] == op_alvo]['Machine Counter'].sum()
            st.info(f"Produção Real Detectada: **{prod_real:,.0f} peças**")

            # 2. Processar Materiais
            # Nota: Ajustar nomes de colunas conforme a inspeção que fizemos
            df_perdas['Código'] = df_perdas['Código'].astype(str)
            
            # Filtro da OP no Real x Stand
            op_mask = df_stand.iloc[:, 0].astype(str).str.contains(str(int(op_alvo)))
            materiais_op = df_stand[op_mask].copy()

            final_rows = []
            for _, row in materiais_op.iterrows():
                sku = str(row.iloc[4]).strip() # Coluna M A T E R I A L CODIGO
                if sku == 'nan' or not sku.isdigit(): continue
                
                desc = row.iloc[5]
                prog_pecas = float(str(row.iloc[3]).replace(',', '.')) if row.iloc[3] != 0 else 0
                total_kg_std = float(str(row.iloc[11]).replace(',', '.')) if row.iloc[11] != 0 else 0
                
                spec_kg_pç = total_kg_std / prog_pecas if prog_pecas > 0 else 0
                
                # Fator de perda
                fator = df_perdas[df_perdas['Código'] == sku]['% da espec'].values
                fator_perda = fator[0] if len(fator) > 0 else 1.0
                
                qtd_calculada = (spec_kg_pç * prod_real) * fator_perda
                
                # Busca Lote (Aba Registros)
                lotes_info = df_reg[(df_reg['OP'].astype(str).str.contains(str(int(op_alvo)))) & 
                                    (df_reg['SKU'].astype(str).str.contains(sku))]
                
                lote_str = ", ".join(lotes_info['LOTE'].astype(str).unique())
                if not lote_str or lote_str == "nan":
                    lote_str = "Verificar OP Anterior"

                final_rows.append({
                    "OP": int(op_alvo),
                    "Código": sku,
                    "Descrição": desc,
                    "Qtd (Kg)": round(qtd_calculada, 2),
                    "Lote": lote_str
                })

            df_final = pd.DataFrame(final_rows)
            st.subheader("Relatório Gerado")
            st.dataframe(df_final, use_container_width=True)

            # --- DOWNLOAD ---
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                df_final.to_excel(writer, index=False, sheet_name='Consumo_PCP')
            
            st.download_button(
                label="📥 Baixar Relatório em Excel",
                data=output.getvalue(),
                file_name=f"Consumo_OP_{op_alvo}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

else:
    st.warning("Aguardando upload de todas as planilhas para iniciar...")

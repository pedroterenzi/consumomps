import streamlit as st
import pandas as pd
import io
import os

# Configuração da Página
st.set_page_config(page_title="PCP Elite - Gestão de Consumo", layout="wide")

# --- CAMINHOS FIXOS DAS PLANILHAS ---
CAMINHOS = {
    "oficial": r"X:\10_Jaguariuna\1.Production\2.(Historico de produtividade)\1.(Produtividade)\Relatório de produção\Relatório das paradas Operacionais(Oficial).xlsm",
    "spec": r"C:\Users\pedro-santos\Downloads\SPEC.xlsx",
    "stand": r"C:\Users\pedro-santos\Documents\real x stand novo.xlsx",
    "requisicao": r"C:\Users\pedro-santos\Downloads\Controle requisição x OP (13).xlsx"
}

# --- FUNÇÕES DE HIGIENE ---
def clean_id(val):
    if pd.isna(val) or val == "": return ""
    return str(val).strip().split('.')[0].lstrip('0')

def parse_num(val):
    if pd.isna(val) or val == "" or str(val).strip() == "-": return 0.0
    s = str(val).strip().replace(',', '.')
    try:
        if s.count('.') > 1:
            parts = s.split('.')
            s = "".join(parts[:-1]) + "." + parts[-1]
        return float(s)
    except: return 0.0

def get_consolidated_specs(df_spec, parent_code):
    parent_clean = clean_id(parent_code)
    first_level = df_spec[df_spec['G1_COD'].apply(clean_id) == parent_clean]
    materiais_unitarios = {}
    scale_factor = 1.0
    
    for _, row in first_level.iterrows():
        comp_code = clean_id(row['G1_COMP'])
        qty = parse_num(row['G1_QUANT'])
        if (not comp_code.startswith(('5', '6'))) and qty > 1:
            scale_factor = qty
            sub_items = df_spec[df_spec['G1_COD'].apply(clean_id) == comp_code]
            for _, sub_row in sub_items.iterrows():
                s_sku = clean_id(sub_row['G1_COMP'])
                if s_sku.startswith(('5', '6')):
                    s_qty = parse_num(sub_row['G1_QUANT'])
                    s_desc = sub_row['DESC_COMP'] if 'DESC_COMP' in df_spec.columns else "MP"
                    materiais_unitarios[s_sku] = {'ratio': s_qty, 'desc': s_desc}
            break

    for _, row in first_level.iterrows():
        comp_code = clean_id(row['G1_COMP'])
        if comp_code.startswith(('5', '6')):
            qty = parse_num(row['G1_QUANT'])
            desc = row['DESC_COMP'] if 'DESC_COMP' in df_spec.columns else "MP"
            ratio_unitario = qty / scale_factor if scale_factor > 0 else qty
            if comp_code not in materiais_unitarios:
                materiais_unitarios[comp_code] = {'ratio': ratio_unitario, 'desc': desc}
                
    return materiais_unitarios, scale_factor

st.title("📊 Gestão de Consumo de MP por OP")
st.markdown("---")

# --- CARREGAMENTO AUTOMÁTICO ---
@st.cache_data(show_spinner="Carregando bases de dados locais...")
def carregar_dados_locais():
    # Verifica se os arquivos existem
    for nome, path in CAMINHOS.items():
        if not os.path.exists(path):
            return None, f"Arquivo não encontrado: {path}"
    
    try:
        df_oficial = pd.read_excel(CAMINHOS["oficial"], sheet_name='Result by order')
        df_skus = pd.read_excel(CAMINHOS["oficial"], sheet_name='Dados SKUs')
        df_spec = pd.read_excel(CAMINHOS["spec"])
        df_perdas = pd.read_excel(CAMINHOS["stand"], sheet_name='Planilha1')
        df_reg = pd.read_excel(CAMINHOS["requisicao"], sheet_name='REGISTROS', skiprows=2)
        return (df_oficial, df_skus, df_spec, df_perdas, df_reg), None
    except Exception as e:
        return None, f"Erro ao ler arquivos: {e}"

dados, erro = carregar_dados_locais()

if erro:
    st.error(erro)
    st.info("Certifique-se de que as pastas estão acessíveis e o servidor X: está mapeado.")
else:
    df_oficial, df_skus, df_spec, df_perdas, df_reg = dados

    df_oficial['OP_REF'] = df_oficial['Nº Ordem'].apply(clean_id)
    df_reg['OP_REF'] = df_reg['OP'].apply(clean_id)
    df_reg['SKU_REF'] = df_reg['SKU'].apply(clean_id)
    
    # Seleção de OP
    ops_disponiveis = sorted(df_oficial['OP_REF'].unique(), reverse=True)
    
    col1, col2 = st.columns(2)
    with col1:
        op_alvo = st.selectbox("Selecione a OP Atual", ops_disponiveis)
    with col2:
        op_anterior = st.text_input("Informe a OP Anterior (Saldo)")

    if st.button("🚀 Gerar Relatório"):
        # 1. Dados da OP
        dados_op = df_oficial[df_oficial['OP_REF'] == op_alvo]
        prod_real_bruta = dados_op['Machine Counter'].sum()
        prod_estoque_real = dados_op['Peças Estoque - Ajuste'].sum()
        cod_pai = clean_id(dados_op['Código'].iloc[0])
        
        # BOM e Escala
        materiais, escala_spec = get_consolidated_specs(df_spec, cod_pai)
        sku_info = df_skus[df_skus.iloc[:, 0].apply(clean_id) == cod_pai]
        fardo_estoque = parse_num(sku_info.iloc[0, 3]) if not sku_info.empty else escala_spec

        pre_report = []

        for sku, info in materiais.items():
            if "MOD" in sku or not sku.isdigit(): continue
            
            # Cálculo da Meta
            if sku.startswith('5905'):
                consumo_meta = (prod_estoque_real / fardo_estoque) * 1.04
                if op_anterior:
                    dados_ant = df_oficial[df_oficial['OP_REF'] == clean_id(op_anterior)]
                    prod_est_ant = dados_ant['Peças Estoque - Ajuste'].sum()
                    consumo_prev_meta = (prod_est_ant / fardo_estoque) * 1.04
                else: consumo_prev_meta = 0
            else:
                f_row = df_perdas[df_perdas['Código'].apply(clean_id) == sku]
                fator = parse_num(f_row['% da espec'].values[0]) if not f_row.empty else 1.0
                consumo_meta = (info['ratio'] * prod_real_bruta) * fator
                if op_anterior:
                    prod_ant_bruta = df_oficial[df_oficial['OP_REF'] == clean_id(op_anterior)]['Machine Counter'].sum()
                    consumo_prev_meta = (info['ratio'] * prod_ant_bruta) * fator
                else: consumo_prev_meta = 0
            
            # Lotes
            ops_busca = [clean_id(op_anterior), op_alvo] if op_anterior else [op_alvo]
            lotes_disp = df_reg[(df_reg['OP_REF'].isin(ops_busca)) & (df_reg['SKU_REF'] == sku)].copy()
            
            saldo_a_preencher = consumo_meta
            reserva_ant = consumo_prev_meta
            
            if lotes_disp.empty:
                pre_report.append({"OP": op_alvo, "Código": sku, "Descrição": info['desc'], "Quantidade": round(consumo_meta, 3), "Lote": "S/ REGISTRO"})
            else:
                for _, l_row in lotes_disp.iterrows():
                    if saldo_a_preencher <= 0: break
                    qtd_ent = parse_num(l_row['QUANTIDADE'])
                    if reserva_ant > 0:
                        gasto = min(reserva_ant, qtd_ent)
                        reserva_ant -= gasto
                        qtd_ent -= gasto
                    if qtd_ent > 0 and saldo_a_preencher > 0:
                        uso = min(saldo_a_preencher, qtd_ent)
                        saldo_a_preencher -= uso
                        pre_report.append({"OP": clean_id(l_row['OP']), "Código": sku, "Descrição": info['desc'], "Quantidade": uso, "Lote": str(l_row['LOTE']).strip()})
                
                if saldo_a_preencher > 1.0:
                    pre_report.append({"OP": "VERIFICAR", "Código": sku, "Descrição": info['desc'], "Quantidade": round(saldo_a_preencher, 3), "Lote": "FALTA REGISTRO"})

        # --- ORGANIZAÇÃO FINAL DAS COLUNAS ---
        df_pre = pd.DataFrame(pre_report)
        if not df_pre.empty:
            # Agrupar Lotes iguais
            df_final = df_pre.groupby(['OP', 'Código', 'Descrição', 'Lote'], as_index=False).agg({'Quantidade': 'sum'})
            df_final['Quantidade'] = df_final['Quantidade'].round(3)
            
            # Ordenar colunas conforme solicitado
            df_final = df_final[['OP', 'Código', 'Descrição', 'Quantidade', 'Lote']]
            
            st.subheader(f"Relatório Final - OP {op_alvo}")
            st.table(df_final)
            
            buffer = io.BytesIO()
            df_final.to_excel(buffer, index=False)
            st.download_button("📥 Baixar Excel PCP", buffer.getvalue(), f"Consumo_PCP_OP_{op_alvo}.xlsx")

st.sidebar.markdown("---")
if st.sidebar.button("🔄 Atualizar Dados"):
    st.cache_data.clear()
    st.rerun()

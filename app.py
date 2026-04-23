import streamlit as st
import pandas as pd
import io
import os
from pathlib import Path

# Configuração da Página
st.set_page_config(page_title="PCP Elite - Gestão de Consumo", layout="wide")

# --- CONFIGURAÇÃO DE CAMINHOS (Windows Pathlib) ---
# Usamos o 'r' antes das aspas para o Windows não confundir as barras
user_path = Path(r"C:/Users/pedro-santos")
downloads_path = user_path / "Downloads"
documents_path = user_path / "Documents"

CAMINHOS_LOCAIS = {
    "spec": downloads_path / "SPEC.xlsx",
    "stand": documents_path / "real x stand novo.xlsx",
    "requisicao": downloads_path / "Controle requisição x OP (13).xlsx"
}

# --- FUNÇÕES DE HIGIENE E TRATAMENTO ---

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
    """ Explode a árvore de materiais e consolida as specs unitárias """
    parent_clean = clean_id(parent_code)
    first_level = df_spec[df_spec['G1_COD'].apply(clean_id) == parent_clean]
    materiais_unitarios = {}
    scale_factor = 1.0
    
    # Passo 1: Achar o fator de escala (Fraldas por pacote)
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

    # Passo 2: Pegar materiais do nível do pacote (Filmes)
    for _, row in first_level.iterrows():
        comp_code = clean_id(row['G1_COMP'])
        if comp_code.startswith(('5', '6')):
            qty = parse_num(row['G1_QUANT'])
            desc = row['DESC_COMP'] if 'DESC_COMP' in df_spec.columns else "MP"
            ratio = qty / scale_factor if scale_factor > 0 else qty
            if comp_code not in materiais_unitarios:
                materiais_unitarios[comp_code] = {'ratio': ratio, 'desc': desc}
                
    return materiais_unitarios, scale_factor

st.title("📊 Gestão de Consumo PCP")
st.markdown("---")

# --- INTERFACE LATERAL (SISTEMA DE FICHEIROS) ---
st.sidebar.header("📁 Estado das Bases de Dados")

arquivos_prontos = True

for chave, caminho in CAMINHOS_LOCAIS.items():
    if not caminho.exists():
        st.sidebar.error(f"❌ Não encontrado: {caminho.name}")
        arquivos_prontos = False
        # Ajuda a encontrar o nome correto se houver erro
        if chave == "spec":
            st.sidebar.info("Sugestões em Downloads:")
            sugestoes = [f.name for f in downloads_path.glob("SPEC*")]
            st.sidebar.write(sugestoes if sugestoes else "Nenhum ficheiro SPEC encontrado")
    else:
        st.sidebar.success(f"✅ {caminho.name} detetado")

# Upload do Oficial (que está na rede X:)
f_oficial = st.sidebar.file_uploader("Suba o Relatório Oficial (Rede X:)", type=["xlsm", "xlsx"])

if arquivos_prontos and f_oficial:
    try:
        with st.spinner('A carregar dados...'):
            df_oficial = pd.read_excel(f_oficial, sheet_name='Result by order')
            df_skus = pd.read_excel(f_oficial, sheet_name='Dados SKUs')
            df_spec = pd.read_excel(CAMINHOS_LOCAIS["spec"])
            df_perdas = pd.read_excel(CAMINHOS_LOCAIS["stand"], sheet_name='Planilha1')
            df_reg = pd.read_excel(CAMINHOS_LOCAIS["requisicao"], sheet_name='REGISTROS', skiprows=2)

        df_oficial['OP_REF'] = df_oficial['Nº Ordem'].apply(clean_id)
        df_reg['OP_REF'] = df_reg['OP'].apply(clean_id)
        df_reg['SKU_REF'] = df_reg['SKU'].apply(clean_id)
        
        ops_disponiveis = sorted(df_oficial['OP_REF'].unique(), reverse=True)
        
        col1, col2 = st.columns(2)
        with col1:
            op_alvo = st.selectbox("Selecione a OP Atual", ops_disponiveis)
        with col2:
            op_anterior = st.text_input("Informe a OP Anterior (para Saldo)")

        if st.button("🚀 Gerar Relatório de Consumo"):
            # 1. Dados da OP
            dados_op = df_oficial[df_oficial['OP_REF'] == op_alvo]
            prod_bruta = dados_op['Machine Counter'].sum()
            prod_estoque = dados_op['Peças Estoque - Ajuste'].sum()
            cod_pai = clean_id(dados_op['Código'].iloc[0])
            
            # Dados de Escala
            materiais, escala_spec = get_consolidated_specs(df_spec, cod_pai)
            sku_info = df_skus[df_skus.iloc[:, 0].apply(clean_id) == cod_pai]
            fardo_estoque = parse_num(sku_info.iloc[0, 3]) if not sku_info.empty else escala_spec

            pre_report = []

            for sku, info in materiais.items():
                if "MOD" in sku or not sku.isdigit(): continue
                
                # --- REGRA ESPECIAL POLYBAG (5905...) ---
                if sku.startswith('5905'):
                    consumo_meta = (prod_estoque / fardo_estoque) * 1.04
                    if op_anterior:
                        d_ant = df_oficial[df_oficial['OP_REF'] == clean_id(op_anterior)]
                        cons_prev = (d_ant['Peças Estoque - Ajuste'].sum() / fardo_estoque) * 1.04
                    else: cons_prev = 0
                else:
                    # REGRA NORMAL
                    f_row = df_perdas[df_perdas['Código'].apply(clean_id) == sku]
                    fator = parse_num(f_row['% da espec'].values[0]) if not f_row.empty else 1.0
                    consumo_meta = (info['ratio'] * prod_bruta) * fator
                    if op_anterior:
                        p_ant = df_oficial[df_oficial['OP_REF'] == clean_id(op_anterior)]['Machine Counter'].sum()
                        cons_prev = (info['ratio'] * p_ant) * fator
                    else: cons_prev = 0
                
                # --- RATEIO DE LOTES ---
                ops_busca = [clean_id(op_anterior), op_alvo] if op_anterior else [op_alvo]
                lotes_disp = df_reg[(df_reg['OP_REF'].isin(ops_busca)) & (df_reg['SKU_REF'] == sku)].copy()
                
                saldo_a_preencher = consumo_meta
                reserva_ant = cons_prev
                
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
                            pre_report.append({
                                "OP": clean_id(l_row['OP']), 
                                "Código": sku, 
                                "Descrição": info['desc'], 
                                "Quantidade": uso, 
                                "Lote": str(l_row['LOTE']).strip()
                            })
                    
                    if saldo_a_preencher > 1.0:
                        pre_report.append({"OP": "VERIFICAR", "Código": sku, "Descrição": info['desc'], "Quantidade": round(saldo_a_preencher, 3), "Lote": "FALTA REGISTRO"})

            # Agrupamento e Exibição Final
            df_pre = pd.DataFrame(pre_report)
            if not df_pre.empty:
                df_final = df_pre.groupby(['OP', 'Código', 'Descrição', 'Lote'], as_index=False).agg({'Quantidade': 'sum'})
                df_final = df_final[['OP', 'Código', 'Descrição', 'Quantidade', 'Lote']]
                st.table(df_final.round(3))
                
                buffer = io.BytesIO()
                df_final.to_excel(buffer, index=False)
                st.download_button("📥 Baixar Excel", buffer.getvalue(), f"PCP_Consumo_OP_{op_alvo}.xlsx")
    except Exception as e:
        st.error(f"Erro ao processar: {e}")

else:
    st.warning("Certifique-se de que os ficheiros locais estão nas pastas corretas e suba o Oficial.")

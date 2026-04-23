import streamlit as st
import pandas as pd
import io

st.set_page_config(page_title="PCP - Consumo Real Total", layout="wide")

# --- FUNÇÕES DE HIGIENE E TRATAMENTO ---

def clean_id(val):
    if pd.isna(val) or val == "": return ""
    return str(val).strip().split('.')[0].lstrip('0')

def parse_num(val):
    """ Converte qualquer formato de número para float de forma segura """
    if pd.isna(val) or val == "" or str(val).strip() == "-": return 0.0
    s = str(val).strip().replace(',', '.')
    try:
        # Se houver mais de um ponto (separador de milhar), remove
        if s.count('.') > 1:
            parts = s.split('.')
            s = "".join(parts[:-1]) + "." + parts[-1]
        return float(s)
    except: return 0.0

def explode_bom(df_spec, parent_code):
    """ Percorre a estrutura de produto e retorna as especificações unitárias """
    items = df_spec[df_spec['G1_COD'].apply(clean_id) == clean_id(parent_code)]
    materials = []
    scale_factor = 1.0 # Padrão: 1 peça por 1 pai
    
    for _, row in items.iterrows():
        comp_code = clean_id(row['G1_COMP'])
        qty = parse_num(row['G1_QUANT'])
        desc = row['DESC_COMP'] if 'DESC_COMP' in df_spec.columns else "MP"
        
        # Identifica o código da fralda (intermediário) para definir o fator de escala
        # Geralmente começa com 3 e tem uma quantidade alta (ex: 88 fraldas por pacote)
        if comp_code.startswith('3') and qty > 1:
            scale_factor = qty
            # Explode a fralda para pegar os componentes internos (polpa, sap, etc)
            sub_items = df_spec[df_spec['G1_COD'].apply(clean_id) == comp_code]
            for _, sub_row in sub_items.iterrows():
                materials.append({
                    'sku': clean_id(sub_row['G1_COMP']),
                    'desc': sub_row['DESC_COMP'] if 'DESC_COMP' in df_spec.columns else "MP",
                    'ratio': parse_num(sub_row['G1_QUANT']) # Espec por fralda
                })
        # Itens que já estão no nível do pacote (ex: polybag, filme externo)
        elif comp_code.isdigit() and (comp_code.startswith('5') or comp_code.startswith('6')):
            # Transforma a espec do pacote em espec por fralda
            materials.append({
                'sku': comp_code,
                'desc': desc,
                'ratio': qty / scale_factor if scale_factor > 0 else qty
            })
            
    return materials, scale_factor

st.title("📊 Relatório de Rastreabilidade PCP")

# --- SIDEBAR ---
st.sidebar.header("Arquivos")
f_oficial = st.sidebar.file_uploader("1. Relatório Oficial", type=["xlsm", "xlsx"])
f_spec = st.sidebar.file_uploader("2. SPEC.xlsx (BOM)", type=["xlsx", "csv"])
f_perdas = st.sidebar.file_uploader("3. Real x Stand (Aba Planilha1)", type=["xlsx"])
f_reg = st.sidebar.file_uploader("4. Controle Requisição", type=["xlsx"])

if f_oficial and f_spec and f_perdas and f_reg:
    with st.spinner('Processando dados...'):
        df_oficial = pd.read_excel(f_oficial, sheet_name='Result by order')
        df_spec = pd.read_excel(f_spec) if f_spec.name.endswith('.xlsx') else pd.read_csv(f_spec)
        df_perdas = pd.read_excel(f_perdas, sheet_name='Planilha1')
        df_reg = pd.read_excel(f_reg, sheet_name='REGISTROS', skiprows=2)

    df_oficial['OP_REF'] = df_oficial['Nº Ordem'].apply(clean_id)
    df_reg['OP_REF'] = df_reg['OP'].apply(clean_id)
    df_reg['SKU_REF'] = df_reg['SKU'].apply(clean_id)
    
    op_alvo = st.selectbox("Selecione a OP Atual", sorted(df_oficial['OP_REF'].unique(), reverse=True))
    op_anterior = st.text_input("Informe a OP Anterior (Saldo)")

    if st.button("🚀 Gerar Consumo por Lote"):
        # 1. Dados da OP e Produção Real
        dados_op = df_oficial[df_oficial['OP_REF'] == op_alvo]
        prod_real_pecas = dados_op['Machine Counter'].sum()
        prod_prev_pecas = df_oficial[df_oficial['OP_REF'] == clean_id(op_anterior)]['Machine Counter'].sum() if op_anterior else 0
        cod_pai = clean_id(dados_op['Código'].iloc[0])

        # 2. Explosão da BOM para obter as specs reais por PEÇA
        lista_materiais, escala = explode_bom(df_spec, cod_pai)
        
        if not lista_materiais:
            st.error(f"Não foram encontrados materiais para o produto {cod_pai}. Verifique o arquivo SPEC.")
        else:
            st.success(f"OP {op_alvo} | Produção: {prod_real_pecas:,.0f} fraldas | Fator: {escala} pçs/pacote")
            
            final_report = []

            for mat in lista_materiais:
                sku = mat['sku']
                if not sku.isdigit(): continue
                
                # Fator de Perda
                f_row = df_perdas[df_perdas['Código'].apply(clean_id) == sku]
                fator = parse_num(f_row['% da espec'].values[0]) if not f_row.empty else 1.0
                
                # Demanda em KG (Spec Unitária x Produção Real x Perda)
                kg_atual = (mat['ratio'] * prod_real_pecas) * fator
                kg_prev = (mat['ratio'] * prod_prev_pecas) * fator
                
                # 3. Lógica FIFO de Lotes
                ops_busca = [clean_id(op_anterior), op_alvo] if op_anterior else [op_alvo]
                lotes_disp = df_reg[(df_reg['OP_REF'].isin(ops_busca)) & (df_reg['SKU_REF'] == sku)].copy()
                
                reserva_ant = kg_prev
                saldo_atual = kg_atual
                
                if lotes_disp.empty:
                    final_report.append({"Cód MP": sku, "Descrição": mat['desc'], "Kg": round(saldo_atual, 3), "Lote": "S/ REGISTRO", "Status": "Verificar Manual"})
                else:
                    for _, l_row in lotes_disp.iterrows():
                        if saldo_atual <= 0: break
                        qtd_pallet = parse_num(l_row['QUANTIDADE'])
                        
                        # Abate o que a anterior "comeu"
                        if reserva_ant > 0:
                            gasto_ant = min(reserva_ant, qtd_pallet)
                            reserva_ant -= gasto_ant
                            qtd_pallet -= gasto_ant
                        
                        # O que sobrou vai para a OP atual
                        if qtd_pallet > 0 and saldo_atual > 0:
                            uso = min(saldo_atual, qtd_pallet)
                            saldo_atual -= uso
                            final_report.append({"Cód MP": sku, "Descrição": mat['desc'], "Kg": round(uso, 3), "Lote": l_row['LOTE'], "Status": f"Entrada {l_row['OP']}"})

                    if saldo_atual > 0.1:
                        final_report.append({"Cód MP": sku, "Descrição": mat['desc'], "Kg": round(saldo_atual, 3), "Lote": "SALDO MÁQUINA", "Status": "Pé de Máquina"})

            df_res = pd.DataFrame(final_report)
            st.table(df_res)
            
            # Exportar Excel
            buffer = io.BytesIO()
            df_res.to_excel(buffer, index=False)
            st.download_button("📥 Baixar Relatório PCP", buffer.getvalue(), f"Consumo_OP_{op_alvo}.xlsx")

else:
    st.info("Carregue os 4 arquivos na barra lateral para iniciar.")

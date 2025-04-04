import streamlit as st
import pandas as pd
import re
import csv
from io import BytesIO

def detectar_colunas(df):
    """Detecção avançada de colunas com tratamento de caracteres especiais"""
    col_map = {'nfe': None, 'pedido': None, 'cliente': None}
    
    nfe_keywords = {'nfe', 'nf', 'notafiscal', 'numero', 'numeronf', 'nº', 'num', 'nota'}
    pedido_keywords = {'pedido', 'numped', 'num_ped', 'numero', 'nº', 'num', 'ped'}
    cliente_keywords = {'cliente', 'nome', 'razaosocial', 'destinatario', 'emitente'}

    for col in df.columns:
        # Normalização completa do nome da coluna
        col_clean = re.sub(r'[º°ª]', ' ', col)  # Substitui símbolos numéricos
        col_clean = re.sub(r'[^a-zA-Z0-9 ]', '', col_clean).lower().strip()
        tokens = set(col_clean.split())
        
        # Prioridade para NFe
        if not col_map['nfe']:
            if 'nfe' in tokens or 'nf' in tokens:
                col_map['nfe'] = col
            elif tokens & nfe_keywords:
                col_map['nfe'] = col
                
        # Detecção de Pedido
        if not col_map['pedido']:
            if 'pedido' in tokens:
                col_map['pedido'] = col
            elif tokens & pedido_keywords:
                col_map['pedido'] = col
                
        # Detecção de Cliente
        if not col_map['cliente'] and (tokens & cliente_keywords):
            col_map['cliente'] = col
    
    return col_map

def read_csv(file):
    """Leitura inteligente de CSV com detecção de delimitador"""
    try:
        content = file.read(1024).decode('utf-8')
        file.seek(0)
        dialect = csv.Sniffer().sniff(content)
        return pd.read_csv(file, delimiter=dialect.delimiter, dtype=str)
    except:
        file.seek(0)
        return pd.read_csv(file, delimiter=';', dtype=str)

# Configuração da interface Streamlit
st.set_page_config(page_title="Comparador de Pedidos", layout="wide")
st.title("📋 Sistema de Comparação Pedidos Enviados")

# Upload de arquivos
with st.expander("🔽 CARREGAR ARQUIVOS", expanded=True):
    col1, col2 = st.columns(2)
    with col1:
        planilha_pedro = st.file_uploader("Planilha Pedro", type=["xlsx", "csv"])
    with col2:
        planilha_dga = st.file_uploader("Planilha DGA", type=["xlsx", "csv"])

if planilha_pedro and planilha_dga:
    try:
        # Leitura dos arquivos
        if planilha_pedro.name.endswith('.xlsx'):
            df_pedro = pd.read_excel(planilha_pedro, dtype=str)
        else:
            df_pedro = read_csv(planilha_pedro)

        if planilha_dga.name.endswith('.xlsx'):
            df_dga = pd.read_excel(planilha_dga, dtype=str)
        else:
            df_dga = read_csv(planilha_dga)

        # Detecção automática de colunas
        cols_pedro = detectar_colunas(df_pedro)
        cols_dga = detectar_colunas(df_dga)

        # Interface para ajuste manual
        with st.expander("⚙ CONFIGURAR COLUNAS", expanded=False):
            st.warning("Verifique e ajuste as colunas detectadas:")

            col1, col2 = st.columns(2)
            with col1:
                st.subheader("Planilha Pedro")
                cols_pedro['nfe'] = st.selectbox(
                    "Coluna NFe Pedro",
                    df_pedro.columns,
                    index=df_pedro.columns.get_loc(cols_pedro['nfe']) if cols_pedro['nfe'] in df_pedro.columns else 0
                )
                cols_pedro['pedido'] = st.selectbox(
                    "Coluna Pedido Pedro",
                    df_pedro.columns,
                    index=df_pedro.columns.get_loc(cols_pedro['pedido']) if cols_pedro['pedido'] in df_pedro.columns else 0
                )

            with col2:
                st.subheader("Planilha DGA")
                cols_dga['nfe'] = st.selectbox(
                    "Coluna NFe DGA",
                    df_dga.columns,
                    index=df_dga.columns.get_loc(cols_dga['nfe']) if cols_dga['nfe'] in df_dga.columns else 0
                )
                cols_dga['pedido'] = st.selectbox(
                    "Coluna Pedido DGA",
                    df_dga.columns,
                    index=df_dga.columns.get_loc(cols_dga['pedido']) if cols_dga['pedido'] in df_dga.columns else 0
                )
                cols_dga['cliente'] = st.selectbox(
                    "Coluna Cliente DGA",
                    df_dga.columns,
                    index=df_dga.columns.get_loc(cols_dga['cliente']) if cols_dga['cliente'] in df_dga.columns else 0
                )

        # Processamento dos dados
        with st.spinner('Processando dados...'):
            # Limpeza e padronização
            for col in [cols_pedro['nfe'], cols_pedro['pedido']]:
                df_pedro[col] = df_pedro[col].astype(str).str.strip().str.replace(r'\D', '', regex=True)
            
            for col in [cols_dga['nfe'], cols_dga['pedido']]:
                df_dga[col] = df_dga[col].astype(str).str.strip().str.replace(r'\D', '', regex=True)

            # Criação de chaves únicas
            df_pedro['chave'] = df_pedro[cols_pedro['nfe']] + '|' + df_pedro[cols_pedro['pedido']]
            df_dga['chave'] = df_dga[cols_dga['nfe']] + '|' + df_dga[cols_dga['pedido']]

            # Merge dos dados
            merged = pd.merge(
                left=df_pedro,
                right=df_dga[[cols_dga['nfe'], cols_dga['pedido'], cols_dga['cliente'], 'chave']],
                on='chave',
                how='inner',
                suffixes=('_Pedro', '_DGA')
            ).drop_duplicates('chave')

            # Criação do dataframe final
            resultado = merged[[
                cols_pedro['nfe'],
                cols_pedro['pedido'],
                cols_dga['cliente']
            ]].copy()
            
            resultado.columns = ['NFe', 'Pedido', 'Cliente']
            resultado['Enviado?'] = True
            resultado['Observação'] = ''
            resultado['Data'] = pd.Timestamp.now().strftime('%d/%m/%Y')

        # Exibição dos resultados
        st.success(f"✅ {len(resultado)} correspondências encontradas")
        
        # Editor de dados
        edited_df = st.data_editor(
            resultado,
            column_config={
                "NFe": "Número da NF-e",
                "Pedido": "Número do Pedido",
                "Cliente": st.column_config.TextColumn(
                    "Cliente",
                    width="large"
                ),
                "Enviado?": st.column_config.CheckboxColumn(
                    "Confirmado",
                    default=True
                ),
                "Data": st.column_config.DateColumn(
                    "Data de Verificação",
                    format="DD/MM/YYYY",
                    disabled=True
                )
            },
            hide_index=True,
            use_container_width=True
        )

        # Exportação
        st.divider()
        if st.button("💾 Gerar Relatório Consolidado"):
            output = BytesIO()
            with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                edited_df.to_excel(writer, index=False, sheet_name='Controle')
                
                # Resumo estatístico
                resumo = pd.DataFrame({
                    'Metrica': ['Total de Pedidos', 'Confirmados', 'Pendentes'],
                    'Valor': [
                        len(edited_df),
                        edited_df['Enviado?'].sum(),
                        len(edited_df) - edited_df['Enviado?'].sum()
                    ]
                })
                resumo.to_excel(writer, index=False, sheet_name='Resumo')
                
                # Formatação
                workbook = writer.book
                format_header = workbook.add_format({'bold': True, 'bg_color': '#5DADE2'})
                
                for sheet in writer.sheets:
                    worksheet = writer.sheets[sheet]
                    worksheet.set_column('A:E', 20)
                    worksheet.autofilter(0, 0, 0, len(edited_df.columns)-1)
                    worksheet.freeze_panes(1, 0)
                    for col_num, value in enumerate(edited_df.columns.values):
                        worksheet.write(0, col_num, value, format_header)

            st.download_button(
                label="⬇️ Baixar Relatório Completo",
                data=output.getvalue(),
                file_name=f"Relatorio_Envios_{pd.Timestamp.now().strftime('%Y%m%d')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

    except Exception as e:
        st.error(f"🚨 Erro no processamento: {str(e)}")
        st.stop()

else:
    st.warning("⚠️ Por favor, carregue ambas as planilhas para iniciar a comparação.")

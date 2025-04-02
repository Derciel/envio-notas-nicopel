import streamlit as st
import pandas as pd
from io import BytesIO

def detectar_colunas(df):
    """Detecta automaticamente colunas relevantes"""
    col_map = {'nfe': None, 'pedido': None, 'cliente': None}
    
    for col in df.columns:
        col_lower = col.lower()
        if not col_map['nfe'] and any(x in col_lower for x in ['nfe', 'nota', 'fiscal', 'nf']):
            col_map['nfe'] = col
        if not col_map['pedido'] and any(x in col_lower for x in ['pedido', 'num.ped', 'num_ped']):
            col_map['pedido'] = col
        if not col_map['cliente'] and any(x in col_lower for x in ['cliente', 'nome', 'razao', 'razão']):
            col_map['cliente'] = col
    
    return col_map

# Interface Streamlit
st.title("📋 Sistema de Comparação Pedidos Enviados")

# Upload das planilhas
with st.expander("🔽 CARREGAR ARQUIVOS", expanded=True):
    col1, col2 = st.columns(2)
    with col1:
        planilha1 = st.file_uploader("Planilha Pedro", type=["xlsx", "csv"])
    with col2:
        planilha2 = st.file_uploader("Planilha DGA", type=["xlsx", "csv"])

if planilha1 and planilha2:
    # Ler os arquivos
    df1 = pd.read_excel(planilha1) if planilha1.name.endswith('.xlsx') else pd.read_csv(planilha1)
    df2 = pd.read_excel(planilha2) if planilha2.name.endswith('.xlsx') else pd.read_csv(planilha2)

    # Detectar colunas automaticamente
    cols1 = detectar_colunas(df1)
    cols2 = detectar_colunas(df2)

    # Widgets para seleção manual de colunas
    with st.expander("⚙ CONFIGURAR COLUNAS", expanded=False):
        st.warning("Confira ou ajuste as colunas detectadas:")
        
        col1, col2 = st.columns(2)
        with col1:
            st.write("**Planilha Pedro**")
            cols1['nfe'] = st.selectbox("Coluna NFe", df1.columns, index=list(df1.columns).index(cols1['nfe']) if cols1['nfe'] in df1.columns else 0)
            cols1['pedido'] = st.selectbox("Coluna Pedido", df1.columns, index=list(df1.columns).index(cols1['pedido']) if cols1['pedido'] in df1.columns else 0)
        
        with col2:
            st.write("**Planilha DGA**")
            cols2['nfe'] = st.selectbox("Coluna NFe", df2.columns, index=list(df2.columns).index(cols2['nfe']) if cols2['nfe'] in df2.columns else 0)
            cols2['pedido'] = st.selectbox("Coluna Pedido", df2.columns, index=list(df2.columns).index(cols2['pedido']) if cols2['pedido'] in df2.columns else 0)
            cols2['cliente'] = st.selectbox("Coluna Cliente", df2.columns, index=list(df2.columns).index(cols2['cliente']) if cols2['cliente'] in df2.columns else 0)

    # Processamento dos dados
    try:
        # Converter para string e limpar espaços
        for col in [cols1['nfe'], cols1['pedido']]:
            df1[col] = df1[col].astype(str).str.strip()
        
        for col in [cols2['nfe'], cols2['pedido']]:
            df2[col] = df2[col].astype(str).str.strip()

        # Criar chaves compostas
        df1['chave'] = df1[cols1['nfe']] + '|' + df1[cols1['pedido']]
        df2['chave'] = df2[cols2['nfe']] + '|' + df2[cols2['pedido']]

        # Merge para obter os clientes
        resultado = pd.merge(
            left=df1,
            right=df2[[cols2['nfe'], cols2['pedido'], cols2['cliente'], 'chave']],
            on='chave',
            how='inner'
        ).drop_duplicates('chave')

        # Selecionar colunas relevantes
        cols_resultado = [
            cols1['nfe'], 
            cols1['pedido'], 
            cols2['cliente']
        ]
        
        resultado_final = resultado[cols_resultado].copy()
        resultado_final.columns = ['NFe', 'Pedido', 'Cliente']
        
        # Adicionar controles
        resultado_final['Enviado?'] = True
        resultado_final['Observação'] = ''
        resultado_final['Data'] = pd.Timestamp.now().date()

        # Mostrar resultados
        st.success(f"✅ {len(resultado_final)} registros correspondentes encontrados")
        
        # Editor de dados
        edited_df = st.data_editor(
            resultado_final,
            column_config={
                "NFe": st.column_config.TextColumn("Nº Nota Fiscal"),
                "Pedido": st.column_config.TextColumn("Nº Pedido"),
                "Cliente": st.column_config.TextColumn("Nome do Cliente"),
                "Enviado?": st.column_config.CheckboxColumn(
                    "Enviado?",
                    help="Desmarque se não foi enviado",
                    default=True
                ),
                "Data": st.column_config.DateColumn(
                    "Data da Verificação",
                    format="DD/MM/YYYY",
                    disabled=True
                )
            },
            hide_index=True,
            use_container_width=True
        )

        # Exportar
        st.divider()
        if st.button("💾 Gerar Relatório Final"):
            output = BytesIO()
            with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                edited_df.to_excel(writer, index=False, sheet_name='Controle')
                
                # Adicionar resumo
                resumo = pd.DataFrame({
                    'Total os pedidos faturados': [len(edited_df)],
                    'Enviados': [edited_df['Enviado?'].sum()],
                    'Pendentes': [len(edited_df) - edited_df['Enviado?'].sum()]
                })
                resumo.to_excel(writer, index=False, sheet_name='Resumo')
            
            st.download_button(
                label="⬇️ Baixar Relatório (.xlsx)",
                data=output.getvalue(),
                file_name=f"controle_envio_{pd.Timestamp.now().strftime('%Y%m%d')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

    except Exception as e:
        st.error(f"❌ Erro: {str(e)}")
        st.stop()

else:
    st.warning("⚠️ Carregue as planilhas para Comparar os pedidos.")
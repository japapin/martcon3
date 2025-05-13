import streamlit as st
import pandas as pd
import numpy as np
from io import BytesIO

# Configuração da página
st.set_page_config(page_title="Análise de Estoque", layout="wide")
st.title("📈 Análise de Cobertura de Estoque")

# Upload do arquivo
uploaded_file = st.file_uploader("Carregue seu arquivo Excel (análise.xlsx)", type=["xlsx"])

if uploaded_file:
    try:
        # Carregar dados
        df = pd.read_excel(uploaded_file)

        # Validação das colunas obrigatórias
        required_cols = ["Filial", "Cobertura Atual", "Vlr Estoque Tmk", "Mercadoria", "Saldo Pedido"]
        if not all(col in df.columns for col in required_cols):
            missing_cols = [col for col in required_cols if col not in df.columns]
            st.error(f"⚠️ Arquivo inválido! Faltam as colunas: {', '.join(missing_cols)}")
            st.stop()

        # Renomear colunas
        df = df.rename(columns={
            "Vlr Estoque Tmk": "valor_estoque",
            "Cobertura Atual": "cobertura_dias",
            "Filial": "filial",
            "Saldo Pedido": "saldo_pedido"
        })

        # Filtrar dados válidos (saldo e cobertura positivos)
        df = df[(df['cobertura_dias'] > 0) & (df['saldo_pedido'] > 0)].copy()

        # 📌 Cobertura Média por Filial
        st.subheader("📌 Cobertura Média por Filial")

        # Função para cálculo seguro da média ponderada
        def calcular_media_ponderada(grupo):
            try:
                return np.average(grupo["cobertura_dias"], weights=grupo["valor_estoque"])
            except:
                return 0

        cobertura = (
            df.groupby("filial")
            .apply(lambda grupo: pd.Series({
                "Cobertura Média Ponderada (dias)": calcular_media_ponderada(grupo),
                "Cobertura Média Simples (dias)": grupo["cobertura_dias"].mean(),
                "Saldo Pedido Total": grupo["saldo_pedido"].sum()
            }))
            .round(2)
            .reset_index()
            .rename(columns={"filial": "Filial"})
        )

        # Formatação da tabela
        styled_cobertura = cobertura.style \
            .format({
                "Cobertura Média Ponderada (dias)": "{:.2f}",
                "Cobertura Média Simples (dias)": "{:.2f}",
                "Saldo Pedido Total": "R$ {:,.2f}"
            }, na_rep="-") \
            .set_properties(**{'text-align': 'center'}) \
            .set_table_styles([{'selector': 'th', 'props': [('text-align', 'center')]}])

        st.dataframe(styled_cobertura, use_container_width=True)

        # 📊 Distribuição por Faixa de Cobertura (usando Saldo Pedido)
        st.subheader("📊 Distribuição por Faixa de Cobertura (Saldo de Pedido)")

        # Criar faixas de cobertura
        df['faixa'] = pd.cut(
            df['cobertura_dias'],
            bins=[0, 15, 30, 45, 60, np.inf],
            labels=["0-15 dias", "16-30 dias", "31-45 dias", "46-60 dias", "Mais de 60 dias"],
            right=False
        )

        # Agrupar por filial e faixa, somando o saldo de pedido
        resumo_valores = df.groupby(['filial', 'faixa'])['saldo_pedido'].sum().unstack().fillna(0)

        # Adicionar coluna TOTAL por filial
        resumo_valores['TOTAL'] = resumo_valores.sum(axis=1)

        # Tabela de Valores Absolutos
        st.markdown("**Valores Absolutos (R$)**")
        styled_resumo_valores = resumo_valores.style \
            .format("R$ {:,.2f}", na_rep="-") \
            .set_properties(**{'text-align': 'center'}) \
            .set_table_styles([{'selector': 'th', 'props': [('text-align', 'center')]}])
        
        st.dataframe(styled_resumo_valores, use_container_width=True)

        # Tabela de Percentuais (separada)
        st.markdown("**Percentuais por Faixa (%)**")
        
        # Calcular percentuais
        resumo_percentuais = resumo_valores.copy()
        for col in resumo_percentuais.columns:
            if col != 'TOTAL':
                resumo_percentuais[col] = (resumo_percentuais[col] / resumo_percentuais['TOTAL'] * 100).round(2)
        
        # Remover coluna TOTAL dos percentuais
        resumo_percentuais = resumo_percentuais.drop(columns=['TOTAL'])
        
        # Formatar com símbolo %
        styled_resumo_percentuais = resumo_percentuais.style \
            .format("{:.2f}%", na_rep="-") \
            .set_properties(**{'text-align': 'center'}) \
            .set_table_styles([{'selector': 'th', 'props': [('text-align', 'center')]}])
        
        st.dataframe(styled_resumo_percentuais, use_container_width=True)

        # 📥 Exportar para Excel
        output = BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            cobertura.to_excel(writer, sheet_name='Cobertura Média', index=False)
            resumo_valores.to_excel(writer, sheet_name='Valores Absolutos')
            resumo_percentuais.to_excel(writer, sheet_name='Percentuais')
        
        st.download_button(
            label="📥 Baixar Relatório Completo (Excel)",
            data=output.getvalue(),
            file_name="relatorio_estoque.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    except Exception as e:
        st.error(f"Ocorreu um erro ao processar o arquivo: {str(e)}")
        st.stop()

else:
    st.warning("Por favor, carregue um arquivo Excel para análise.")
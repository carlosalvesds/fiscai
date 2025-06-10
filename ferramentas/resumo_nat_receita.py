# ferramentas/resumo_nat_receita.py

import streamlit as st
import pandas as pd
from io import BytesIO, StringIO

def app():
    st.title("📊 Resumo por Código de Natureza da Receita")
    st.markdown("""
Essa ferramenta permite processar arquivos .txt ou .html contendo dados fiscais detalhados por item, com o objetivo de gerar um resumo agrupado por Código de Natureza da Receita.
Essa funcionalidade é útil para análises de conferência tributária, auditoria interna e cruzamentos fiscais relacionados à receita declarada por tipo de operação.
""")

    uploaded_file = st.file_uploader("Envie seu arquivo (.txt)", type=None)

    if uploaded_file:
        filename = uploaded_file.name.lower()

        try:
            if filename.endswith(".txt"):
                df = pd.read_csv(uploaded_file, sep=",", header=None, quotechar='"', encoding="utf-8")
                df.columns = [
                    "Documento", "Descrição", "Cod_Item", "Nat_Receita", "Cód. STB",
                    "NCM", "CST_PIS", "CST_COFINS", "CFOP", "Qtde",
                    "Valor_Produto", "Desconto", "Valor_Total"
                ]
            elif filename.endswith(".htm") or filename.endswith(".html"):
                conteudo = uploaded_file.read().decode("utf-8", errors="ignore")
                tabelas = pd.read_html(StringIO(conteudo))
                if not tabelas:
                    st.warning("Nenhuma tabela encontrada.")
                    st.stop()
                df = tabelas[0]
                df.columns = [str(c).strip() for c in df.columns]
            else:
                st.error("Formato de arquivo não suportado.")
                st.stop()

            df["Valor_Total"] = (
                df["Valor_Total"]
                .astype(str)
                .str.replace(".", "", regex=False)
                .str.replace(",", ".", regex=False)
                .str.replace("R$", "", regex=False)
                .str.replace(" ", "", regex=False)
            )
            df["Valor_Total"] = pd.to_numeric(df["Valor_Total"], errors="coerce")

            linhas_iniciais = len(df)
            df = df[df["Valor_Total"].notnull()]
            linhas_validas = len(df)

            total_geral = df["Valor_Total"].sum().round(2)
            st.info(f"💰 Valor total geral (sem filtros): R$ {total_geral:,.2f}")

            st.subheader("🎯 Filtro CST_PIS")
            cst_pis_unicos = sorted(df["CST_PIS"].astype(str).dropna().unique())
            cst_pis_selecionado = st.multiselect("Selecione os CST_PIS desejados (deixe vazio para total geral)", cst_pis_unicos)

            df_filtrado = df.copy()
            if cst_pis_selecionado:
                df_filtrado = df[df["CST_PIS"].astype(str).isin(cst_pis_selecionado)]

            total_filtrado = df_filtrado["Valor_Total"].sum().round(2)
            st.success(f"🔎 Total após filtro: R$ {total_filtrado:,.2f}")

            resumo = (
                df_filtrado.groupby("Nat_Receita")["Valor_Total"]
                .sum()
                .reset_index()
                .rename(columns={"Nat_Receita": "Cód. Nat Receita", "Valor_Total": "Total Valor Contábil"})
            )
            resumo["Total Valor Contábil"] = resumo["Total Valor Contábil"].round(2)

            st.subheader("📋 Resumo agrupado por Código de Natureza")
            resumo_exibicao = resumo.copy()
            resumo_exibicao["Total Valor Contábil"] = resumo_exibicao["Total Valor Contábil"].apply(
                lambda x: f"R$ {x:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
            )
            st.dataframe(resumo_exibicao)

            def gerar_excel(df):
                output = BytesIO()
                with pd.ExcelWriter(output, engine="openpyxl") as writer:
                    df.to_excel(writer, index=False, sheet_name="Resumo")
                return output.getvalue()

            excel_bytes = gerar_excel(resumo)
            st.download_button("📥 Baixar resumo como Excel", data=excel_bytes, file_name="resumo_nat_receita.xlsx")

            if linhas_iniciais != linhas_validas:
                st.warning(f"{linhas_iniciais - linhas_validas} linhas com valores inválidos foram ignoradas.")

        except Exception as e:
            st.error(f"❌ Erro ao processar o arquivo: {e}")

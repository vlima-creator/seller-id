import re
import pandas as pd
import requests
import streamlit as st
from io import BytesIO

st.set_page_config(page_title="Extrator Seller ID Mercado Livre")

st.title("Extrator Seller ID Mercado Livre")

arquivo = st.file_uploader("Envie uma planilha Excel (.xlsx) com uma coluna chamada 'url'", type=["xlsx"])

def extrair_dados(url):
    try:
        r = requests.get(
            url,
            headers={"User-Agent":"Mozilla/5.0"},
            timeout=30
        )

        html = r.text

        seller_id = ""
        loja = ""

        m = re.search(r'"seller_id":\s*(\d+)', html)
        if m:
            seller_id = m.group(1)

        n = re.search(r'"nickname":"([^"]+)"', html)
        if n:
            loja = n.group(1)

        return seller_id, loja

    except Exception:
        return "", ""

if arquivo:
    df = pd.read_excel(arquivo)

    if "url" not in df.columns:
        st.error("A planilha precisa ter uma coluna chamada url")
    else:
        resultado = []

        progresso = st.progress(0)

        total = len(df)

        for i, row in df.iterrows():
            seller_id, loja = extrair_dados(str(row["url"]))

            resultado.append({
                "url": row["url"],
                "seller_id": seller_id,
                "loja": loja
            })

            progresso.progress((i + 1) / total)

        resultado_df = pd.DataFrame(resultado)

        st.success("Processamento concluído")
        st.dataframe(resultado_df)

        output = BytesIO()
        resultado_df.to_excel(output, index=False)

        st.download_button(
            "Baixar resultado",
            data=output.getvalue(),
            file_name="resultado_seller_id.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
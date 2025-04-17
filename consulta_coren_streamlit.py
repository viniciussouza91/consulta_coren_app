
import streamlit as st
import pandas as pd
import time
import tempfile
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from webdriver_manager.chrome import ChromeDriverManager

st.set_page_config(page_title="Consulta Coren-SP", layout="centered")

st.title("🩺 Consulta de Profissionais - Coren-SP")
st.markdown("Automação para busca em lote via planilha.")

uploaded_file = st.file_uploader("📤 Envie a planilha .xlsx com a coluna `termo_pesquisa`", type=["xlsx"])

if uploaded_file:
    df = pd.read_excel(uploaded_file)
    
    if "termo_pesquisa" not in df.columns:
        st.error("❌ A planilha deve conter a coluna 'termo_pesquisa'.")
    else:
        termos = df['termo_pesquisa'].dropna().tolist()
        st.success(f"✅ {len(termos)} termos encontrados para consulta.")

        if st.button("🚀 Executar Consulta"):
            st.info("⏳ Iniciando automação... isso pode levar alguns minutos.")
            progress_bar = st.progress(0)

            # Configurar Selenium
            options = webdriver.ChromeOptions()
            options.add_argument('--headless')
            options.add_argument('--disable-gpu')
            options.add_argument('--no-sandbox')
            service = Service(ChromeDriverManager().install())
            driver = webdriver.Chrome(service=service, options=options)

            url = "https://servicos.coren-sp.gov.br/consulta-de-profissionais/"
            todos_resultados = []

            for i, termo in enumerate(termos):
                try:
                    driver.get(url)
                    time.sleep(2)

                    campo_nome = driver.find_element(By.ID, "nome")
                    botao_buscar = driver.find_element(By.ID, "btnConsultar")

                    campo_nome.clear()
                    campo_nome.send_keys(termo)
                    botao_buscar.click()
                    time.sleep(3)

                    tabela = driver.find_element(By.ID, "tabelaResultados")
                    linhas = tabela.find_elements(By.TAG_NAME, "tr")[1:]

                    for linha in linhas:
                        colunas = linha.find_elements(By.TAG_NAME, "td")
                        if len(colunas) >= 5:
                            todos_resultados.append({
                                'Nome': colunas[0].text,
                                'Inscrição': colunas[1].text,
                                'Categoria': colunas[2].text,
                                'Situação': colunas[3].text,
                                'Data de Inscrição': colunas[4].text,
                                'Termo Pesquisado': termo
                            })
                except:
                    pass

                progress_bar.progress((i + 1) / len(termos))

            driver.quit()

            if todos_resultados:
                df_resultados = pd.DataFrame(todos_resultados)

                st.success("✅ Consulta finalizada com sucesso!")
                st.dataframe(df_resultados)

                # Exportar resultados
                with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as tmp:
                    df_resultados.to_excel(tmp.name, index=False)
                    st.download_button("📥 Baixar resultados em Excel", data=open(tmp.name, "rb"), file_name="resultado_coren.xlsx")
            else:
                st.warning("⚠️ Nenhum resultado encontrado para os termos enviados.")

import streamlit as st
import json
from Gear_com_revisao_V28 import main_function  # ajuste para o nome real da função principal do Gear

st.set_page_config(page_title="Gear Revisão Espaciada", page_icon="📚")

st.title("📅 Gerador de Cronograma – Gear com Revisão Espaciada")

# Lê ou cria config padrão
try:
    with open("scheduler_config.json", "r", encoding="utf-8") as f:
        config = json.load(f)
except FileNotFoundError:
    config = {}

minutos_por_dia = st.number_input("Minutos de estudo por dia", min_value=30, max_value=600, value=config.get("minutos_por_dia", 120))
dias_por_semana = st.number_input("Dias de estudo por semana", min_value=1, max_value=7, value=config.get("dias_por_semana", 5))
data_inicio = st.date_input("Data de início", value=None)
data_prova = st.date_input("Data da prova", value=None)
tipo_prova = st.selectbox("Tipo de prova", ["TEA", "TSA", "ME1", "ME2", "ME3"], index=1)

temas_path = st.file_uploader("Arquivo de temas (.xlsx)", type="xlsx")
aulas_path = st.file_uploader("Arquivo de aulas (.xlsx)", type="xlsx")
capa_path = st.file_uploader("Capa (PNG)", type="png")
orient_path = st.file_uploader("Orientações (PDF)", type="pdf")
template_path = st.file_uploader("Template .dotx (opcional)", type="dotx")

if st.button("Gerar Cronograma"):
    st.write("Gerando cronograma...")

    config = {
        "minutos_por_dia": minutos_por_dia,
        "dias_por_semana": dias_por_semana,
        "data_inicio": data_inicio.strftime("%d/%m/%Y"),
        "data_prova": data_prova.strftime("%d/%m/%Y"),
        "tipo_prova": tipo_prova,
        "review_offsets": [30],  # fixa ou adicione controles para isso
    }

    # chama a função principal do Gear
    main_function(config)
    st.success("Cronograma gerado com sucesso!")

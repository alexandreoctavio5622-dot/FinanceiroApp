import streamlit as st
import pandas as pd
from datetime import datetime
import re
import os
from io import BytesIO

# CONFIGURAÇÃO
EXCEL_PATH = "Formulario.xlsx"
SHEET_NAME = "Janeiro-26"

st.title("Controle Financeiro Pessoal")
st.write("Adicione lançamentos e baixe a planilha atualizada no final.")

def parse_valor(valor_str):
    if not valor_str: return 0.0
    valor_str = re.sub(r"[R$\s]", "", str(valor_str)).replace(",", ".")
    try: return float(valor_str)
    except: return 0.0

# Inicializar o DataFrame na sessão para manter os dados enquanto o app estiver aberto
if 'df_temp' not in st.session_state:
    if os.path.exists(EXCEL_PATH):
        st.session_state.df_temp = pd.read_excel(EXCEL_PATH, sheet_name=SHEET_NAME)
    else:
        st.error("Arquivo base não encontrado!")

with st.form(key='form_lancamento'):
    data = st.date_input("Data", value=datetime.today())
    descricao = st.text_input("Descrição")
    nfp = st.text_input("NFP (opcional)")
    codigo = st.text_input("Código")
    forma_pagto = st.selectbox("Forma de Pagto.", ["débito", "crédito", "dinheiro", "VA", "cartão CEA pay"])
    debito = st.text_input("Débito Conta Corrente", value="")
    credito = st.text_input("Crédito Conta Corrente", value="")
    submit_button = st.form_submit_button(label='Adicionar Lançamento')

if submit_button:
    nova_linha = {
        st.session_state.df_temp.columns[0]: data.strftime("%d/%m/%Y"),
        st.session_state.df_temp.columns[1]: descricao,
        st.session_state.df_temp.columns[2]: nfp,
        st.session_state.df_temp.columns[3]: codigo,
        st.session_state.df_temp.columns[4]: forma_pagto,
        st.session_state.df_temp.columns[5]: parse_valor(debito),
        st.session_state.df_temp.columns[6]: parse_valor(credito)
    }
    st.session_state.df_temp = pd.concat([st.session_state.df_temp, pd.DataFrame([nova_linha])], ignore_index=True)
    st.success("Lançamento adicionado à lista temporária!")

# Mostrar os últimos lançamentos realizados nesta sessão
st.write("### Lançamentos da Sessão")
st.dataframe(st.session_state.df_temp.tail(5))

# Botão de Download
st.write("---")
buffer = BytesIO()
with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
    st.session_state.df_temp.to_excel(writer, sheet_name=SHEET_NAME, index=False)

st.download_button(
    label="📥 Baixar Planilha Atualizada",
    data=buffer.getvalue(),
    file_name=f"Financeiro_Atualizado_{datetime.now().strftime('%d-%m-%Y')}.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
)

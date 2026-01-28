import streamlit as st
import pandas as pd
from datetime import datetime
import re
import os

# CONFIGURAÇÃO
EXCEL_PATH = "Formulario.xlsx"
SHEET_NAME = "Janeiro-26"

st.title("Controle Financeiro Pessoal")
st.write("Adicione lançamentos de Débito ou Crédito na planilha.")

def parse_valor(valor_str):
    if not valor_str: return 0.0
    valor_str = re.sub(r"[R$\s]", "", str(valor_str)).replace(",", ".")
    try: return float(valor_str)
    except: return 0.0

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
    try:
        if not os.path.exists(EXCEL_PATH):
            st.error(f"Arquivo '{EXCEL_PATH}' não encontrado!")
        else:
            # Lendo a planilha com Pandas (mais tolerante a erros de XML)
            df = pd.read_excel(EXCEL_PATH, sheet_name=SHEET_NAME)
            
            # Criando a nova linha
            nova_linha = {
                df.columns[0]: data.strftime("%d/%m/%Y"),
                df.columns[1]: descricao,
                df.columns[2]: nfp,
                df.columns[3]: codigo,
                df.columns[4]: forma_pagto,
                df.columns[5]: parse_valor(debito),
                df.columns[6]: parse_valor(credito)
            }
            
            # Adicionando ao DataFrame
            df = pd.concat([df, pd.DataFrame([nova_linha])], ignore_index=True)
            
            # Salvando de volta no Excel
            with pd.ExcelWriter(EXCEL_PATH, engine='openpyxl', mode='w') as writer:
                df.to_excel(writer, sheet_name=SHEET_NAME, index=False)
            
            st.success("Lançamento adicionado com sucesso!")
            st.info("Nota: Os dados são salvos temporariamente no servidor do Streamlit.")
            
    except Exception as e:
        st.error(f"Erro ao processar Excel: {e}")

# app_financeiro.py
import streamlit as st
import xlwings as xw
from datetime import datetime
import re

# ===========================
# CONFIGURAÇÃO
# ===========================
EXCEL_PATH = r"C:\Users\alexandre\Desktop\Pasta\Pessoal\Formulario.xlsx"
SHEET_NAME = "Janeiro-26"

# ===========================
# TÍTULO DO APP
# ===========================
st.title("Controle Financeiro Pessoal")
st.write("Adicione lançamentos de Débito ou Crédito na planilha.")

# ===========================
# FUNÇÃO PARA LIMPAR VALORES MONETÁRIOS
# ===========================
def parse_valor(valor_str):
    """
    Converte string do tipo "R$ 2.000,50" ou "2000,50" para float 2000.50
    """
    if not valor_str:
        return None
    # Remove R$, espaços e substitui vírgula por ponto
    valor_str = re.sub(r"[R$\s]", "", valor_str)
    valor_str = valor_str.replace(",", ".")
    try:
        return float(valor_str)
    except:
        return None

# ===========================
# FORMULÁRIO DE LANÇAMENTO
# ===========================
with st.form(key='form_lancamento'):
    data = st.date_input("Data", value=datetime.today())
    descricao = st.text_input("Descrição")
    nfp = st.text_input("NFP (opcional)")
    codigo = st.text_input("Código")
    forma_pagto = st.selectbox("Forma de Pagto.", ["débito", "crédito", "dinheiro", "VA", "cartão CEA pay"])
    
    # Cofres / Colunas de valores
    debito = st.text_input("Débito Conta Corrente", value="")
    credito = st.text_input("Crédito Conta Corrente", value="")


    submit_button = st.form_submit_button(label='Adicionar Lançamento')

# ===========================
# PROCESSAR ENVIO
# ===========================
if submit_button:
    try:
        # Abrir planilha
        wb = xw.Book(EXCEL_PATH)
        if SHEET_NAME not in [s.name for s in wb.sheets]:
            st.error(f"Aba '{SHEET_NAME}' não encontrada!")
        else:
            sheet = wb.sheets[SHEET_NAME]

            # Encontrar última linha preenchida na coluna A (Data)
            col_a = sheet.range("A:A").value
            col_a = [c for c in col_a if c is not None]
            last_row = len(col_a)
            next_row = last_row + 1

            # Preencher dados
            sheet.range(f"A{next_row}").value = data  # formato Excel reconhece como data
            sheet.range(f"B{next_row}").value = descricao
            sheet.range(f"C{next_row}").value = nfp
            sheet.range(f"D{next_row}").value = codigo
            sheet.range(f"E{next_row}").value = forma_pagto
            
            # Valores convertidos para float
            sheet.range(f"F{next_row}").value = parse_valor(debito)
            sheet.range(f"G{next_row}").value = parse_valor(credito)


            # Salvar e fechar
            wb.save()
            wb.close()
            st.success("Lançamento adicionado com sucesso!")

    except Exception as e:
        st.error(f"Erro ao gravar no Excel: {e}")

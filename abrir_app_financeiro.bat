@echo off
REM Definindo o caminho do Python
set PYTHON=C:\Users\alexandre\AppData\Local\Python\bin\python.exe

REM Caminho do seu arquivo Streamlit
set APP=C:\Users\alexandre\Desktop\Pasta\Pessoal\app_financeiro.py

REM Executando o Streamlit
%PYTHON% -m streamlit run "%APP%"

pause

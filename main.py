from fastapi import FastAPI
from pydantic import BaseModel
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import FileResponse
from datetime import datetime
import sqlite3

app = FastAPI()

# Libera acesso externo (necessário para navegador)
app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

DATABASE = "financeiro.db"

# Criar tabela automaticamente
def criar_tabela():
    conn = sqlite3.connect(DATABASE)
    cursor = conn.cursor()
    cursor.execute("""
        CREATE TABLE IF NOT EXISTS lancamentos (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            data TEXT,
            descricao TEXT,
            codigo TEXT,
            tipo TEXT,
            valor REAL
        )
    """)
    conn.commit()
    conn.close()

criar_tabela()

class Lancamento(BaseModel):
    data: str
    descricao: str
    codigo: str
    valor: float
    tipo: str  # debito ou credito


# Página principal (abre o index.html)
@app.get("/")
def home():
    return FileResponse("index.html")


# Adicionar lançamento
@app.post("/lancamento/")
def adicionar_lancamento(l: Lancamento):

    try:
        datetime.strptime(l.data, "%Y-%m-%d")
    except:
        return {"erro": "Formato de data inválido. Use YYYY-MM-DD"}

    conn = sqlite3.connect(DATABASE)
    cursor = conn.cursor()

    cursor.execute("""
        INSERT INTO lancamentos (data, descricao, codigo, tipo, valor)
        VALUES (?, ?, ?, ?, ?)
    """, (l.data, l.descricao, l.codigo, l.tipo.lower(), l.valor))

    conn.commit()
    conn.close()

    return {"mensagem": "Lançamento salvo com sucesso"}


# Listar lançamentos
@app.get("/lancamentos/")
def listar_lancamentos():

    conn = sqlite3.connect(DATABASE)
    cursor = conn.cursor()

    cursor.execute("""
        SELECT id, data, descricao, codigo, tipo, valor
        FROM lancamentos
        ORDER BY data DESC
    """)

    dados = cursor.fetchall()

    conn.close()

    return {"dados": dados}

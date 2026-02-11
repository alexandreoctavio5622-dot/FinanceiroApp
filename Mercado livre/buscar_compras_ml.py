import requests
import pandas as pd
from datetime import datetime
from ml_token_manager import obter_token

print("Obtendo token...")
ACCESS_TOKEN = obter_token()

headers = {
    "Authorization": f"Bearer {ACCESS_TOKEN}"
}

print("Obtendo usuário...")
user_resp = requests.get(
    "https://api.mercadolibre.com/users/me",
    headers=headers
)

user_id = user_resp.json()["id"]

dados = []
ano_inicial = 2010
ano_final = datetime.now().year

for ano in range(ano_inicial, ano_final + 1):
    print(f"\n🔎 Buscando compras do ano {ano}...")

    date_from = f"{ano}-01-01T00:00:00.000-00:00"
    date_to   = f"{ano}-12-31T23:59:59.000-00:00"

    offset = 0
    limit = 50

    while True:
        url = (
            "https://api.mercadolibre.com/orders/search?"
            f"buyer={user_id}"
            f"&order.date_created.from={date_from}"
            f"&order.date_created.to={date_to}"
            f"&offset={offset}"
            f"&limit={limit}"
        )

        resp = requests.get(url, headers=headers)
        result = resp.json()

        orders = result.get("results", [])
        if not orders:
            break

        for order in orders:
            data = order["date_created"]
            valor = order["total_amount"]
            status = order["status"]

            for item in order["order_items"]:
                produto = item["item"]["title"]

                dados.append({
                    "Produto": produto,
                    "Data": data,
                    "Valor Pedido": valor,
                    "Status": status
                })

        offset += limit

print("\nTotal de registros encontrados:", len(dados))

df = pd.DataFrame(dados)
df.to_excel("compras_mercadolivre_completo.xlsx", index=False)

print("📊 Relatório gerado com sucesso!")

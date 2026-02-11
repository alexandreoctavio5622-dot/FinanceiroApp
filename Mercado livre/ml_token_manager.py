import requests
import json
import time

CLIENT_ID = "3441038985074666"
CLIENT_SECRET = "NDNhav6uBkFJLoRh4N3BquO6zR6G6Zup"
REFRESH_TOKEN = "TG-6980a7eedc5a1300018f166a-38135237"

TOKEN_FILE = "token.json"


def renovar_token():
    url = "https://api.mercadolibre.com/oauth/token"

    payload = {
        "grant_type": "refresh_token",
        "client_id": CLIENT_ID,
        "client_secret": CLIENT_SECRET,
        "refresh_token": REFRESH_TOKEN
    }

    response = requests.post(url, data=payload)
    data = response.json()

    if "access_token" in data:
        data["expires_at"] = time.time() + data["expires_in"]

        with open(TOKEN_FILE, "w") as f:
            json.dump(data, f)

        print("🔄 Token renovado com sucesso.")
        return data["access_token"]
    else:
        raise Exception(f"Erro ao renovar token: {data}")


def obter_token():
    try:
        with open(TOKEN_FILE, "r") as f:
            data = json.load(f)

        if time.time() > data["expires_at"]:
            return renovar_token()

        return data["access_token"]

    except:
        return renovar_token()

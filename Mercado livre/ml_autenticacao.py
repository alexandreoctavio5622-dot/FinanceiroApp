import requests

CLIENT_ID = "3441038985074666"
CLIENT_SECRET = "NDNhav6uBkFJLoRh4N3BquO6zR6G6Zup"
CODE = "TG-6980a7926bef2b0001e6d102-38135237"

url = "https://api.mercadolibre.com/oauth/token"

payload = {
    "grant_type": "authorization_code",
    "client_id": CLIENT_ID,
    "client_secret": CLIENT_SECRET,
    "code": CODE,
    "redirect_uri": "https://oauth.pstmn.io/v1/callback"
}

response = requests.post(url, data=payload)

print("Status:", response.status_code)
print("Resposta completa:", response.text)

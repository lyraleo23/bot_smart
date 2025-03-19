import requests
import json

def atualizar_req_miliapp(TOKEN_MILIAPP, req_obj):
    url = f'https://api.fmiligrama.com/vendas/smart?token={TOKEN_MILIAPP}'

    payload = json.dumps(req_obj)
    headers = {
        'Accept': 'application/json',
        'Content-Type': 'application/json'
    }

    response = requests.request("POST", url, headers=headers, data=payload)
    print(response.text)

    return

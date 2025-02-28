import requests
import json

def consulta_api(req, num_ped):
    #Api tiny
    url = "https://api.fmiligrama.com/vendas/smart"
    num_ped = str(num_ped)

    #Dados para api
    payload = json.dumps({
        "numero_tiny": num_ped,
        "req_smart": req,
        "cnpj": '07413904000198'
    })
    #Cabeçalho para envio da api
    headers = {
        'Accept': 'application/json',
        'Content-Type': 'application/json'
    }
    #Comando de envio para api
    response = requests.request("POST", url, headers=headers, data=payload)
    
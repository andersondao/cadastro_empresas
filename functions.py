import json

from requests import request


def consulta_cnpj(cnpj):
    url = f"https://www.receitaws.com.br/v1/cnpj/{cnpj}"
    querystring = {
        "token": "XXXXXXXX-XXXX-XXXX-XXXX-XXXXXXXXXXXX",
        "cnpj": "06990590000123",
        "plugin": "RF",
    }
    response = request("GET", url, params=querystring)

    resp = json.loads(response.text)
    # print(resp)
    return (
        resp["nome"],
        resp["logradouro"],
        resp["numero"],
        resp["complemento"],
        resp["bairro"],
        resp["municipio"],
        resp["uf"],
        resp["cep"],
        resp["telefone"],
        resp["email"],
    )

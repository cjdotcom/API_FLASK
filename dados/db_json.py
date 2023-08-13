from datetime import datetime
import json

def get(cod):
    with open('', 'r') as f:
        pass
    pass

def post(cod, nome_produto):
    dt_criacao = datetime.now() #data atual
    form = f"""
    "codigo": "{cod}",
    "name": "{nome_produto}",
    "dataCriação": "{dt_criacao}"
    """

    cad_prod = json.loads(form)

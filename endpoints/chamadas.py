from flask_restx import Resource, Api
from flask import request
from server.instance import server
from openpyxl import load_workbook
from datetime import datetime
import pandas as pd
import json

app, api = server.app, server.api

@api.route('/Status')
class ApiStatus(Resource):
    def get(self,):
        dia = datetime.now() #data atual
        lista = """{
            "Status": "Okay",
            "Date": ""
            }"""
        db = json.loads(lista)
        db['Date'] = str(dia)
        return db

@api.route('/Products')
class ProdutoByOrder(Resource):
    def get(self,):
        caminhoExcel = "db_excel.xlsx"
        wb = load_workbook(caminhoExcel)
        sheet = wb['DADOS']
        lista = []
        for i in range(2, sheet.max_row):
            linha = [x.value for x in sheet[i]]
            if linha[0] != None:
                lista.append(linha)
            else:
                break

        t = {}
        for i in lista:
            t[i[0]] = [i[1], i[2]]

        form = """
            {
            "codigo":"",
            "name":"",
            "dtRegistro":""
            }
        """
        tdb = json.loads(form)

        retorno = []
        for k, v in t.items():
            info = tdb.copy()

            info['codigo'] = k
            info['name'] = v[0]
            info['dtRegistro'] = v[1]

            retorno.append({"Products":info})

        df = pd.DataFrame(retorno)
        return json.loads(df.to_json())

    def post(self,):
        pass

    def put(self,):
        pass
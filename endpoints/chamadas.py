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
class Produtos(Resource):
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
    
@api.route('/Product')
class ProdutoByCodigo(Resource):
    def get(self,):
        codigo = request.args.get('codigo')
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

        erro_cod = """
            {
                "erro": "A chamada /Product requer o parametro <codigo>", 
                "param": "invalid"
            }
        """
        if codigo != None:
            df = pd.DataFrame(retorno)
            df_real = None
            for db_cod in df['Products'].values:
                if db_cod['codigo'] == int(codigo):
                    df_real = json.dumps({"Product":db_cod})
            
            if df_real != None:
                return json.loads(df_real)
            else:
                erro_prod = """
                    {
                        "erro": "Código não cadastrado!"
                    }
                """
                return json.loads(erro_prod)
        else:
            return json.loads(erro_cod)
        
@api.route('/CadProduct')
class CadastroProduto(Resource):
    def post(self,):
        codigo = request.args['codigo']
        nomeProduto = request.args['nomeProduto']

        caminhoExcel = "db_excel.xlsx"
        wb = load_workbook(caminhoExcel)
        sheet = wb['DADOS']
        lista = []
        c = 0
        for i in range(2, sheet.max_row):
            linha = [x.value for x in sheet[i]]
            if linha[0] != None:
                lista.append(linha)
                c += 1
            else:
                break

        t = {}
        for i in lista:
            t[i[0]] = [i[1], i[2]]

        gatilho = None
        for k, v in t.items():
            if int(codigo) == k:
                gatilho = True
                break
            else:
                gatilho = False

        if gatilho == False:
            linha = c+1
            sheet.cell(row=linha, column=1, value=int(codigo))
            sheet.cell(row=linha, column=2, value=str(nomeProduto))
            sheet.cell(row=linha, column=3, value=datetime.today().strftime("%d/%m/%Y %H:%M:%S"))

            wb.save(caminhoExcel)

            succes = """
                {
                    "info":"Produto cadastrado!"
                }
            """
            return json.loads(succes)
        else:
            erro = """
                {
                    "error":"Produto já existe."
                }
            """
            return json.loads(erro)
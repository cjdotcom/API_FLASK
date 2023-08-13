from flask import Flask
from flask_restx import Api
# from waitress import serve

class Server():
    def __init__(self,):
        self.app = Flask(__name__)
        self.api = Api(self.app,
                       title='API_FLASK_PRODUTOS'
                       )
    
    def run(self,):
        # serve(self.api, port=4000) #SERVIDOR WSGI PARA PRODUÇÃO
        self.app.run(host="0.0.0.0", port=4000, debug=True) #SERVIDOR LOCAL PARA MANUTENÇÃO

server = Server()

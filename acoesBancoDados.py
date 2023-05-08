import pyodbc

class dadosConexao:
    def __init__(self, ipServidor,bancoDados,usuario,senha):
        self.ipServidor = ipServidor
        self.bancoDados = bancoDados
        self.usuario = usuario
        self.senha = senha



class acaoBancoDados(dadosConexao):
    def __init__(self, ipServidor, bancoDados, usuario, senha,query):
        dadosConexao.__init__(self, ipServidor, bancoDados, usuario, senha)
        self.query = query

    def executa(self):
        cnxn = pyodbc.connect(
            'DRIVER={ODBC Driver 18 for SQL Server};SERVER=' + self.ipServidor + ';DATABASE=' + self.bancoDados + ';UID=' + self.usuario + ';PWD=' + self.senha + ';TrustServerCertificate=YES')
        cursor = cnxn.cursor()
        try:
            if ('SELECT' in self.query and 'INSERT' not in self.query) or 'SET NOCOUNT ON' in self.query:
                cursor.execute(self.query)
                resultado = cursor.fetchall()
                return resultado
            else:
                cursor.execute(self.query)
                cursor.commit()
                return 'ACAO REALIZADA COM SUCESSO'
        except:
            return 'ERRO AO REALIZAR AÇÃO'


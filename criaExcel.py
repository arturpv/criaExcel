import os
from openpyxl import Workbook
import datetime
import acoesBancoDados

def cria_excel():
    consulta = '''  
        SELECT
            [COLUNA]
        FROM TB_ANALITICO_MONITORIA WITH(NOLOCK)   
        WHERE [COLUNA] = '2023-03-06T15:38'
    '''

    # Conecta ao banco de dados e executa a consulta
    with acoesBancoDados.acaoBancoDados('server', 'banco', 'login','senha', consulta) as conexao_banco:
        executa_dados_para_excel = conexao_banco.executa()

    # Cria o arquivo em Excel
    planilha_excel = Workbook()

    # Monta a tabela
    planilha = planilha_excel.active
    planilha.append(['Nome da sua coluna'])

    for linha in executa_dados_para_excel:
        planilha.append(list(linha))

    # Define a largura das colunas
    for coluna in range(1, 10):
        letra_coluna = chr(64 + coluna)
        largura = len(planilha.cell(row=1, column=coluna).value) + 5
        planilha.column_dimensions[letra_coluna].width = largura

    # Salva o arquivo
    dt_hoje = datetime.date.today()
    nome_arquivo = f'nome_do_seu_excel.xlsx'
    caminho = os.path.abspath("caminho para salvar o documento, como exemplo C:\documento\  ")
    planilha_excel.save(os.path.join(caminho, nome_arquivo))

    # Fecha o arquivo
    planilha_excel.close()
    
    
def main():
    cria_excel()

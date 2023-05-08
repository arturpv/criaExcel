def criaExcel():
    query = 		f'''  
                             SELECT
                                [COLUNA]
                             FROM TB_ANALITICO_MONITORIA WITH(NOLOCK)   
                             WHERE [COLUNA] = '2023-03-06T15:38'


                              '''
    conexaoBanco = acoesBancoDados.acaoBancoDados('server', 'banco', 'login','senha', query)
    executaDadosParaExcel = conexao.executa()

    # GERA O ARQUIVO EM EXCEL
    arquivo_excel = Workbook()

    # Monta a tabela
    planilha = arquivo_excel.active
    planilha.append(['Nome da sua coluna'])

    for linha in acompanhamentoGeral:
        planilha.append(list(linha))

    cont = 1
    while cont < 10:
        column = str(chr(64 + cont))
        tamanho = len(planilha.cell(row=1, column=cont).value) + 5
        planilha.column_dimensions[column].width = tamanho
        cont += 1

    # Salva o arquivo
    dtHoje = datetime.date.today()
    nomeArquivo = f'nome_do_seu_excel.xlsx'
    caminho = os.path.abspath("caminho para salvar o documento, como exemplo C:\documento\  ")
    arquivo_excel.save(f'{caminho}\{nomeArquivo}')
    arquivo_excel.close()

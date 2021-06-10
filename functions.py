import pandas as pd
from datetime import datetime
from time import sleep

tabela = pd.read_excel(r'C:\Users\bruno\Documents\ex_1_moore\database.xlsx', sheet_name='bd')

def extratificao_valor():
    print('-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=')
    print('\n        Extratificação dos dados por valor')
    sleep(1.5)
    valor_total_departamento = tabela[['NOME_DEPARTAMENTO', 'VALOR']].groupby('NOME_DEPARTAMENTO').sum('VALOR')
    valor_total_departamento = valor_total_departamento.sort_values(by='VALOR', ascending=False)
    print(valor_total_departamento)
    print('-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=')
    sleep(1.5)
        

def aging_list():
    print('-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=')
    print('\n        Aging List: ')
    sleep(1.5)
    for i in range(tabela['DTA_EMISSAO'].count()):
        data_emissao = tabela['DTA_EMISSAO'][i].date()
        data_venc = tabela['DTA_VENCIMENTO'][i].date()
        data_emissao_convertida = datetime.strptime(str(data_emissao), '%Y-%m-%d').date()
        data_venc_convertida = datetime.strptime(str(data_venc), '%Y-%m-%d').date()
        dias_totais = (data_venc_convertida - data_emissao_convertida).days
        nome_departamento = tabela['NOME_DEPARTAMENTO'][i]
        valor_recebimento = tabela['VALOR'][i]
        cliente = tabela['CLIENTE'][i]
        print(f'Faltam {dias_totais} dias para o Departamento {nome_departamento} receber R${valor_recebimento} do cliente {cliente}')
        sleep(.5)
    print('-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=')
    sleep(1.5)
    

def analise_estatistica():
    print('-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=')
    print('\n        Média de valor por departamento')
    media_departamento = tabela[['NOME_DEPARTAMENTO', 'VALOR']].groupby('NOME_DEPARTAMENTO').mean('VALOR')
    media_departamento = media_departamento.sort_values(by='VALOR', ascending=False)
    print(f'\n      {media_departamento}')
    
    sleep(1.5)
    print('\n        Mediana de valor por departamento')
    mediana_departamento = tabela[['NOME_DEPARTAMENTO', 'VALOR']].groupby('NOME_DEPARTAMENTO').median('VALOR')
    mediana_departamento = mediana_departamento.sort_values(by='VALOR', ascending=False)
    print(f'\n      {mediana_departamento}')
    print('-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=')
    sleep(1.5)


def programa():
    while True:
        print('\033[1;33m -=-=-=-=-=-=-=-=-=-=-=-=-= SISTEMA MOORE -=-=-=-=-=-=-=-=-=-=-=-=-= ')
        while True:   
            try:
                n = int(input('''\n\033[1;37m        ESCOLHA UMA OPÇÃO: \n 
            [1] PARA EXTRATIFICAÇÃO DA TABELA POR VALOR \n 
            [2] PARA AGING LIST \n 
            [3] PARA ANÁLISES ESTATISTICAS \n
            [4] PARA SAIR \n
            INSIRA AQUI: '''))
                break
            except ValueError:
                print('\n\n         Digite um valor válido! ')

        if n == 1:
            extratificao_valor()
        elif n == 2:
            aging_list()
        elif n == 3:
            analise_estatistica()
        elif n == 4:
            break
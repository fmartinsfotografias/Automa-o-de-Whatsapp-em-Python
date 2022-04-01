import requests
import json
from time import sleep
from datetime import datetime
from openpyxl import *


def cotacao():
    #API Cotação de valores
    cotações = requests.get("https://economia.awesomeapi.com.br/last/USD-BRL,EUR-BRL,BTC-BRL")
    cotações = cotações.json()
    cotações_dolar = float(cotações['USDBRL']["bid"])
    cotações_Euro = float(cotações['EURBRL']["bid"])
    cotações_BitCoin = float(cotações['BTCBRL']["bid"])


    print('COTAÇÕES MONETARIAS'.center(60))
    print('='*60)
    print(f'Dolar agora: R$ {cotações_dolar:.2f}')
    print(f'Euro agora: R$ {cotações_Euro:.2f}')
    print(f'BitCoin agora: R$ {cotações_BitCoin:.3f}')

    sleep(1)

    link_Ip = f"https://api.hgbrasil.com/weather?key=f3f2a093&user_ip=remote"
    clima = requests.get(link_Ip)
    clima = clima.json()

    # VariaveisClima
    cidade_clima = clima['results']['city']
    cidade_data = clima['results']['date']
    cidade_hora = clima['results']['time']
    cidade_temp = clima['results']['temp']
    cidade_descrição = clima['results']['description']
    cidade_humidade = clima['results']['humidity']
    cidade_vento = clima['results']['wind_speedy']
    cidade_NascerDoSol = clima['results']['sunrise']
    cidade_PorDoSol = clima['results']['sunset']

    cidade_Proximos_Dias = clima['results']['forecast']

    print('=' * 60)
    print('CLIMA LOCAL'.center(60))
    print('=' * 60)
    print(f'Cidade: {cidade_clima}')
    print(f'Data: {cidade_data}')
    print(f'Hora: {cidade_hora}')
    print(f'Temperatura: {cidade_temp}º')
    print(f'Descrição: {cidade_descrição}')
    print(f'Humidade: {cidade_humidade}%')
    print(f'Velocidade do vento: {cidade_vento}')
    print(f'Nascer do sol: {cidade_NascerDoSol}')
    print(f'Por do Sol: {cidade_PorDoSol}')
    sleep(1)
    print('=' * 60)
    print('PROXIMOS DIAS'.center(60))
    print('=' * 60)
    sleep(1)

    for c in cidade_Proximos_Dias:
        print(f'{c}  ')
    print('='*60)
    print(f'Enviando Relatorio...')


    wb = Workbook()
    sh = wb.active
    wb.save(filename='relatorios.xlsx')

    wb = Workbook()
    sh = wb.active
#NOMES
    sh['A1'] = 'NOME'
    sh['A2'] = 'Fernando'
    sh['A3'] = 'Najela'
    sh['A4'] = 'Erison'
    sh['A5'] = 'Miriam'
#TELEFONES
    sh['B1'] = 'TELEFONE'
    sh['B2'] = 5547999214977
    sh['B3'] = 5547988188395
    sh['B4'] = 5547992910815
    sh['B5'] = 5547992498797

#MENSAGENS
    sh['D1'] = 'MENSAGEM'
    sh['D2'] = f"'DOLAR R$: '{cotações_dolar} 'EURO R$: '{cotações_Euro} 'BTC R$: '{cotações_BitCoin} " \
               f"'CIDADE:' {cidade_clima} 'DATA: '{cidade_data} 'TEMPERATURA:' {cidade_temp} " \
               f"'HUMIDADE: '{cidade_humidade} 'HORA: '{cidade_hora}"
    sh['D3'] = f"'DOLAR R$: '{cotações_dolar} 'EURO R$: '{cotações_Euro} 'BTC R$: '{cotações_BitCoin} " \
               f"'CIDADE:' {cidade_clima} 'DATA: '{cidade_data} 'TEMPERATURA:' {cidade_temp} " \
               f"'HUMIDADE: '{cidade_humidade} 'HORA: '{cidade_hora}"
    sh['D4'] = f"'DOLAR R$: '{cotações_dolar} 'EURO R$: '{cotações_Euro} 'BTC R$: '{cotações_BitCoin} " \
               f"'CIDADE:' {cidade_clima} 'DATA: '{cidade_data} 'TEMPERATURA:' {cidade_temp} " \
               f"'HUMIDADE: '{cidade_humidade} 'HORA: '{cidade_hora}"
    sh['D5'] = f"'DOLAR R$: '{cotações_dolar} 'EURO R$: '{cotações_Euro} 'BTC R$: '{cotações_BitCoin} " \
               f"'CIDADE:' {cidade_clima} 'DATA: '{cidade_data} 'TEMPERATURA:' {cidade_temp} " \
               f"'HUMIDADE: '{cidade_humidade} 'HORA: '{cidade_hora}"
    wb.save(filename='relatorios.xlsx')

    return cotacao

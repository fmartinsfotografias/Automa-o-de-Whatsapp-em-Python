from cotacao import cotacao
from time import sleep
import selenium
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from selenium import webdriver
import time
import urllib
import pandas as pd

print('=' * 60)
print('INFORMAÇÕES'.center(60))
print('=' * 60)
while True:
    sleep(1)
    cotacao()
    sleep(1)

    #Fonte do Relatorio
    data = pd.read_excel("relatorios.xlsx")

    #Navegação
    navegador = webdriver.Chrome()
    navegador.get("https://web.whatsapp.com/")

    while len(navegador.find_elements(By.ID,"side")) < 1:
        time.sleep(1)

    # já estamos com o login feito no whatsapp web
    for i, mensagem in enumerate(data['MENSAGEM']):
        nome = data.loc[i, "NOME"]
        numero = data.loc[i, "TELEFONE"]
        texto = urllib.parse.quote(f"Oi {nome}! {mensagem}")
        link = f"https://web.whatsapp.com/send?phone={numero}&text={texto}"
        navegador.get(link)
        while len(navegador.find_elements(By.ID, "side")) < 1:
            time.sleep(7)
        navegador.find_element(By.XPATH, '//*[@id="main"]/footer/div[1]/div/span[2]/div/div[2]/div[1]/div/div[2]').send_keys(Keys.ENTER)
        time.sleep(10)
    navegador.close()
    time.sleep(600)









#!/usr/bin/env python
# coding: utf-8

# In[21]:


#1 -importar as bibliotecas
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
servico = Service(ChromeDriverManager().install())
options = webdriver.ChromeOptions()
options.add_argument('--headless') # modo headless

import pandas as pd
import win32com.client as win32
from datetime import datetime
import time
from PIL import Image
import pyautogui as gui

#from pathlib import Path
import os


# In[2]:


#2 - acessar o site da fundamentus
browser = webdriver.Chrome(service=servico, options=options)
browser.get('https://www.fundamentus.com.br/detalhes.php?papel=RBRR11')


# In[23]:


#3 - acessar a base de dados .xlsx
df = pd.read_excel(r'C:\Users\W10\Desktop\Python\Arquivos_estudo\Projetos\Meus_projetos\Carteira Fiis\Projeto - Carteira de fiis.xlsx')
#display(df)


# In[4]:


#3 - pegar as informações dos fiis
df['Cotação'] = 0
df['Dividendo por cota'] = 0
df['Dividend Yield'] = 0
df['Oscilação do mês'] = 0
df['Oscilação 30 dias'] = 0
df['Oscilação 12 meses'] = 0

for linha in df.index:
    fii = df.loc[linha, 'Fiis']
    browser.find_element(By.XPATH, '//*[@id="completar"]').send_keys(f'{fii}', Keys.ENTER)
    valor_cota = browser.find_element(By.XPATH, '/html/body/div[1]/div[2]/table[1]/tbody/tr[1]/td[4]/span').text
    dividendos_cota = browser.find_element(By.XPATH, '/html/body/div[1]/div[2]/table[3]/tbody/tr[3]/td[6]/span').text
    dividend_yield = browser.find_element(By.XPATH, '/html/body/div[1]/div[2]/table[3]/tbody/tr[3]/td[4]/span').text
    oscilacao_mes = browser.find_element(By.XPATH, '/html/body/div[1]/div[2]/table[3]/tbody/tr[3]/td[2]/span').text
    oscilacao_30dias = browser.find_element(By.XPATH, '/html/body/div[1]/div[2]/table[3]/tbody/tr[4]/td[2]/span/font').text
    oscilacao_12meses = browser.find_element(By.XPATH, '/html/body/div[1]/div[2]/table[3]/tbody/tr[5]/td[2]/span/font').text
    
    df.loc[linha, 'Cotação'] = valor_cota
    df.loc[linha, 'Dividendo por cota'] = dividendos_cota
    df.loc[linha, 'Dividend Yield'] = dividend_yield
    df.loc[linha, 'Oscilação do mês'] = oscilacao_mes
    df.loc[linha, 'Oscilação 30 dias'] = oscilacao_30dias
    df.loc[linha, 'Oscilação 12 meses'] = oscilacao_12meses

browser.quit()
#display(df)
#df.info()


# In[5]:


#4 - tratamento da base de dados
df['Cotação'] = df['Cotação'].str.replace(',', '.').astype('float')
df['Dividendo por cota'] = df['Dividendo por cota'].str.replace(',', '.').astype('float')
df = df.set_index('Fiis')

#display(df)
#df.info()


# In[6]:


#5 - acrescentar no df: valor atual de todo o investimento
df['Valor Atual'] = df['Cotas'] * df['Cotação']

valor_total = df['Valor Atual'].sum()
#print(f'Valor total atual da carteira: R$ {valor_total:,.2f}.')


# In[7]:


#6 - acrescentar no df: valor esperado de dividendos
df['Valor Dividendos Anual'] = df['Cotas'] * df['Dividendo por cota']

dividendos_total = (df['Valor Dividendos Anual'].sum()) / 12
#print(f'Valor total de dividendos esperado por mês: R$ {dividendos_total:,.2f}.')


# ###### Importação dos gráficos dos históricos de cotações

# In[18]:


gui.alert('NÃO TOCAR NO MOUSE OU TECLADO')
browser = webdriver.Chrome(service=servico)

# criar função para printar e editar a imagem

def capturar_imagem(periodo_do_grafico):
    #printar a página
    time.sleep(1)
    browser.save_screenshot(fr'C:\Users\W10\Desktop\Python\Arquivos_estudo\Projetos\Meus_projetos\Carteira Fiis\gráficos cotações\{fii} - {periodo_do_grafico}.png')
    #abrir, editar, salvar o print
    imagem = Image.open(fr'C:\Users\W10\Desktop\Python\Arquivos_estudo\Projetos\Meus_projetos\Carteira Fiis\gráficos cotações\{fii} - {periodo_do_grafico}.png')
    left = 500
    top = 200
    right = 1320
    bottom = 800
    imagem = imagem.crop((left, top, right, bottom))
    imagem.save(fr'C:\Users\W10\Desktop\Python\Arquivos_estudo\Projetos\Meus_projetos\Carteira Fiis\gráficos cotações\{fii} - {periodo_do_grafico}.png')

# abrir os gráficos de cada fii (1 mês, 3 meses, 6 meses, 1 ano) e aplicar a função
    
for fii in df.index:
    
    browser.get(f'https://www.fundamentus.com.br/cotacoes.php?papel={fii}')
    browser.maximize_window() 
    time.sleep(3)

    #gui.click(x=570, y=350)
    #capturar_imagem('1m')

    #gui.click(x=609, y=354)
    #capturar_imagem('3m')

    gui.click(x=641, y=352)
    capturar_imagem('6m')
    
    #gui.click(x=722, y=355)
    #capturar_imagem('1y')
        
browser.quit()


# In[22]:


# enviar email
data = datetime.now().strftime('%d/%m/%Y')

outlook = win32.Dispatch('outlook.application')
e = outlook.CreateItem(0)
e.To = 'bep_rafael@hotmail.com'
e.Subject = f'Análise da Carteira de Fiis - {data}'
e.HTMLBody = f'''
<p>{df.to_html()}</p>
<p>Valor total atual da carteira:<strong> R$ {valor_total:,.2f}</strong></p>
<p>Valor total de dividendos esperado por mês:<strong> R$ {dividendos_total:,.2f}</strong></p>
'''
path = r'C:\Users\W10\Desktop\Python\Arquivos_estudo\Projetos\Meus_projetos\Carteira Fiis\gráficos cotações'
lista_arquivos = os.listdir(path)
#pasta_imagens = Path(r'C:\Users\W10\Desktop\Python\Arquivos_estudo\Projetos\Meus_projetos\Carteira Fiis\gráficos cotações')
for arquivo in lista_arquivos:
    caminho_anexo = fr"C:\Users\W10\Desktop\Python\Arquivos_estudo\Projetos\Meus_projetos\Carteira Fiis\gráficos cotações\{arquivo}"
    e.Attachments.Add(str(caminho_anexo)) 
e.Send()

gui.alert('FIM DA AUTOMAÇÃO :D')


# In[ ]:





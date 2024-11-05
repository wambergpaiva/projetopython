#!/usr/bin/env python
# coding: utf-8

# # Automação Web e Busca de Informações com Python
# 
# #### Desafio: 
# 
# Trabalhamos em uma importadora e compramos e vendemos commodities:
# - Soja, Milho, Trigo, Petróleo, etc.
# 
# Precisamos pegar na internet, de forma automática, a cotação de todas as commodites e ver se ela está abaixo do nosso preço ideal de compra. Se tiver, precisamos marcar como uma ação de compra para a equipe de operações.
# 
# Base de Dados: https://drive.google.com/drive/folders/1O9aE_Hen5smf_a6TsbVF6DM0ms4Ykj6s?usp=share_link
# 
# Para isso, vamos criar uma automação web:
# 
# - Usaremos o selenium
# - Importante: baixar o webdriver

# In[43]:


# firefox -> geckodriver
# chrome -> chmedriver
# edge -> msedgedrive

from selenium import webdriver
import pyautogui
import pyperclip
import time

#passo a passo
# Passo 1: Abrir o navegador
from selenium import webdriver

navegador = webdriver.Edge()
navegador.get("https://www.google.com.br/")


# In[44]:


# Passo 2: Importar a base de dados
import pandas as pd
df = pd.read_excel('commodities.xlsx')
display(df)


# In[45]:


for linha in df.index:
    produto = df.loc[linha, 'Produto']
    print(produto)
    produto = produto.replace('ó', 'o').replace('á', 'a').replace('ã', 'a').replace('ç', 'c').replace('é','e').replace('ú', 'u')

    link = f'https://www.melhorcambio.com/{produto}-hoje'
    print(link)
    

    navegador.get(link)


    preco = navegador.find_element('xpath', '//*[@id="comercial"]').get_attribute('value')
    preco = preco.replace('.', '').replace(',', '.')
    print(preco)
    df.loc[linha, 'Preço Atual'] = float(preco)

print('Acabou')
display(df)
    

# .click() -> clicar
# .send_keys('texto') -> escrever
# .get_atribute() -> pegar um valor

# Passo 3: Para cada produto da base de dados
# Passo 4: Pesquisar o preço do produto
# Passo 4: Atualizar o preço da base de dados
# Passo 6: Decidir quais predutos comprar


# In[46]:


#Preencher a coluna de comprar

df['Comprar'] = df['Preço Atual'] <= df['Preço Ideal']
display(df)

#Exportar a base de dados para o excel
df.to_excel('commodities_atualizado.xlsx', index=False)



# In[47]:


import pyautogui
import pyperclip
import pandas as pd
import time

# Minimizar o navegador 
pyautogui.click(x=836, y=32)
pyautogui.sleep(3)

# Leitura da planilha Excel com as cotações atuais
df = pd.read_excel('commodities_atualizado.xlsx')

# Abrir uma nova aba no navegador e acessar o Gmail
pyautogui.hotkey('ctrl', 't')
pyperclip.copy('https://mail.google.com/mail/u/0/#inbox')
pyautogui.hotkey('ctrl', 'v')
pyautogui.press('enter')
time.sleep(5)

# Clicar no botão "Escrever" no Gmail (ajuste coordenadas conforme necessário)
pyautogui.click(x=60, y=144)  # Coordenadas ajustáveis conforme sua tela

time.sleep(2)  # Aguardar o Gmail carregar o campo "Escrever"

# Digitar o e-mail do destinatário
pyperclip.copy('tricolorflu34+diretoria@gmail.com')  # Usar pyperclip para colar o e-mail
pyautogui.hotkey('ctrl', 'v')
pyautogui.press('tab')  # Seleciona o e-mail

# Pular para o campo do assunto
pyperclip.copy('Relatório de Commodities Atualizado')
pyautogui.hotkey('ctrl', 'v')  # Colar o assunto

# Pausa para garantir que o campo do assunto foi preenchido
time.sleep(1)

# Clicar no campo do corpo do e-mail (ajuste as coordenadas conforme necessário)
pyautogui.click(x=1562, y=768)  # Coordenadas ajustáveis para o campo do corpo do e-mail

# Preparar o texto do corpo do e-mail com os dados
texto = f"""
Prezados, bom dia

Segue abaixo o relatório atualizado de commodities com os preços ideais de compra:

"""

# Adicionar os produtos e indicar se é para comprar ou não
for index, row in df.iterrows():
    if row['Preço Atual'] <= row['Preço Ideal']:
        status = "COMPRAR"
    else:
        status = "NÃO COMPRAR"
    
    # Montar a linha do e-mail com o status
    texto += f"Produto: {row['Produto']} - Preço Atual: R${row['Preço Atual']} - Preço Ideal: R${row['Preço Ideal']} - {status}\n"

texto += """

Atenciosamente,
Wamberg Paiva - Engenheiro de Produção
"""

# Copiar e colar o texto no corpo do e-mail
pyperclip.copy(texto)
pyautogui.hotkey('ctrl', 'v')

# Enviar o e-mail (atalho de teclado: Ctrl + Enter)
pyautogui.hotkey('ctrl', 'enter')


# In[48]:


pyautogui.sleep(5)
pyautogui.position() 


# In[49]:


get_ipython().system('pip install selenium')


# In[ ]:





# In[ ]:





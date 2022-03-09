#!/usr/bin/env python
# coding: utf-8

# # Automação de Sistemas e Processos com Python
# 
# ### Desafio:
# 
# Todos os dias, o nosso sistema atualiza as vendas do dia anterior.
# O nosso trabalho diário, como analista, é enviar um e-mail para a diretoria, assim que começar a trabalhar, com o faturamento e a quantidade de produtos vendidos no dia anterior. Essa tarefa é repetitiva e não agrega valor ao processo. Portanto, uma solução seria *automatizá-lo*.
# 
# Para resolver isso, vamos usar o pyautogui, uma biblioteca de automação de comandos do mouse e do teclado, além das bibliotecas pandas, time e piperclip. 
# 
# ### O que iremos fazer:
# 
# 1. Entrar no sistema da empresa para acessar a BD.
# 2. Navegar no sistema até a BD.
# 3. Exportar (fazer download) e importar a BD.
# 4. Calcular os indicadores (faturamento total e quantidade de produtos vendidos).
# 5. Enviar o email à diretoria com os indicadores.
# 
# ### Importando as bibliotecas que serão usadas:
# 
# * pyautogui: ferramenta de automação de mouse, teclado e tela. 
# * pyperclip: permite copiar e colar via python.
# * time: facilita no desenvolvimento de códigos que envolvem tempo de espera.

# In[63]:


get_ipython().system('pip install pyautogui')
get_ipython().system('pip install pyperclip')


# In[64]:


import pyautogui
import pyperclip
import time
import pandas as pd


# <b> 
#     
# 1. Entrando no sistema
#     
# 2. Navegar no sistema até a BD
#     
# 3. Exportar (fazer download) e importar a BD

# In[69]:


pyautogui.PAUSE = 1 # intervalo de 1seg entre comandos
pyautogui.alert("Vai começar, aperte OK!")

# Abrir um navegador de internet usando o teclado

pyautogui.press("winleft") # clicando no botão windows 
pyautogui.write("chrome") # abrindo o navegador
pyautogui.press("enter")
pyautogui.click(x=844, y=629, clicks = 1) # clicando no meu usuário chrome
pyautogui.hotkey('ctrl', 't') # abrindo nova guia

# Copiar link do Drive onde está nossa bd usando o teclado
# Entrar no drive onde está a bd usando o teclado

link = "https://drive.google.com/drive/folders/149xknr9JvrlEnhNWO49zPcw0PW5icxga"
pyperclip.copy(link) # armazenando na memória o valor da variável link (ctrl+c)
pyautogui.hotkey('ctrl', 'v') # colando no navegador
pyautogui.press("enter")
time.sleep(3) # o computador vai aguardar 7seg antes de executar a próxima linha de código possibilitando carregar a página

# Abrir a pasta do drive usando o mouse 

pyautogui.click(x=384, y=294, clicks = 2) # clicando na pasta
time.sleep(3)

# Baixar planilha Vendas - Dez.xlsx usando mouse 

pyautogui.click(x=420, y=303, clicks = 1, button = "right")
pyautogui.click(x=578, y=877, clicks = 1) # fazer download
time.sleep(3)


# ### Vamos agora ler o arquivo baixado para pegar os indicadores
# 
# - Faturamento
# - Quantidade de Produtos
# 
# 

# In[70]:


# Lendo o arquivo 

vendas = pd.read_excel(r'C:\Users\fabiabc\Downloads\Vendas - Dez.xlsx')
                       # r indica que o texto é uma raw string, ou seja, sem caracteres especiais
display(vendas)

# Calculando os indicadores

faturamento = vendas['Valor Final'].sum()
print((f"""O faturamento total é R${faturamento:,.2f}"""))

qtde_produtos = vendas['Quantidade'].sum()
print((f"""A quantidade de produtos é R${qtde_produtos:,}"""))


# ### Vamos agora enviar um e-mail pelo gmail

# In[71]:


# Abrir um nova guia no navegador 

pyautogui.hotkey('ctrl','t')

# Acessar o gmail 

pyautogui.write("https://mail.google.com/mail/u/0/?tab=rm&ogbl#inbox")
pyautogui.press('enter')
time.sleep(3)

# Escrever um novo email 

pyautogui.click(x=52, y=227, clicks = 1)
pyautogui.write('fabia')
pyautogui.press('tab')
pyautogui.press('tab')
assunto = "Relatório de vendas de dezembro"
pyperclip.copy(assunto)
pyautogui.hotkey('ctrl', 'v')
    
pyautogui.press('tab')
corpo = f"""
Prezados, bom dia

O faturamento de ontem foi de R${faturamento:,.2f}.
A quantidade de produtos vendidos foi {qtde_produtos:,}.

Abrçs,
Fabia"""
pyperclip.copy(corpo)
pyautogui.hotkey('ctrl', 'v')
pyautogui.hotkey('ctrl', 'enter')


# #### Use esse código para descobrir qual a posição de um item que queira clicar
# 
# - Lembre-se: a posição na sua tela é diferente da posição na minha tela

# In[72]:


time.sleep(5)
pyautogui.position()


# %% [markdown]
# # Automação de Sistemas e Processos com Python
# 
# ### Desafio:
# 
# Todos os dias, o nosso sistema atualiza as vendas do dia anterior.
# O seu trabalho diário, como analista, é enviar um e-mail para a diretoria, assim que começar a trabalhar, com o faturamento e a quantidade de produtos vendidos no dia anterior
# 
# E-mail da diretoria: seugmail+diretoria@gmail.com<br>
# Local onde o sistema disponibiliza as vendas do dia anterior: Aqui você bota o link do drive com seu relatorio
# 
# Para resolver isso, vamos usar o pyautogui, uma biblioteca de automação de comandos do mouse e do teclado
# 
# 
# Referência do pyautogui: https://pyautogui.readthedocs.io/en/latest/quickstart.html

# %%
import pyautogui
import time 
import pyperclip

# Para esse desafio usaremos as bibliotecas
# pyautogui -> para automações
# pandas -> dados
# openpyxl -> trabalha com arquivos Excel

# pyautogui.click -> clicar com o mouse
# pyautogui.write -> escrever um texto
# pyautogui.press -> apertar 1 tecla
# pyautogui.hotkey -> combinação de teclas

pyautogui.PAUSE = 0.5

# Passo a passo do desafio
# Passo 1: entrar no sistema da empresa
# Abri o chrome
pyautogui.press("win")
pyautogui.write("chrome")
pyautogui.press("enter")

time.sleep(2)

# Digitei o link do sistema
LINK_RELATORIO = "COLE_AQUI_O_LINK_DO_RELATORIO"
pyautogui.write(LINK_RELATORIO)

# Apertei enter e esperei
pyautogui.press("enter")
time.sleep(2)



# %% [markdown]
# ### Vamos agora ler o arquivo baixado para pegar os indicadores
# 
# - Faturamento
# - Quantidade de Produtos

# %%
# Passo 2: Navegar no sistema para encontrar a base de dados
time.sleep(2)
pyautogui.click(x=343, y=432, clicks=2)
time.sleep(1)

# Passo 3: Exportar a base de dados (baixar o arquivo)
pyautogui.click(x=343, y=432, clicks=1) # seleciona o arquivo
pyautogui.click(x=1268, y=426, clicks=1) # clica nos 3 pontos
time.sleep(1)
pyautogui.click(x=1139, y=360, clicks=1) #baixa o arquivo

import pandas

# Passo 4: Calcular os indicadores (faturamento e quantidade de produtos vendidos)
# Abrir a base de dados
from pathlib import Path

ARQUIVO = "Vendas - Dez.xlsx"
caminho = Path.home() / "Downloads" / ARQUIVO
tabela = pandas.read_excel(caminho)

# Ver informações da base de dados
display(tabela)

# Somar o faturamento de todos os produtos = somar a coluna valor final
faturamento = tabela["Valor Final"].sum()

# Somar a quantidade de produtos = somar coluna de quantidade
quantidade_de_produtos = tabela["Quantidade"].sum()



# %% [markdown]
# ### Vamos agora enviar um e-mail pelo gmail

# %%
# Passo 5: Enviar as informações por e-mail
# Abrir uma nova aba
pyautogui.hotkey("ctrl", "t")

# Entrar no email
pyautogui.write("http://mail.google.com/")
pyautogui.press("enter")
time.sleep(8)

# Clicar no botão escrever email
pyautogui.click(x=92, y=214)
time.sleep(5)

# Para quem enviar
DESTINATARIO = "email@exemplo.com"
pyautogui.write(DESTINATARIO)
pyautogui.press("tab")
pyautogui.press("tab")
time.sleep(2)

# Qual o asunto do email
pyperclip.copy("Relatório de vendas de hoje")
pyautogui.hotkey("ctrl", "v")
pyautogui.press("tab")


# Qual o corpo do email
texto = f"""
Prezados,

Segue em anexo o relatório de vendas de hoje.

Faturamento:R${faturamento:,.2f}
Quantidade de produtos vendidos:{quantidade_de_produtos:,}

Qualquer dúvida estou à disposição.
Abs,
Gabriel Ferreira
"""
pyperclip.copy(texto)
pyautogui.hotkey("ctrl", "v")

# Enviar o email
pyautogui.press("tab")
pyautogui.press("enter")
pyautogui.press("enter")

# %% [markdown]
# #### Use esse código para descobrir qual a posição de um item que queira clicar
# 
# - Lembre-se: a posição na sua tela é diferente da posição na minha tela

# %%
time.sleep(5)
print(pyautogui.position())



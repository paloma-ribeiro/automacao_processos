"""
Automação de Sistemas e Processos com Python

Aula 01

Problema: Todos os dias, o nosso sistema atualiza as vendas do dia anterior. O seu trabalho,
como analista, é enviar um email para a diretoria, assim que começa a trabalhar, com o
faturamento e a quantidade de produtos vendidos no dia anterior.

E-mail da diretoria: pythonimpressionador+diretoria@gmail.com

Local onde o sistema disponibiliza as vendas do dia anterior:
https://drive.google.com/drive/folders/149xknr9JvrlEnhNWO49zPcw0PW5icxga

Para resolver isso, vamos usar o pyautogui, uma biblioteca de automação de comandos do mouse
e teclado.
"""

import pyautogui
import pyperclip
import time
import pandas as pd

pyautogui.PAUSE = 2  # Force a 2 seconds delay in the execution of the commands

# Passo 1: Entrar no sistema da empresa (no caso o Google drive)

pyautogui.press('win')  # Open the start menu
pyautogui.write('chrome')  # Open the browser Google Chrome
pyautogui.press('enter')  # Press enter
pyautogui.write(
    'https://drive.google.com/drive/folders/149xknr9JvrlEnhNWO49zPcw0PW5icxga')  # Write the address in the browser
pyautogui.press('enter')  # Press enter

# Outra maneira de fazer o passo 1, mencionado acima.

# pyautogui.hotkey('ctrl', 't')
# pyperclip.copy('https://drive.google.com/drive/folders/149xknr9JvrlEnhNWO49zPcw0PW5icxga')
# pyautogui.hotkey('ctrl', 'v')
# pyautogui.press('enter')

# Passo 2: Navegar no sistema até encontrar a base de dados
# OBS: pyautogui.position(): Get the position of the button on the screen

time.sleep(6)  # Add a 6 seconds pause in the execution of the commands (wait for the screen to load)
pyautogui.click(x=453, y=307, clicks=2)  # double-click on the screen, in the position informed (folder)

# Passo 3: Exportar a base de vendas

time.sleep(3)  # Add a 3 seconds pause in the execution of the commands (wait for the screen to load)
pyautogui.click(x=395, y=410)  # click on the screen, in the position informed (file)
pyautogui.click(x=1073, y=177)  # click on the screen, in the position informed (menu)
pyautogui.click(x=834, y=613)  # click on the screen, in the position informed (download)

# Passo 4: Calcular os indicadores (faturamento e quatidade de produtos vendidos)

time.sleep(7)  # Add a 7 seconds pause in the execution of the commands (wait for download)
tabela = pd.read_excel(
    r'C:\Users\palom\Downloads\Vendas - Dez.xlsx')  # read the downloaded file in the downloads folder

faturamento = tabela['Valor Final'].sum()  # sum of all values in the Valor Final column
quantidade = tabela['Quantidade'].sum()  # sum of all values in the Quantidade column

# Passo 5: Enviar um email para a diretoria com os indicadores

# Abrir o email no navegador
pyautogui.hotkey('ctrl', 't')  # open the new tab
pyperclip.copy('https://mail.google.com/mail/u/1/?pli=1#inbox')  # Copy the email link
pyautogui.hotkey('ctrl', 'v')  # Paste the email link
pyautogui.press('enter')  # Press enter
time.sleep(9)  # Add a 9 second pause in the execution of the commands (wait for the email to load)

# Clicar no botão escrever
pyautogui.click(x=114, y=191)  # click on the screen, in the position informed (new email)
time.sleep(4)

# Escrever destinatário (quem vai receber o email)
pyautogui.write('pythonimpressionador+diretoria@gmail.com')  # write the recipient's email address
pyautogui.press('tab')  # select recipient email

# Escrever o assunto
pyautogui.press('tab')  # jump to subject field
pyperclip.copy('Relatório de Vendas')  # write the subject of the email
pyautogui.hotkey('ctrl', 'v')
time.sleep(2)

# Escrever o corpo do email
pyautogui.press('tab')  # jump to email body

texto = f"""
Prezados, bom dia!

O faturamento de ontem foi de: R$ {faturamento:.2f}
A quantidade de produtos foi de: {quantidade}

Abs
"""

pyperclip.copy(texto)
pyautogui.hotkey('ctrl', 'v')
time.sleep(2)

# Enviar o email
pyautogui.hotkey('ctrl', 'enter')

import win32com.client as win32
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium import webdriver
import time
import urllib
import pandas as pd
import PySimpleGUI as sg

layout = [
    [sg.Text('Selecione o arquivo Excel', font=("Helvetica", 13),
             text_color='', justification='center')],
    [sg.InputText(size=(40, 1)), sg.FileBrowse(button_text='Procurar')],
    [sg.Text('Selecione o email para envio do relatorio', font=(
        "Helvetica", 13), text_color='', justification='center')],
    [sg.InputText(size=(40, 1))],
    [sg.Submit(button_text='Enviar', button_color=('white', 'green')),
     sg.Cancel(button_text='Cancelar', button_color=('white', 'red'))]
]

window = sg.Window('Arquivo Excel', layout)

while True:
    event, values = window.Read()
    if event is None or event == 'Enviar':
        if values[0] == '':
            sg.popup('Nenhum arquivo selecionado')
        else:
            filename = values[0]
        email_env = values[1]
        break
    elif event == 'Cancelar':
        break

window.close

contatos_df = pd.read_excel(filename)

# o resto do código aqui

navegador = webdriver.Chrome()

navegador.get("https://web.whatsapp.com/")

navegador.maximize_window()

while len(navegador.find_elements(By.ID, 'side')) < 1:
    time.sleep(5)

teste = ''

for i, mensagem in enumerate(contatos_df['Mensagem']):

    pessoa = contatos_df.loc[i, "Pessoa"]

    numero = contatos_df.loc[i, "Numero"]

    texto = urllib.parse.quote(f"Oi {pessoa}! {mensagem}")

    link = f"https://web.whatsapp.com/send?phone={numero}&text={texto}"

    navegador.get(link)

    time.sleep(5)

    try:

        if len(navegador.find_elements(By.CLASS_NAME, '_2Nr6U')) > 0:

            navegador.find_element(By.XPATH,
                                   '//*[@id="app"]/div/span[2]/div/span/div/div/div/div/div/div[2]/div/div/div/div').click()
            msg = f'Mensagem não enviada.'
            print(msg)
            time.sleep(5)
        else:

            navegador.find_element(By.XPATH,
                                   '//*[@id="main"]/footer/div[1]/div/span[2]/div/div[2]/div[1]/div/div/p/span').send_keys(Keys.ENTER)
            msg = f'Mensagem enviada com sucesso.'
            print(msg)
        time.sleep(5)

    except:
        print(f'Erro no envio')
        continue

    teste += f'<tr> <td>{pessoa}</td> <td>{numero}</td> <td>{msg}</td></tr>'

outlook = win32.Dispatch('outlook.application')

email = outlook.CreateItem(0)

email.To = str(email_env)

email.Subject = 'Relatório dos numeros enviados'

email.HTMLBody = f'''

<html>
    <head>
        <style>
            table, th, td {{
                border: 1px solid black;
                border-collapse: collapse;
            }}
            th, td {{
                padding: 8px;
                text-align: left;
            }}
            th {{
                background-color: lightgray;
            }}
        </style>
    </head>
    <body>
        <h1 style="text-align: center;">Relatório dos Números Enviados</h1>
        <br>
        <table style="width:100%;">
            <tr>
                <th>Nome</th>
                <th>Número</th>
                <th>Status de Envio</th>
            </tr>
            {teste}
        </table>
    </body>
</html>
'''

email.Send()

print('Email Enviado')

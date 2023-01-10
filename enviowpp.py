import win32com.client as win32
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium import webdriver
import time
import urllib
import PySimpleGUI as sg
import pandas as pd

# Função para importar o arquivo xlsx

try:
    def import_file():
        layout = [[sg.Text('Selecione o arquivo xlsx:'), sg.Input(), sg.FileBrowse()],
                  [sg.OK(), sg.Cancel()]]

        window = sg.Window('Importar arquivo').Layout(layout)
        event, values = window.Read()
        window.Close()

        if event == 'OK':
            filename = values[0]
            contatos_df = pd.read_excel(filename)
            return contatos_df
        else:
            return None

    # Função para exibir mensagem de erro

    def error(msg):
        sg.PopupError(msg)

    # Função para exibir a tela de relatório

    def show_report(teste):
        layout = [[sg.Text('Relatório dos números enviados')],
                  [sg.Text('Nome'), sg.Text('Numero'),
                   sg.Text('Status de Envio')],
                  [sg.Listbox(values=teste, size=(50, 20))],
                  [sg.OK()]]

        window = sg.Window('Relatório').Layout(layout)
        event, values = window.Read()
        window.Close()

    def main():
        global contatos_df
        contatos_df = import_file()
        if contatos_df is None:
            return

except:
    print("Algo deu errado")

finally:

    main()

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
                msg = f'A mensagem enviada para o numero {numero} nao foi efetuada.'
                print(msg)
                time.sleep(5)
            else:

                navegador.find_element(By.XPATH,
                                       '//*[@id="main"]/footer/div[1]/div/span[2]/div/div[2]/div[1]/div/div/p/span').send_keys(Keys.ENTER)
                msg = f'A mensagem enviada para o numero {numero} foi efetuada.'
                print(msg)
            time.sleep(5)

        except:
            print(f'Erro no envio')
            continue

        teste += f'<tr> <td>{pessoa}</td> <td>{numero}</td> <td>{msg}</td></tr>'

    outlook = win32.Dispatch('outlook.application')

    email = outlook.CreateItem(0)

    email.To = 'digomesrique@hotmail.com'

    email.Subject = 'Relatório dos numeros enviados'

    email.HTMLBody = f'''
        
    <html>
            <body>
                <h1>Relatorio dos numeros enviados</h1>
                <br>
                <table border="1">
        <tr>
            <td>Nome</td>
            <td>Numero:</td>
            <td>Status de Envio:</td>
            
        </tr>
        {teste}
    
    </table>
            </body>
        </html>
        
        '''
    email.Send()

    print('Email Enviado')

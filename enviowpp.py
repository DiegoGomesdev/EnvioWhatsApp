import urllib
import time
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
import pandas as pd
import win32com.client as win32

contatos_df = pd.read_excel("Enviar.xlsx")
print(contatos_df)

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

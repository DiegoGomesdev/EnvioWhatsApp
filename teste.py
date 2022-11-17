import urllib
import time
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from IPython.display import display
import pandas as pd
import win32com.client as win32

contatos_df = pd.read_excel("enviar3.xlsx")


# quantidade de produtos vendidos por loja
pessoa = contatos_df['Pessoa'].to_list()
print(pessoa)

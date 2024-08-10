"""AUTOMATIZAR CHAT WHATSAPP """

import openpyxl
from urllib.parse import quote
import webbrowser
from time import sleep
import pyautogui
import os 


#CASO NECESSARIO, UTILIZE O CODIGO ABAIXO PARA ABRIR O SITE ANTES DE ENVIAR AS MENSAGENS DANDO TEMPO PARA FAZER O LOGIN NO WHATSAPP WEB 

#webbrowser.open('https://web.whatsapp.com/')
#sleep(5)

#LER PLANILHA

workbook = openpyxl.load_workbook('clientes.xlsx')
pagina_clientes = workbook['Sheet1']

for linha in pagina_clientes.iter_rows(min_row=2):
  #nome, telefone
  nome = linha[0].value
  telefone = linha[1].value
   
  mensagem = f'Olá {nome} Hoje temos uma promoção que é a sua cara, venha conferir no link do ifood. https://www.ifood.com.br/delivery/campinas-sp/pizzaria-da-mama-nossa-senhora-aparecida/569f4917-f837-4d09-a655-54acbd137bb0'

try:
  link_mensagem_whatsapp = f'https://web.whatsapp.com/send?phone={telefone}&text={quote(mensagem)}'
  webbrowser.open(link_mensagem_whatsapp)
  (sleep(5))
  seta = pyautogui.locateCenterOnScreen('seta.png')
  sleep(3)
  pyautogui.click(seta[0],seta[1])
  sleep(3)
  pyautogui.hotkey('ctrl','w')
  sleep(3)
except:
  print(f'Não foi possivel enviar mensagem para {nome}')
  with open('erros.csv','a',newline='',encoding='utf-8') as arquivo:
       arquivo.write(f'{nome},{telefone}')

#CRIAR PROMOÇÃO DE PIZZAS E ENVIAR  PARA CLIENTES
#COM BASE NOS DADOS DA PLANILHA

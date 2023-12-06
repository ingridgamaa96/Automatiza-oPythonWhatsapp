"""
Preciso automatizar minhas mensagens para meus clientes, para poder mandar mensagens de cobrança em determinado dia com dias de vencimento diferentes

"""

# descrever os passos manuais e depois transformar isso em código
import openpyxl
from urllib.parse import quote
import webbrowser
from time import sleep
from datetime import datetime
import pyautogui

webbrowser.open('https://web.whatsapp.com/')
sleep(30)
#ler planilha e guardar informações sobre nome, telefone e  data vencimento
workbook = openpyxl.load_workbook('clientes.xlsx.xlsx')
paginas_clientes = workbook['Planilha1']

for linha in paginas_clientes.iter_rows(min_row=2):
    #nome, telefone, vencimento

    #com base nos dados da planilha
    nome = linha[0].value
    telefone = linha[1].value
    dataVencimento = linha[2].value

    dataVencimento = datetime.now()  # Supondo que dataVencimento seja um objeto datetime
    mensagem = f'Olá {nome} seu boleto vence no dia {dataVencimento.strftime("%d/%m/%Y")} evite bloqueios'

   # mensagem = f'Olá {nome} seu boleto vence no dia {dataVencimento.strftime("%v/%m/%Y")} evite bloqueios'

    # criar links personalizados do whatsapp e enviar mensagens para cada cliente
    link_mensagem_whatsapp = f'https://web.whatsapp.com/send?phone={telefone}&text={quote(mensagem)}'
    webbrowser.open(link_mensagem_whatsapp)
    sleep(15)
    try:
            
        botaoEnviar = pyautogui.locateAllOnScreen('botaoEnviar.png')
        sleep (6)
        pyautogui.click(botaoEnviar[0],botaoEnviar[1]) 
        sleep(5)
        pyautogui.hotkey('ctrl', 'w')
        sleep(5)
    except:
        print (f 'Não foi possível enviar mensagem para {nome}')
        with open ('errros.csv', 'a', newline='',encoding='utf-8') as arquivo:
            arquivo.write(f'{nome}, {telefone}')


 





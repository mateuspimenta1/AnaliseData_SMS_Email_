from selenium import webdriver
from selenium.webdriver.common.by import By
import time
import pandas as pd
import pyautogui as pag
import openpyxl
import twilio
from twilio.rest import Client


#ATENÇÃO!!!!
#NECESSÁRIO CONTA NO TWILLIO.COM
#P/ O ALGORITIMO FUNCIONAR, LEMBRE-SE DE TROCAR AS INFORMAÇÕES DO:
#TWILLIO: ACCOUT_SID, AUTH_TOKEN, NUMERO VIRTUAL, E SEU NUMERO DE TESTE.
#GMAIL ACCOUNT: LOGIN E SENHA DO SEU EMAIL DE ENVIO, E O EMAIL TESTE P/ RECEBIMENTO.

#$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$



#Troque os valores do token e SID para os valores da sua propia conta twilio.
account_sid = "xxxxxx-xxxxx-xxxxxxx-xxxxxx"  #TROQUE AQUI.
auth_token = "xxxxx-xxxxx-xxxx-xxxxx-xxxx"  #TROQUE AQUI.
client = Client(account_sid, auth_token)

#Define a lista pra começar a ler a base de dados e imprimir no terminal.
listaMeses = ['Janeiro', 'Fevereiro', 'Março', 'Abril', 'Maio', 'Junho']
for mes in listaMeses:
    tabelaVendas = pd.read_excel(f'{mes}.xlsx')
    print(tabelaVendas)

    #Localizador feito pra achar algum possivel vendedor que tenha Vendido > que 50000,
    #no mes, e imprimir nome , mes e vendas que bateram a meta.
    if (tabelaVendas['Vendas'] > 50000).any():
        vendedor = tabelaVendas.loc[tabelaVendas['Vendas'] > 50000, 'Vendedor'].values[0]
        vendas = tabelaVendas.loc[tabelaVendas['Vendas'] > 50000, 'Vendas'].values[0]
        print(f'No mes {mes} alguém bateu a meta. Vendedor: {vendedor}, Vendas: {vendas}  \n')

        #Nesse momento, acessamos a biblioteca do Twillio pra criar e enviar
        # a SMS com as informações
        #Necessário utilizar suas informações do Twillio em from= e to=
        #encontrados no site que criou sua conta.
        #mesma tela do SID e Token.
        message = client.messages.create(
            body=f"No mes {mes} alguém bateu a meta.\nVendedor: {vendedor} \nVendas: {vendas}  \n",
            from_="+XXXXX-XXXXX",  #TROQUE AQUI NUMERO VIRTUAL
            to="+XXXXXX-XXXXX"  #TROQUE AQUI NUMERO TESTE
        )
    else:
        print(f'No mes {mes} não encontramos \n ____________________________ \n')


print("Sms enviada, preparando serviço web")
print('um momento.....')
time.sleep(5)
# Abre o Chrome pelo Edge, Entra no gmail, Loga Acc e Senha, abre a criação de email.
navegador = webdriver.Edge()
navegador.get("https://google.com")
time.sleep(3)
navegador.find_element(By.XPATH, '//*[@id="gb"]/div/div[1]/div/div[1]/a').click()
time.sleep(3)
pag.typewrite("XXXXX@gmail.com")   #TROQUE AQUI PRA EMAIL DE ENVIO DO GMAIL
navegador.find_element(By.XPATH, '//*[@id="identifierNext"]/div/button/span').click()
time.sleep(4)
pag.typewrite("XXXXXXXXX")   #TROQUE AQUI SENHA EMAIL ENVIO
navegador.find_element(By.XPATH, '//*[@id="passwordNext"]/div/button/span').click()
time.sleep(5)
navegador.find_element(By.XPATH, 'html/body/div[7]/div[3]/div/div[2]/div[1]/div[1]/div/div').click()
time.sleep(3)

# define endereço, Assunto, e Inicia o Cabeçalho do email.
pag.typewrite("xxxxxxxx@outlook.com") #Troque AQUI ENDEREÇO DE ENVIO
time.sleep(2)
pag.press('enter')
pag.press('tab')
time.sleep(2)
pag.typewrite("No-reply - Ti.empresarial")
time.sleep(2)
pag.press('tab')
time.sleep(2)
pag.typewrite(f'Relatório Semestral + Bonus Melhor Vendedor  \n')
pag.typewrite(f'No mes {mes} alguem bateu a meta.\n Vendedor: {vendedor} Vendas: {vendas}  \n')



#Loop for pra imprimir todas as planinhas, e somar as vendas.
totalvendas = 0
for mes in listaMeses:
    tabelaVendas = pd.read_excel(f'{mes}.xlsx')
    pag.typewrite(f'{mes}\n')
    pag.typewrite(f'{tabelaVendas}\n\n ')
    totalvendas = totalvendas + tabelaVendas['Vendas'].sum()
    time.sleep(1.5)

#Imprimir Total de vendas e enviar o email.
time.sleep(3)
pag.typewrite(f'\nTotal de vendas no primeiro semestre: \n R${totalvendas}')
time.sleep(2)
with pag.hold('ctrl'):
    pag.press('enter')
print('\n Automação concluida com sucesso.')


#BACKSPACE '\b'
#TAB       '\t'
#ENTER     '\n'
#RETURN    '\r'
#ESC       '\x1b'
#DELETE    '\x7f'

"""
#Lembrando que os tempos de espera podem variar de acordo com seu poder computacional.
# No caso de computadores mais lentos, aumentar valores no time.sleep()
# Em cada linha que o mesmo se encontra.
"""
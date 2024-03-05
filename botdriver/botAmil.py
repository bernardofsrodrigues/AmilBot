from selenium import webdriver
import time
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import pandas as pd
import openpyxl
import PySimpleGUI as sg
#import mysql.connector 
#import requests


df = usuario = senha = arquivo = conteudo = None

def executar_aplicativo():
    global df, usuario, senha, arquivo, conteudo

    layout = [
        [sg.Text('Usuário:'), sg.Input(key='usuario', size=(20,1))],
        [sg.Text('Senha:  '), sg.Input(key='senha',password_char='*', size=(20,1))],
        [sg.Text('Selecione um arquivo Excel:')],
        [sg.Input(key='-FILE-', enable_events=True), sg.FileBrowse(size=(6))],
        [sg.Button('Enviar', size=(6))],
    ]

    window = sg.Window('Carregar Arquivo Excel', layout)

    while True:
        event, values = window.read()

        if event == sg.WINDOW_CLOSED:
            break
        elif event == 'Enviar':
            usuario = values['usuario']
            senha = values['senha']
            arquivo = values['-FILE-']

            if arquivo:
                try:
                    df = pd.read_excel(arquivo)

                    print(f"\nUsuário: {usuario}")
                    print(f"Senha: {senha}\n")
                    print("Conteúdo do arquivo Excel:")
                    print(df, "\n")

                except Exception as e:
                    print(f"Erro ao carregar o arquivo: {e}")


if __name__ == '__main__':
    executar_aplicativo()
'''
def realizarConexao():
    try:
        dados = mysql.connector.connect (
            user='root',
            password='',
            host='127.0.0.1',
            database='retorno_contratos',
        )

        if dados.is_connected():
            print("Conexão bem-sucedida!")

    except mysql.connector.Error as err:
        print(f"Erro: {err}")

realizarConexao()
'''

#df = pd.read_excel('C:\\Users\\User\\OneDrive\\Área de Trabalho\\trikas\\botdriver\\basewalter.xlsx')

def realizarConsulta(usuario, senha, arquivo):#7z7iojn319.06493896622
    if arquivo:
        planilha = arquivo.replace("/", "\\")
        try:
            workbook = openpyxl.load_workbook(planilha)
            sheet = workbook.active
        except (TypeError, FileNotFoundError, openpyxl.utils.exceptions.InvalidFileException) as e:
            print(f"\nErro ao carregar a planilha: {e}\n")

        driver = webdriver.Edge()

        driver.get("https://portalcorretor.amil.com.br/portal/web/servicos/usuario/corretor/login")

        if usuario is not None:
            wait = WebDriverWait(driver, 120)
            usuario_input = wait.until(EC.presence_of_element_located((By.XPATH, '/html/body/div[1]/section/form/fieldset/dl/dd[1]/input')))
            usuario_input.send_keys(usuario) 

        if senha is not None:
            senha_input = wait.until(EC.presence_of_element_located((By.XPATH, '/html/body/div[1]/section/form/fieldset/dl/dd[2]/input')))
            senha_input.send_keys(senha) 

        login_button = wait.until(EC.element_to_be_clickable((By.XPATH, '/html/body/div[1]/section/form/div/button[2]')))
        login_button.click()

        menu_element = wait.until(EC.presence_of_element_located((By.XPATH, '/html/body/div[1]/nav/ul/li[2]/a')))
        menu_element.click()

        all_handles = driver.window_handles
        driver.switch_to.window(all_handles[1])

        element_to_click = wait.until(EC.element_to_be_clickable((By.XPATH, '/html/body/div[2]/div[3]/div')))
        element_to_click.click()

        time.sleep(3)

        '''
        driver.find_element(By.XPATH,'/html/body/div[2]/h1[4]/a').click()
        time.sleep(5)
        driver.find_element(By.XPATH,'/html/body/div[2]/div[4]/h2/a').click()
        time.sleep(5)
        driver.find_element(By.XPATH,'/html/body/div[2]/div[4]/div/h3/a').click()
        time.sleep(5)
        '''

        driver.switch_to.frame('menu')
        driver.find_element(By.XPATH,'/html/body/div[2]/h1[4]').click()
        time.sleep(1.5)
        driver.find_element(By.XPATH,'/html/body/div[2]/div[4]/h2/a').click()
        time.sleep(1.5)
        driver.find_element(By.XPATH,'/html/body/div[2]/div[4]/div/h3[1]').click()
        time.sleep(1.5)
        driver.switch_to.default_content()

        time.sleep(3)

        workbook = openpyxl.Workbook()
        sheet = workbook.active

        sheet.append(["Contrato", "Fatura", "Ciclo", "Referencia", "Vencimento", "Pagamento", "Valor", "Dias_atraso"])

        for index, row in df.iterrows():
            contrato = row['contrato'] 

            driver.switch_to.default_content()

            driver.switch_to.frame('principal')
            driver.find_element(By.XPATH,'/html/body/fieldset[2]/center/table/tbody/tr/td/table/tbody/tr/td/b/form[1]/table/tbody/tr[1]/td[2]/input[1]').clear()
            time.sleep(1)
            driver.find_element(By.XPATH, '/html/body/fieldset[2]/center/table/tbody/tr/td/table/tbody/tr/td/b/form[1]/table/tbody/tr[1]/td[2]/input[1]').send_keys(str(contrato))
            time.sleep(1)
            driver.find_element(By.XPATH,'/html/body/fieldset[2]/center/table/tbody/tr/td/table/tbody/tr/td/b/form[1]/table/tbody/tr[2]/td[2]/input[1]').send_keys('01/2020')
            time.sleep(1)
            driver.find_element(By.XPATH,'/html/body/fieldset[2]/center/table/tbody/tr/td/table/tbody/tr/td/b/form[1]/table/tbody/tr[2]/td[2]/input[2]').send_keys('06/2025')
            time.sleep(1)

            driver.switch_to.default_content()

            driver.switch_to.frame('toolbar')
            driver.find_element(By.XPATH,'/html/body/div/span[1]/img').click()
            time.sleep(2)

            driver.switch_to.default_content()

            driver.switch_to.frame('principal')
            time.sleep(5)

            cabeca = driver.find_element(By.XPATH,f'/html/body/table[2]/tbody/tr[1]')
            var1 = cabeca.text.split()
            print("\n\n")
            print("Contrato",var1)
            time.sleep(0.1)


            elements = driver.find_elements(By.XPATH, '/html/body/table[2]/tbody/tr')
            time.sleep(3)

            x = 2
            y = 2
            count = len(elements)
            for x in range(count):
                try:
                    fatura = driver.find_element(By.XPATH, f'/html/body/table[2]/tbody/tr[{y}]') 
                    retorno_fatura = fatura.text.split()
                    print(contrato, retorno_fatura)

                    sheet.append([contrato] + retorno_fatura)
                except:
                    None

                y += 1

            driver.switch_to.default_content()

            driver.switch_to.frame('toolbar')
            driver.find_element(By.XPATH,'/html/body/div/span/img').click()
            time.sleep(2)

            
        workbook.save('botdriver\\resultados.xlsx')

        driver.quit()
        time.sleep(5)

    else:
        print("\nArquivo não especificado. Certifique-se de que a variável 'arquivo' está definida corretamente.\n")

realizarConsulta(usuario, senha, arquivo)

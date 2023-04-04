import requests
import xlsxwriter  
from selenium import webdriver
import selenium.webdriver.support.expected_conditions as EC
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.wait import WebDriverWait
from time import sleep
from datetime import date
from selenium.webdriver.common.keys import Keys


options = Options()
options.add_experimental_option("detach", True)

navegador = webdriver.Chrome(ChromeDriverManager().install(), options=options)

navegador.get("https://sig.ufla.br/modulos/processos_seletivos_alunos/candidatos_alunos/acesso/chamadas.php")

sleep(5)

navegador.find_element(By.XPATH, '//*[@id="cod_processo_seletivo"]').click()

for i in range(5):
    navegador.find_element(By.XPATH, '//*[@id="cod_processo_seletivo"]').send_keys(Keys.ARROW_UP)

navegador.find_element(By.XPATH, '//*[@id="cod_processo_seletivo"]').send_keys(Keys.ENTER)

sleep(5)
navegador.find_element(By.XPATH, '//*[@id="enviar"]').click()

sleep(5)

tabelas = navegador.find_elements(By.CLASS_NAME, 'tabela')
print(len(tabelas))


curso=[]
aluno = []
modalidade=[]


for i in range(1, len(tabelas)-1):
    print(i)

    print('//*[@id="centro"]/table[' + str(i) + ']/caption')

    curso_aux =(navegador.find_element(By.XPATH, '//*[@id="centro"]/table[' + str(i) + ']/caption').text.split("-")[1].split(" na")[0])
    
    alunos = int(navegador.find_element(By.XPATH, '//*[@id="centro"]/table[' + str(i) + ']/tfoot/tr/td').text.split(" ")[1])

    for j in range(1, alunos+1):
        curso.append(curso_aux)
        aluno.append(navegador.find_element(By.XPATH, '//*[@id="centro"]/table[' + str(i) + ']/tbody/tr[' + str(j)+']/td[1]').text)
        modalidade.append(navegador.find_element(By.XPATH, '//*[@id="centro"]/table[' + str(i) + ']/tbody/tr[' + str(j)+']/td[2]').text)
        

print(curso)
print(aluno)
print(modalidade)

#Definir o caminho para salvar o arquivo .xlsx      
book = xlsxwriter.Workbook('')     
sheet = book.add_worksheet()  

#Cabeçalho do arquivo
sheet.write(0, 0, 'Aluno')
sheet.write(0, 1, 'Curso')
sheet.write(0, 2, 'Modalidade')


#Todas as listas possuem o mesmo número de elementos
for i in range(len(curso)):
    sheet.write(i+1, 0, aluno[i])
    sheet.write(i+1, 1, curso[i])
    sheet.write(i+1, 2, modalidade[i])
    
book.close()

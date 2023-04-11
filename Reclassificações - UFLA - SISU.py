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

#Processos seletivos da UFLA
navegador.get("https://sig.ufla.br/modulos/processos_seletivos_alunos/candidatos_alunos/acesso/chamadas.php")

sleep(5)

curso=[]
aluno = []
modalidade=[]
campus=[]
cliques=[]
chamada = []

navegador.find_element(By.XPATH, '//*[@id="cod_processo_seletivo"]').click()

opcoes = (navegador.find_element(By.XPATH, '//*[@id="cod_processo_seletivo"]').text).split("\n")
print(opcoes)


n_cliques = []

for j in range(len(opcoes)):
    
    if(opcoes[j].split(" ")[0] == "SISU" and opcoes[j].split(" ")[2] == "2023/1"):
        print(j)

        cliques.append(opcoes[j])

        print(len(opcoes)-j - 1)

        n_cliques.append(len(opcoes)-j - 1)

print("Numero de cliques\n")
for t in range(len(n_cliques)):
    print(n_cliques[t])

print("\n")

navegador.find_element(By.XPATH, '//*[@id="cod_processo_seletivo"]').click()


for chamadas in range (19):

    navegador.find_element(By.XPATH, '//*[@id="chamada"]').click() 
    navegador.find_element(By.XPATH, '//*[@id="chamada"]').send_keys(Keys.ARROW_DOWN)
    navegador.find_element(By.XPATH, '//*[@id="chamada"]').send_keys(Keys.ENTER)


    for a in range(len(n_cliques)):

        navegador.find_element(By.XPATH, '//*[@id="chamada"]').click()


        print("n de cliques")
        
        sleep(5)

        #navegador.find_element(By.XPATH, '//*[@id="cod_processo_seletivo"]').click()

        for b in range(n_cliques[a]):
            navegador.find_element(By.XPATH, '//*[@id="cod_processo_seletivo"]').send_keys(Keys.ARROW_UP)

        navegador.find_element(By.XPATH, '//*[@id="cod_processo_seletivo"]').send_keys(Keys.ENTER)

        sleep(5)

        navegador.find_element(By.XPATH, '//*[@id="enviar"]').click()

        sleep(5)

        navegador.find_element(By.XPATH, '//*[@id="enviar"]').click()
        
        sleep(10)

        tabelas = navegador.find_elements(By.CLASS_NAME, 'tabela')
        print(len(tabelas))

        for i in range(1, len(tabelas)):
            print(i)

            print('//*[@id="centro"]/table[' + str(i) + ']/caption')

            curso_aux =(navegador.find_element(By.XPATH, '//*[@id="centro"]/table[' + str(i) + ']/caption').text.split("-")[1].split(" na")[0])
            
            alunos = int(navegador.find_element(By.XPATH, '//*[@id="centro"]/table[' + str(i) + ']/tfoot/tr/td').text.split(" ")[1])

            for j in range(1, alunos+1):
                campus.append(cliques[a])
                curso.append(curso_aux)
                aluno.append(navegador.find_element(By.XPATH, '//*[@id="centro"]/table[' + str(i) + ']/tbody/tr[' + str(j)+']/td[1]').text)
                modalidade.append(navegador.find_element(By.XPATH, '//*[@id="centro"]/table[' + str(i) + ']/tbody/tr[' + str(j)+']/td[2]').text)
                chamada.append(chamadas+1)

        print(campus)
        print(curso)
        print(aluno)
        print(modalidade)
        print(chamada)

        for b in range(n_cliques[a]):
            navegador.find_element(By.XPATH, '//*[@id="cod_processo_seletivo"]').send_keys(Keys.ARROW_DOWN)
            navegador.find_element(By.XPATH, '//*[@id="cod_processo_seletivo"]').send_keys(Keys.ENTER)

book = xlsxwriter.Workbook('')     
sheet = book.add_worksheet()  

#Cabeçalho do arquivo
sheet.write(0, 0, 'Aluno')
sheet.write(0, 1, 'Curso')
sheet.write(0, 2, 'Modalidade')
sheet.write(0, 2, 'Campus')
sheet.write(0, 2, 'Chamada')

#Todas as listas possuem o mesmo número de elementos
for i in range(len(curso)):
    sheet.write(i+1, 0, aluno[i])
    sheet.write(i+1, 1, curso[i])
    sheet.write(i+1, 2, modalidade[i])
    sheet.write(i+1, 3, campus[i])
    sheet.write(i+1, 4, chamada[i])
    
book.close()
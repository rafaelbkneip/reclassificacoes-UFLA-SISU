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

#Garantir o carregamento
sleep(5)

#Definir listas
curso = []
aluno = []
modalidade = []
campus = []
cliques = []
chamada = []
n_cliques = []

#Abrir a drop-down com as opções de processos seletivos
navegador.find_element(By.XPATH, '//*[@id="cod_processo_seletivo"]').click()

#Extrair o texto da drop-down
opcoes = (navegador.find_element(By.XPATH, '//*[@id="cod_processo_seletivo"]').text).split("\n")
print(opcoes)

#Varrer as opções, conferindo quais são referentes ao SISU de 2023 e guardar informações do seu texto e número de cliques necessário para acessá-la
for j in range(len(opcoes)):
    if(opcoes[j].split(" ")[0] == "SISU" and opcoes[j].split(" ")[2] == "2023/1"):
        #Guarda o texto da opção em questão
        cliques.append(opcoes[j])
        #Para uma dada opção do drop-down de opções, sua possição, de baixo para cima na lista, a partir da primeira posição, é dada pelo 
        #tamanho de opções menos sua posição na lista menos 1
        n_cliques.append(len(opcoes)- j - 1)

#Abrir novamente a caixa de opções
navegador.find_element(By.XPATH, '//*[@id="cod_processo_seletivo"]').click()

#Conferir as primeiras 20 classificações da UFLA para o SISU 
for chamadas in range (21):

    #A cada iteração, escolher a próxima reclassificação, que na lista drop-down de chamada estará imediatamente na linha inferior
    navegador.find_element(By.XPATH, '//*[@id="chamada"]').click() 
    navegador.find_element(By.XPATH, '//*[@id="chamada"]').send_keys(Keys.ARROW_DOWN)
    navegador.find_element(By.XPATH, '//*[@id="chamada"]').send_keys(Keys.ENTER)

    #Acessar as opções de processos seletivos do SISU 2023
    for a in range(len(n_cliques)):
        navegador.find_element(By.XPATH, '//*[@id="chamada"]').click()

        sleep(5)

        #Selecionar a referida opção do SISU, clicando o número de vezes necessário para tal
        for b in range(n_cliques[a]):
            navegador.find_element(By.XPATH, '//*[@id="cod_processo_seletivo"]').send_keys(Keys.ARROW_UP)
        navegador.find_element(By.XPATH, '//*[@id="cod_processo_seletivo"]').send_keys(Keys.ENTER)

        #Garantir o carregamento
        sleep(5)
        #Enviar opções e extrair relatório
        navegador.find_element(By.XPATH, '//*[@id="enviar"]').click()

        #Garantir o carregamento
        sleep(5)
        navegador.find_element(By.XPATH, '//*[@id="enviar"]').click()
        
        #Garantir o carregamento
        sleep(10)

        #Os aprovados estão alocados na página em elementos de classe 'tabela'
        tabelas = navegador.find_elements(By.CLASS_NAME, 'tabela')

        #Para cada uma das tabelas com notas dos alunos:
        for i in range(1, len(tabelas)):
            #Salvar o nome curso
            curso_aux =(navegador.find_element(By.XPATH, '//*[@id="centro"]/table[' + str(i) + ']/caption').text.split("-")[1].split(" na")[0])
            #Salvar a quantidade de alunos aprovados naquele curso, através da frase "Total: X candidatos (para o curso/chamada indicado)"
            alunos = int(navegador.find_element(By.XPATH, '//*[@id="centro"]/table[' + str(i) + ']/tfoot/tr/td').text.split(" ")[1])

            #Para cada um dos alunos:
            for j in range(1, alunos+1):
                #Salvar campus
                campus.append(cliques[a])
                #Salvar curso
                curso.append(curso_aux)
                #Salvar nome
                aluno.append(navegador.find_element(By.XPATH, '//*[@id="centro"]/table[' + str(i) + ']/tbody/tr[' + str(j)+']/td[1]').text)
                #Salvar modalidade
                modalidade.append(navegador.find_element(By.XPATH, '//*[@id="centro"]/table[' + str(i) + ']/tbody/tr[' + str(j)+']/td[2]').text)
                #Salvar a chamada em questão
                chamada.append(chamadas+1)

        #Retornar a posição original do quadro de opções, permitindo que a próxima opção do SISU seja selecionada
        for b in range(n_cliques[a]):
            navegador.find_element(By.XPATH, '//*[@id="cod_processo_seletivo"]').send_keys(Keys.ARROW_DOWN)
            navegador.find_element(By.XPATH, '//*[@id="cod_processo_seletivo"]').send_keys(Keys.ENTER)

#Abrir um arquivo .xlsx a partir do caminho
book = xlsxwriter.Workbook('Caminho')     
sheet = book.add_worksheet()  

#Cabeçalho do arquivo
sheet.write(0, 0, 'Aluno')
sheet.write(0, 1, 'Curso')
sheet.write(0, 2, 'Modalidade')
sheet.write(0, 3, 'Campus')
sheet.write(0, 4, 'Chamada')

#Todas as listas possuem o mesmo número de elementos
for i in range(len(curso)):
    sheet.write(i+1, 0, aluno[i])
    sheet.write(i+1, 1, curso[i])
    sheet.write(i+1, 2, modalidade[i])
    sheet.write(i+1, 3, campus[i])
    sheet.write(i+1, 4, chamada[i])

#Fechar o arquivo   
book.close()
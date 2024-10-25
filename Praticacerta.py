from selenium import webdriver
from selenium.webdriver.common.by import By
import time 
import pandas as pd
import win32com.client as win32
import datetime




#Abrindo o navegador
driver = webdriver.Chrome()
driver.get("https://www.bbc.com/portuguese")
time.sleep(5)

#Pegando os títulos da matérias
nome_materia= driver.find_element(By.XPATH,"/html/body/div[2]/div/div/main/div/div/section[1]/div/ul/li[1]/div/div[2]/h3/a").text

nome_materia2= driver.find_element(By.XPATH,"/html/body/div[2]/div/div/main/div/div/section[3]/div/ul/li[1]/div/div[2]/h3/a").text

nome_materia3= driver.find_element(By.XPATH,"/html/body/div[2]/div/div/main/div/div/section[1]/div/ul/li[3]/div/div[2]/h3/a").text

nome_materia4= driver.find_element(By.XPATH,"/html/body/div[2]/div/div/main/div/div/section[1]/div/ul/li[4]/div/div[2]/h3/a").text

nome_materia5= driver.find_element(By.XPATH,"/html/body/div[2]/div/div/main/div/div/section[3]/div/ul/li[2]/div/div[2]/h3/a").text

#Criando uma tabela com os títulos das matérias
Tabela1 = {
    'Nome Materia': [nome_materia, nome_materia2,nome_materia3,nome_materia4,nome_materia5],}

#Verificando se tem a palavra economia
#economia_total = Tabela1.stack().str.contains('economia', case=False).sum()
economia_total= 0

#Criando uma nova tabela com a contagem de notícias economias
Tabela = {
    'Nome Materia': [nome_materia, nome_materia2, nome_materia3, nome_materia4, nome_materia5],
    'Noticias economia': [economia_total,economia_total,economia_total,economia_total,economia_total]
}


data= datetime.datetime.now()
#Ciando um Dataframe a partir daquela tabela
df = pd.DataFrame(Tabela)
nome_arquivo= (f'NoticiasBBC_{data}.xlsx')
#Criando uma planilha no excel
tabelaCerta= Tabela.to_excel(nome_arquivo)
df.to_excel('C:\Users\raikasilva-ieg\OneDrive - Instituto Germinare\Área de Trabalho\Tech', index = False)
print('Funcionando!')


#Enviar e-mail 
outlook = win32.Dispatch("outlook.application")
email_out = outlook.CreateItem(0)
 
 
email_out.To= "raika.silva@germinare.org.br"
 
email_out.Subject = "Destaques da BBC Brasil: Análise Diária"
variavel = (f'Notícias do dia: {nome_materia,nome_materia2,nome_materia3,nome_materia4,nome_materia5} A quantidade total é: {economia_total} {tabelaCerta}')
 
email_out.htmlBody= variavel

email_out.Send()
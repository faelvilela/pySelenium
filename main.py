from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.common.by import By
import time, shutil, os
from datetime import datetime, date, timedelta

menos15 = (date.today() - timedelta(days=15)).strftime('%Y-%m-%d')
data = date.today().strftime('%d_%m_%Y')
data2 = date.today().strftime('%m_%d_%Y').replace("_0", "_")
nome = 'planilha_excel_'+data+'.xlsx'
path = 'C:/Users/Administrator/Downloads/'
target_dir = 'Q:/TI/'

options = Options()
# Eliminar erro Bluetooth adapter
options.add_experimental_option("excludeSwitches", ["enable-logging"])
options.add_experimental_option("detach", True)
# Para abrir o navegador com um perfil especifico
# options.add_argument('--profile-directory=Profile 1')
# options.add_argument("user-data-dir=C:\\Users\\Administrator\\AppData\\Local\\Google\\Chrome\\User Data\\")

navegador = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=options)  # Setando navegador

navegador.get('https://site.com.br/#/signIn/')
login = navegador.find_element(By.XPATH, '//*[@id="conteudo-interno"]/div[1]/div/form/div[1]/div/input')
login.send_keys('login')
senha = navegador.find_element(By.XPATH, '//*[@id="conteudo-interno"]/div[1]/div/form/div[2]/div/input')
senha.send_keys('senha')
navegador.find_element(By.XPATH, '//*[@id="conteudo-interno"]/div[1]/div/form/div[4]/button').click()

time.sleep(5)
filtrar = navegador.find_element(By.XPATH, '//*[@id="conteudo-interno"]/div/div[1]/div[2]/div/div[2]/button[1]').click()
data = navegador.find_element(By.XPATH, '//*[@id="collpseFilter"]/div/div/form/div[3]/div/div/input[1]')
navegador.execute_script("arguments[0].value = '"+menos15+"';", data)
filtrar2 = navegador.find_element(By.XPATH, '//*[@id="collpseFilter"]/div/div/form/div[6]/button[2]').click()
time.sleep(15)
exportar = navegador.find_element(By.XPATH, '//*[@id="conteudo-interno"]/div/div[1]/div[2]/div/div[2]/button[2]').click()
time.sleep(60)
navegador.quit()

#Movendo arquivos Baixados
file_names = os.listdir(path)
for file_name in file_names:
    if file_name.endswith('.xlsx'):
        shutil.move(os.path.join(path, file_name), target_dir)
os.startfile('Macros.xlsm')




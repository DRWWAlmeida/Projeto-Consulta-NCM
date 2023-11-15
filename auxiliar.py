import pandas as pd
import openpyxl
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
import time

planilha = pd.read_excel('PRODUTOS 10-2023 - teste - Copia.xls', dtype=str)
#print(planilha)

driver = webdriver.Chrome()
driver.get("https://www.econeteditora.com.br//pis_cofins/index.php?form[a]=0")


user_input = driver.find_element(By.XPATH,"//input[@name='Log']")
user_input.send_keys('gdc47781')
time.sleep(3)
password_input = driver.find_element(By.XPATH,"//input[@name='Sen']")
password_input.send_keys('X32023')
time.sleep(10)


workbook = openpyxl.Workbook()
workbook.create_sheet("consulta")
sheet_consulta = workbook['consulta']

for linha in planilha.index:
    booking_date = (planilha.loc[linha, "Dt. Escrituação"]) 
    number = (planilha.loc[linha, "Número"])
    model = (planilha.loc[linha, "Mod."])
    product_name = (planilha.loc[linha, "Nome Produto"])
    ncm = (planilha.loc[linha, "Classificação"])
    total_value = (planilha.loc[linha, "Valor Total Item"])
    discount = (planilha.loc[linha, "Descontos"])
    #ajusted_value = total_value - discount

    driver.switch_to.frame("ifram")
    ncm_input = driver.find_element(By.XPATH, "//input[@name='form[ncm]']")
    ncm_input.clear()
    time.sleep(1)
    ncm_input.send_keys(ncm)
    input()
import openpyxl
from selenium import webdriver
from selenium.webdriver.common.by import By
from time import sleep

# Entra na planilha e exatrai os dados
planilhaClientes = openpyxl.load_workbook("dados_clientes.xlsx") # Abre a planilha
paginaClientes = planilhaClientes["Sheet1"]

for linha in paginaClientes.iter_rows(min_row = 2, values_only = True):
    nome, valor, cpf, vencimento = linha
    driver = webdriver.Chrome()
    driver.get("https://consultcpf-devaprender.netlify.app/")
    sleep(5) # Atrasa o carregamento do site em 5 segundos
    campoPesquisa = driver.find_element(By.XPATH, "//input[@id='cpfInput']")
    sleep(1) # Pausa para o campo ser carregado
    campoPesquisa.send_keys(cpf)
    sleep(1) # Pausa para o campo ser preenchido
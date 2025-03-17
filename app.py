import openpyxl
from selenium import webdriver
from selenium.webdriver.common.by import By
from time import sleep

# Entra na planilha e exatrai os dados
planilhaClientes = openpyxl.load_workbook("dados_clientes.xlsx") # Abre a planilha
paginaClientes = planilhaClientes["Sheet1"]

driver = webdriver.Chrome()
driver.get("https://consultcpf-devaprender.netlify.app/")

for linha in paginaClientes.iter_rows(min_row = 2, values_only = True):
    
    print(linha)
    
    nome, valor, cpf, vencimento = linha
    sleep(5) # Atrasa o carregamento do site em 5 segundos

    campoPesquisa = driver.find_element(By.XPATH, "//input[@id='cpfInput']")
    sleep(1) # Pausa para o campo ser carregado
    campoPesquisa.clear()
    
    campoPesquisa.send_keys(cpf)
    sleep(1) # Pausa para o campo ser preenchido

    botaoConsultar = driver.find_element(By.XPATH, "//button[@class='btn btn-custom btn-lg btn-block mt-3']")
    sleep(1) # Pausa para o botão ser carregado

    botaoConsultar.click()
    sleep(4) # Pausa para a consulta ser feita

    campoStatus = driver.find_element(By.XPATH, "//span[@id='statusLabel']")
    
    if campoStatus.text == "em dia":
        dataPagamento = driver.find_element(By.XPATH, "//p[@id='paymentDate']")
        metodoPagamento = driver.find_element(By.XPATH, "//p[@id='paymentMethod']")
        
        dataPagamentoData = dataPagamento.text.split()[3]
        metodoPagamentoMetodo = metodoPagamento.text.split()[3]
        
        planilhaFechamento = openpyxl.load_workbook("dados_clientes_fechamento.xlsx") # Abre a planilha
        paginaFechamento = planilhaFechamento["Sheet1"]

        paginaFechamento.append([nome, valor, cpf, vencimento, "em dia", dataPagamentoData, metodoPagamentoMetodo])
        
        planilhaFechamento.save("dados_clientes_fechamento.xlsx")
    else:
        planilhaFechamento = openpyxl.load_workbook("dados_clientes_fechamento.xlsx") # Abre a planilha
        paginaFechamento = planilhaFechamento["Sheet1"]
        
        paginaFechamento.append([nome, valor, cpf, vencimento, "pendente"])
        
        planilhaFechamento.save("dados_clientes_fechamento.xlsx")

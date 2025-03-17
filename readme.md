### Automação de Consulta de Pagamentos e Registro em Excel  

Este projeto é um **script Python** que utiliza **Selenium** para acessar um site de consulta de CPF e verificar o status de pagamento de clientes. Os dados coletados são armazenados automaticamente em uma **planilha Excel**, facilitando a gestão financeira.  

---

## 📌 **Funcionamento do Aplicativo**  

1. **Leitura da planilha de clientes:**  
   - O script abre um arquivo `dados_clientes.xlsx` contendo informações como **nome, valor devido, CPF e data de vencimento**.  

2. **Automação da consulta:**  
   - O Selenium inicia um navegador **Chrome** e acessa o site de consulta de CPF.  
   - Para cada cliente, o CPF é inserido no campo de pesquisa, e o botão de consulta é acionado.  

3. **Captura do status de pagamento:**  
   - O script verifica se o status do CPF é **"em dia"** ou **"pendente"**.  
   - Se o pagamento estiver **em dia**, o **método e a data do pagamento** também são extraídos.  

4. **Registro no Excel:**  
   - Os dados são adicionados automaticamente na planilha `dados_clientes_fechamento.xlsx`.  
   - Clientes em dia são registrados com a **data e o método de pagamento**.  
   - Clientes inadimplentes recebem a tag **"pendente"**.  

---

## 🔧 **Tecnologias Utilizadas**  

✅ **Python** – Linguagem principal do script  
✅ **Selenium** – Automação de navegador para interagir com o site  
✅ **OpenPyXL** – Manipulação de arquivos Excel  
✅ **Time (sleep)** – Controle de tempo para garantir carregamento correto da página  

---

## 📌 **Vantagens do Sistema**  

🚀 **Automação total** – Reduz trabalho manual e erros humanos  
📊 **Registro estruturado** – Organiza dados financeiros em Excel  
🔍 **Rastreamento de pagamentos** – Identifica clientes adimplentes e inadimplentes automaticamente  

Esse script pode ser facilmente adaptado para diferentes **sites de consulta** e personalizações conforme a necessidade da empresa.
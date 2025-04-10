#Ler dados da planilha
# Inserir cada célula de cada linha em uma lista em um campo do sistema 
import openpyxl
import pyautogui
 

workbookk = openpyxl.load_workbook('vendas_de_produtos.xlsx')
vendas_sheet = workbookk['vendas']

for linha in vendas_sheet.iter_rows(min_row=2):    
    #...nome                              
    pyautogui.click(1808, 452, duration=1.5)
    pyautogui.write(linha[0].value) # Nome do vendedor
    #...produto
    pyautogui.click(1815,476, duration=1.5)
    pyautogui.write(linha[1].value) # Nome do produto
    #...quantidade
    pyautogui.click(1813, 497, duration=1.5)    
    pyautogui.write(str(linha[2].value)) # Preço do produto
    #...categoria
    pyautogui.click(1883, 532, duration=1.5)    
    pyautogui.write(linha)
    pyautogui.click(1752,549, duration=1.5) # Clicar no botão de salvar
    pyautogui.click(1256,581, duration=1.5) # Clicar no botão de salvar
    print(linha[0].value) # Nome do vendedor
    print(linha[1].value) # Nome do produto 
    print(linha[2].value) # Preço do produto
    print(linha[3].value) # Data da venda
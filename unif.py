import tkinter as tk
from tkinter import ttk
import openpyxl
import pyautogui
import threading
import time

# Função para preencher automaticamente usando dados do Excel
def preencher_formulario():
    time.sleep(3)  # Tempo pra você colocar a janela do formulário em primeiro plano

    # Abre a planilha
    workbook = openpyxl.load_workbook('vendas_de_produtos.xlsx')
    sheet = workbook['vendas']

    for linha in sheet.iter_rows(min_row=2):
        cliente = linha[0].value
        produto = linha[1].value
        quantidade = str(linha[2].value)
        categoria = linha[3].value

        # Clique e preenchimento dos campos
        pyautogui.click(1808, 452, duration=0.5)
        pyautogui.write(cliente or '')

        pyautogui.click(1815, 476, duration=0.5)
        pyautogui.write(produto or '')

        pyautogui.click(1813, 497, duration=0.5)
        pyautogui.write(quantidade or '')

        pyautogui.click(1883, 532, duration=0.5)
        pyautogui.write(categoria or '')

        pyautogui.click(1752, 549, duration=0.5)  # Botão Salvar

        print(f"Preenchido: {cliente}, {produto}, {quantidade}, {categoria}")
        time.sleep(1)

# Função do botão "Iniciar Preenchimento"
def iniciar_preenchimento():
    thread = threading.Thread(target=preencher_formulario)
    thread.start()

# GUI
janela = tk.Tk()
janela.title("Cadastro de Produto")
janela.geometry("300x200")

tk.Label(janela, text="Cliente").grid(row=0, column=0, padx=10, pady=5, sticky="w")
entry_cliente = tk.Entry(janela)
entry_cliente.grid(row=0, column=1)

tk.Label(janela, text="Produto").grid(row=1, column=0, padx=10, pady=5, sticky="w")
entry_produto = tk.Entry(janela)
entry_produto.grid(row=1, column=1)

tk.Label(janela, text="Quantidade").grid(row=2, column=0, padx=10, pady=5, sticky="w")
entry_quantidade = tk.Entry(janela, width=5)
entry_quantidade.grid(row=2, column=1, sticky="w")

tk.Label(janela, text="Categoria do Produto").grid(row=3, column=0, padx=10, pady=5, sticky="w")
combo_categoria = ttk.Combobox(janela, values=["Alimentos", "Bebidas", "Limpeza", "Outros"])
combo_categoria.grid(row=3, column=1)

btn_iniciar = tk.Button(janela, text="Iniciar Preenchimento", bg="green", fg="white", command=iniciar_preenchimento)
btn_iniciar.grid(row=4, column=0, columnspan=2, pady=10)

janela.mainloop()
git status

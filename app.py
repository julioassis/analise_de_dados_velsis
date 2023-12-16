import tkinter as tk
from tkinter import filedialog
from bs4 import BeautifulSoup
import pandas as pd
import ttkbootstrap as tb

def extrair_dados():
    # Obter o HTML do campo de entrada
    html = txt_html.get("1.0", "end-1c")

    # Extrair os dados das tabelas HTML
    soup = BeautifulSoup(html, "html.parser")
    tables = soup.find_all("table")

    # Criar um dicionário para armazenar os dados de todas as tabelas
    table_data = {}

    # Percorrer todas as tabelas encontradas na página
    for i, table in enumerate(tables):
        table_data["Tabela {}".format(i+1)] = {}

        # Criar listas para armazenar os dados de cada coluna da tabela atual
        columns = []
        for th in table.find("thead").find_all("th"):
            columns.append(th.text.strip())

        # Percorrer as linhas da tabela atual e extrair os dados
        rows = []
        for row in table.find("tbody").find_all("tr"):
            cells = row.find_all("td")
            row_data = [cell.text.strip() for cell in cells]
            rows.append(row_data)

        # Armazenar os dados da tabela atual no dicionário
        for j, column in enumerate(columns):
            table_data["Tabela {}".format(i+1)][column] = [row[j] for row in rows]

    # Salvar os dados em um arquivo Excel
    file_path = filedialog.asksaveasfilename(defaultextension=".xlsx")
    with pd.ExcelWriter(file_path) as writer:
        # Percorrer o dicionário de dados das tabelas
        for table_name, data in table_data.items():
            # Criar um DataFrame a partir dos dados da tabela atual
            df_table = pd.DataFrame(data)
            # Salvar o DataFrame como uma planilha no arquivo Excel
            df_table.to_excel(writer, sheet_name=table_name, index=False)

    # Exibir uma mensagem de conclusão
    lbl_message.config(text="Extração concluída. O arquivo Excel foi salvo com sucesso.")

# Criar a janela principal
root = tb.Window(themename='morph')
root.title("Extração de Dados de HTML")
root.geometry('850x380')
root.resizable(False, False)

# Criar um rótulo para o campo de entrada de HTML
lbl_html = tk.Label(root, text="Insira o código HTML:")
lbl_html.pack()

# Criar o campo de entrada de HTML
txt_html = tk.Text(root, height=10, width=80)
txt_html.pack()

# Criar o botão para extrair os dados
btn_extract = tk.Button(root, text="Extrair Dados", command=extrair_dados)
btn_extract.pack(pady=20)

# Criar um rótulo para exibir a mensagem de conclusão
lbl_message = tk.Label(root, text="")
lbl_message.pack()

# Iniciar o loop de eventos
root.mainloop()
""" Conversor arquivos Excel XLS para CSV """

# Bibliotecas utilizadas.

import os
from tkinter import Tk, filedialog, Label, Button, messagebox
import pandas as pd


class ExcelToCsvConverter:
    """ Telas """

    def __init__(self, root):
        self.root = root
        self.root.title("Conversor XLSX/CSV")  # Label/Titulo do seu app.
        self.root.geometry("550x150")  # Posição da tela do APP no monitor.

        self.label = Label(
            root, text="Selecione a pasta contendo os arquivos em Excel:")  # Texto solicitando a pasta de origem.
        self.label.pack(pady=10)

        self.button = Button(root, text="Selecionar",
                             command=self.select_folder)  # Texto botão de seleção.
        self.button.pack(pady=10)

        self.button = Button(root, text="Converter para CSV",
                             command=self.convert_excel_to_csv)  # Texto do botão de ação.
        self.button.pack(pady=10)

        self.status_label = Label(root, text="")
        self.status_label.pack(pady=10)

    def select_folder(self):
        """Seleção de Pasta"""

        self.folder_path = filedialog.askdirectory()
        self.label.config(text=f"Diretório: {self.folder_path}")  # Label "Diretório".

    def convert_excel_to_csv(self):
        """Conversão de Excel para CSV"""

        excel_files = [f for f in os.listdir(
            self.folder_path) if f.endswith('.xlsx')]

        if not excel_files:
            messagebox.showinfo(
                "Erro", "Nenhum arquivo Excel encontrado na pasta.")  # Texto da mensagem de erro caso não haja arquivos xlsx na pasta.
            return

        csv_directory = os.path.join(self.folder_path, "CSV")  # Texto do nome da pasta de destino que será crtiada onde ficarão os novos arquivos convertidos.
        if not os.path.exists(csv_directory):
            os.makedirs(csv_directory)

        for excel_file in excel_files:
            excel_path = os.path.join(self.folder_path, excel_file)
            df = pd.read_excel(excel_path)

            csv_file = os.path.splitext(excel_file)[0] + '.csv'
            csv_path = os.path.join(csv_directory, csv_file)

            df.drop(0, axis=0, inplace=True)
            df.to_csv(csv_path, index=False, header=False,  # Parâmetros de conversão do arquivo, removendo "Cabeçalho", "Linha em branco", codificação e separação por ";" caso queira é só subistiuir por ","
                      encoding="ISO-8859-1", sep=";")

        messagebox.showinfo(
            "Finalizado", "Conversão concluída com sucesso!")  # Popup de Mensagem de conclusão.

        print("Conversão concluída:")  # Mensagem de conclusão no terminal.


if __name__ == "__main__":
    root = Tk()
    # Obtém a largura e a altura da tela
    screen_width = root.winfo_screenwidth()
    screen_height = root.winfo_screenheight()

    # Obtém a largura e a altura da janela
    window_width = 550  # substitua pelo valor desejado
    window_height = 200  # substitua pelo valor desejado

    # Calcula as coordenadas para centralizar a janela
    x = (screen_width - window_width) // 2
    y = (screen_height - window_height) // 2

    # Define as dimensões e a posição da janela
    root.geometry(f"{window_width}x{window_height}+{x}+{y}")

    app = ExcelToCsvConverter(root)
    root.mainloop()

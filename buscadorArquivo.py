import tkinter as tk
from tkinter import filedialog

def buscadorDeArquivo():

    root = tk.Tk() # cria uma janela
    root.withdraw() # esconde a janela

    # abre o explorador de arquivos e escolhe o arquivo txt
    file_path = filedialog.askopenfilename() 

    root.destroy() # fecha a janela
    return file_path

if __name__ == "__main__":
    file_path = buscadorDeArquivo()
    print(file_path)
import tkinter as tk
from tkinter import filedialog

def buscadorArquivo():

    root = tk.Tk() # cria uma janela
    root.withdraw() # esconde a janela

    # abre o explorador de arquivos e escolhe o arquivo txt
    file_path = filedialog.askopenfilename() 

    root.destroy() # fecha a janela
    return file_path
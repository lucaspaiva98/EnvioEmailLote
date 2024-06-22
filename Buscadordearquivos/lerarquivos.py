from Buscadordearquivos.buscadorArquivo import buscadorArquivo
import pandas as pd  # Importa a biblioteca pandas
import os  # Importa a biblioteca os

def executarleitura():
    # Abrir o arquivo
    file_path = buscadorArquivo()
    # file_path = os.path.join(os.getcwd(), 'Dados p. preencher.xlsx')
    data = pd.read_excel(file_path, dtype=str)  # leitura do arquivo excel em formato texto

    return data
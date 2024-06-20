from buscadorArquivo import buscadorDeArquivo
import pandas as pd
import os

def abrirArquivo(file_path):

    with open(file_path, "r") as file:
        data = file.readlines()

    data = [line.strip() for line in data]

    # mostrar a lista linha por linha
    for line in data:
        print(line)
    return data

def __init__(self):
    def executarleitura():
        # Abrir o arquivo
        # file_path = buscadorDeArquivo()
        file_path = os.path.join(os.getcwd(), 'Dados p. preencher.xlsx')
        data = pd.read_excel(file_path, dtype=str)  # tudo como texto

        # Dados do usuário
        self.nome = data['Nome_Funcionario'].tolist()
        self.setor = data['Setor'].tolist()
        self.funcao = data['Funcao'].tolist()
        self.celular = data['Celular Institucional'].tolist()
        self.telefone = data['Contato_Empresarial'].tolist()
        self.ramal = data['Ramal'].tolist()
        self.email = data['Email'].tolist()

if __name__ == "__main__":
    nome, setor, funcao, celular, telefone, ramal, email = executarleitura()
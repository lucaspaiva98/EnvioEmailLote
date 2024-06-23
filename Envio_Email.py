# Envio de Email em Lote usando outlook
# Outlook object com funçao

import win32com.client as win32
from email.mime.image import MIMEImage
import os
from Buscadordearquivos.lerarquivos import executarleitura
import tkinter as tk
from tkinter import filedialog
import pandas as pd
i = 0

def coleta_dados():
    data = executarleitura()

    # Dados do usuário para lista
    nome = data['Nome_Funcionario'].fillna('-').tolist()
    setor = data['Setor'].fillna('-').tolist()
    funcao = data['Funcao'].fillna('-').tolist()
    celular = data['Celular Institucional'].fillna('-').tolist()
    telefone = data['Contato_Empresarial'].fillna('-').tolist()
    ramal = data['Ramal'].fillna('-').tolist()
    email = data['Email'].fillna('-').tolist()

    return nome, setor, funcao, celular, telefone, ramal, email

def variaveistxt():
    nome, setor, funcao, celular, telefone, ramal, email = coleta_dados()

    txtPrincipal = '''<p style='font-family:calibri;'>Caro colaborador(a),<br><br>
    Com o objetivo de padronizar as assinaturas de e-mail dos funcionários do Grupo Camed,
    <b>estamos enviando sua assinatura ajustada em anexo, juntamente com o manual contendo o passo a passo a ser seguido</b> para a atualização.<br>
    Será necessário <b>SALVAR a imagem enviada em anexo em seu computador</b>
    e <b>ACESSAR o manual referente à ferramenta de e-mail utilizada: Outlook, Thunderbird ou Webmail (OWA).</b><br></p>'''

    dados = '''<p style='font-family:calibri;'>Seguem as informações da assinatura em anexo:<br><br>
    <u>Nome</u>: {}<br>
    <u>Setor</u>: {}<br>
    <u>Função</u>: {}<br>
    <u>Celular</u>: {}<br>
    <u>Telefone</u>: {}<br>
    <u>Ramal</u>: {}<br>
    <u>E-mail</u>: {}<br></p>'''.format(nome[i], setor[i], funcao[i], celular[i], telefone[i], ramal[i], email[i])

    texto3 = '''<p style='font-family:calibri;'>Em caso de dúvidas na execução do procedimento ou em caso de mudança na assinatura enviada,
    retornar este e-mail clicando em <b>Responder a Todos</b>.
    <br><br>Desde já, agradecemos sua colaboração!<br></p>'''

    return txtPrincipal, dados, texto3, nome, email

def Remetente(ass_remetente):
    # Caminho para a imagem
    image_path = os.path.join(os.getcwd(), ass_remetente)

    # Abrindo a imagem em modo binário
    with open(image_path, 'rb') as f:
        img_Ass = f.read()
    
    # Cria um objeto MIMEImage
    img = MIMEImage(img_Ass)

    # Adiciona um 'Content-ID' para a imagem
    img.add_header('Content-ID', '<{}>'.format(ass_remetente))

    # Gera a string HTML para a assinatura
    ass_html = '''<span style='font-family:calibri;'<br>Atenciosamente <span><br>
    <img src="cid:{}">'''.format(ass_remetente)

    return img, ass_html

def envia_email():
    txtPrincipal, dados, texto3, nome, email = variaveistxt()

    assRemetenteEnvio = 'AssLucasP.png'
    img, assLucasP = Remetente(assRemetenteEnvio)

    outlook = win32.Dispatch('outlook.application') # Abrir o Outlook
    mail = outlook.CreateItem(0)  # Criar um novo email

    mail.Display()  # Exibir o email
    
    mail.To = email[i] # Destinatário
    mail.CC = '' # Com Cópia
    
    mail.Subject = 'Assinatura de E-mail - {}'.format(nome[i])  # Assunto

    # Corpo do Email
    mail.HTMLBody = txtPrincipal + dados + texto3 + assLucasP

    # Anexos usar imagem assLucasP da pasta do projeto
    assRemetente = os.getcwd() + '\\AssLucasP.png'
    mail.Attachments.Add(assRemetente)

    #mail.Send() # Enviar Email

def exibir_dados():
    nome, setor, funcao, celular, telefone, ramal, email = coleta_dados()

    # Imprimir dados do usuário
    print("{}, {}".format(nome[0], nome[1]))
    print("{}, {}".format(setor[0], setor[1]))
    print("{}, {}".format(funcao[0], funcao[1]))
    print("{}, {}".format(celular[0], celular[1]))
    print("{}, {}".format(telefone[0], telefone[1]))
    print("{}, {}".format(ramal[0], ramal[1]))
    print("{}, {}".format(email[0], email[1]))

if __name__ == '__main__':
    envia_email()
    # exibir_dados()

# print('Emails Enviados com Sucesso')
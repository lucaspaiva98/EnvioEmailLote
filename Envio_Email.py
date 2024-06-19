# Envio de Email em Lote usando outlook
# Outlook object com funlçao

import win32com.client as win32
from email.mime.image import MIMEImage
import os
import time

def path_AssRemetente(Ass_Remetente):
    # Caminho para a imagem
    image_path = os.path.join(os.getcwd(), Ass_Remetente)

    # Abrindo a imagem em modo binário
    with open(image_path, 'rb') as f:
        img_Ass = f.read()
    
    # Cria um objeto MIMEImage
    img = MIMEImage(img_Ass)

    # Adiciona um 'Content-ID' para a imagem
    img.add_header('Content-ID', '<{}>'.format(Ass_Remetente))

    return img

def variaveis():

    assremetenteEnvio = 'AssLucasP.png'

    assLucasP = '''<span style='font-family:calibri;'><br>Atenciosamente,<span><br>
    <img src="file:///E:/Bkp/pc/Programas camed/Lucas P/MPython/bkp/bk1/EnvioEmailLote/{}">'''.format(assremetenteEnvio)

    txtPrincipal = '''<p style='font-family:calibri;'>Caro colaborador(a),<br><br>
    Com o objetivo de padronizar as assinaturas de e-mail dos funcionários do Grupo Camed,
    <b>estamos enviando sua assinatura ajustada em anexo, juntamente com o manual contendo o passo a passo a ser seguido</b> para a atualização.<br>
    Será necessário <b>SALVAR a imagem enviada em anexo em seu computador</b>
    e <b>ACESSAR o manual referente à ferramenta de e-mail utilizada: Outlook, Thunderbird ou Webmail (OWA).</b><br></p>'''

    txtPrincipal2 = '''<p style='font-family:calibri;'>Caro colaborador(a),<br><br>
    <b>Estamos enviando sua assinatura ajustada em anexo, juntamente com o manual contendo o passo a passo a ser seguido</b> para a atualização.<br>
    Será necessário <b>salvar a imagem da sua assinatura em seu computador</b><br>
    e <b>acessar o manual referente à ferramenta de e-mail utilizada: Outlook, Thunderbird ou Webmail (OWA).</b><br></p>'''

    txtPrincipal3 = '''<p style='font-family:calibri;'>Caro colaborador(a),<br><br>
    <b>Alguns ajustes foram realizados em sua assinatura de e-mail.</b><br>
    Portanto, <b>você deverá substituir a assinatura em uso pela que se encontra em anexo.</b><br></p>'''

    return txtPrincipal, txtPrincipal2, txtPrincipal3, assLucasP

def envia_email():
    txtPrincipal, txtPrincipal2, txtPrincipal3, assLucasP = variaveis()

    outlook = win32.Dispatch('outlook.application') # Abrir o Outlook
    mail = outlook.CreateItem(0)  # Criar um novo email

    mail.Display()  # Exibir o email
    
    mail.To = 'lucasprs@camed.com.br' # Destinatário
    mail.CC = '' # Com Cópia
    
    mail.Subject = 'Assunto do Email' # Assunto
    
    mail.Body = 'Corpo do Email' # Corpo do Email
    mail.HTMLBody = txtPrincipal + assLucasP

    # Anexos usar imagem assLucasP da pasta do projeto
    assRemetente = os.getcwd() + '\\AssLucasP.png'
    mail.Attachments.Add(assRemetente)

    #mail.Send() # Enviar Email

envia_email()

print('Emails Enviados com Sucesso')
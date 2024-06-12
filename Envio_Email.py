# Envio de Email em Lote usando outlook
# Outlook object com funlçao

import win32com.client as win32
import os
import time

def variaveis():
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

    assLucasP = '''<span style="font-family:calibri;"><br>Atenciosamente,</span><br>
    <img src="file:///E:/Bkp/pc/Programas camed/Lucas P/MPython/Envio_Email_Lote/ImgLucasP.png">'''


    return txtPrincipal, txtPrincipal2, txtPrincipal3, assLucasP

def envia_email():
    outlook = win32.Dispatch('outlook.application')
    mail = outlook.CreateItem(0)  # Criar um novo email
    mail.Display()  # Exibir o email
    mail.To = 'guilhermesla@camed.com.br' # Destinatário
    mail.CC = '' # Com Cópia
    mail.Subject = 'Assunto do Email' # Assunto
    mail.Body = 'Corpo do Email' # Corpo do Email
    mail.HTMLBody = variaveis()[0] + variaveis()[3]

    # Anexos
    #attachment = os.getcwd() + '\\imgLucasP.png'
    #mail.Attachments.Add(attachment)

    mail.Send() # Enviar Email

i = 0
while i < 1:
    envia_email()
    i += 1
    print(i)
print('Emails Enviados com Sucesso')
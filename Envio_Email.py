# Envio de Email em Lote usando outlook
# Outlook object com funlçao

import win32com.client as win32
import os

def envia_email():
    outlook = win32.Dispatch('outlook.application')
    mail = outlook.CreateItem(0)  # Criar um novo email
    mail.To = '' # Destinatário
    mail.Subject = 'Assunto do Email' # Assunto
    mail.Body = 'Corpo do Email' # Corpo do Email
    mail.HTMLBody = '<h2>Corpo do Email</h2>' # Corpo do Email em HTML

    # Anexos
    attachment = os.getcwd() + '\\arquivo.txt'
    mail.Attachments.Add(attachment)

    mail.Send() # Enviar Email

envia_email()
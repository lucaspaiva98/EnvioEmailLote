Sub enviar_EMAILS()
    Dim resultado As VbMsgBoxResult
    Dim arquivo As String
    'Dim obj_outlook As Outlook.Application
    
    Application.ScreenUpdating = False
    Remetente = Cells(7, 1).Value
    
    '---------------------------------------'
    'Destinatarios CC
    Lucas = "lucasprs@camed.com.br"
    Wollney = "WollneyMR@camedseguros.com.br"
    '---------------------------------------'
    
    'ok = "no"
    
    'Do Until ok = "yes"
        If Remetente = "VAZIO" Then
            MsgBox ("INFORMAR REMETENTE")
            Cells(7.1).Select
        Else
            If Cells(9, 1) = "" Then
                MsgBox ("Primeiro Funcionário não encontrado")
                Cells(9, 1).Select
            Else
                If Cells(4, 9) = "FALHA" Then
                    MsgBox ("Problema(s) encontrado(s) na planilha")
                'Else
                    'ok = "yes"
                End If
            End If
        End If
    'Loop
        
        '----Assinatura remetente----
        'Buscando a assinatura do Remetente
        If Remetente = "Lucas" Then
            ass = "<span style='font-family:calibri;'><br>Atenciosamente,</span><br>" _
                & "<img src='O:\Ger_Tecnologia\Tecnologia\LucasP\Assinaturas\NOVA INTRANET\ENVIO\Ass_remetentes\LucasP.png'>"
        Else
            If Remetente = "Wollney" Then
                ass = "<span style='font-family:calibri;'><br>Atenciosamente,</span><br>" _
                    & "<img src='O:\Ger_Tecnologia\Tecnologia\LucasP\Assinaturas\NOVA INTRANET\ENVIO\Ass_remetentes\WollneyM.png'>"
            End If
        End If
        '-----Fim Ass remetente--------
        
        resultado = MsgBox("Quem está enviando é: " & Cells(7, 1) & ". -> Tem certeza que deseja prosseguir com esta ação?", vbYesNo, "Envio de E-mails")
        
        If resultado = vbYes Then
        
            Set obj_outlook = CreateObject("Outlook.Application")
            
            lin = 9 'linha inicial
            cont = 0 'contador
            
            'Quantidade de linhas
            Do Until Cells(lin, 1) = ""
                cont = cont + 1
                lin = lin + 1
            
            Loop
            
            'Barra de progresso
            ThisWorkbook.Sheets(1).PBar1.Max = cont
            ThisWorkbook.Sheets(1).PBar1.Value = 0
            
            lin = 9
            '---------------------------------------------------------------------------------------'
            Do Until Cells(lin, 1) = ""
                
                Set Email = obj_outlook.createitem(0)
                
                'Destinatario - Copia - CCO (linha, coluna)
                Email.To = Cells(lin, 7).Value
                Email.cc = "LucasPRS@camed.com.br;WollneyMR@camedseguros.com.br;RodrigoDD@camed.com.br;joaolsr@camed.com.br"
                'Email.display
                'Email.bcc = "maxfbl@camed.com.br;joaoprc@camed.com.br"
                
                
                '--------Assunto----------
                Email.Subject = "ASSINATURA DE E-MAIL - " & Cells(lin, 1)
                'Email.Subject = "CORREÇÃO NA ASSINATURA DE E-MAIL - " & Cells(lin, 1)
                
                '-----------TEXTOS PARA ALTERAÇÃO DE ASSINATURA----------
                
                texto1 = "<p style='font-family:calibri;'>Caro colaborador,<br><br>" _
                & "Com o objetivo de promovermos a padronização das assinaturas dos e-mails dos funcionários do Grupo Camed, " _
                & "<b>estamos enviando sua assinatura ajustada em anexo, juntamente com o manual que contém o passo a passo que deverá ser seguido</b> para a atualização desta.<br>" _
                & "Logo, <b>será necessário SALVAR a imagem enviada em anexo em seu computador</b> " _
                & "e <b>ACESSAR o manual referente a ferramenta de e-mail utilizado: Outlook, Thunderbird ou Webmail (OWA).</b><br></p>"
                
                'texto1 = "<p style='font-family:calibri;'>Caro colaborador,<br><br>" _
                & "<b>Estaremos enviando sua assinatura ajustada em anexo juntamente o manual com o passo a passo que deverá ser seguido</b>, para a atualização da mesma.<br>" _
                & "Com isso, <b>será necessário salvar a imagem de sua assinatura em seu computador</b><br> " _
                & "E acessar o manual referente a ferramenta de e-mail utilizado,se Outlook, Thunderbird ou Webmail(OWA).<br></p>"
                
                'texto1 = "<p style='font-family:calibri;'>Caro colaborador,<br><br>" _
                & "<b>Alguns ajustes foram realizados em sua assinatura de EMAIL.</b><br>" _
                & "Com isso, <b>você deverá estar substituindo a assinatura em uso por a que se encontra em anexo!</b><br> "
                
                
                '------DADOS---------
                texto2 = "<p style='font-family:calibri;'>Seguem as informações da assinatura em anexo:<br><br>" _
                & "<u>Nome</u>: " & Cells(lin, 1).Value & "<br>" _
                & "<u>Setor</u>: " & Cells(lin, 2).Value & "<br>" _
                & "<u>Função</u>: " & Cells(lin, 3).Value & "<br>" _
                & "<u>Celular</u>: " & Cells(lin, 4).Value & "<br>" _
                & "<u>Telefone</u>: " & Cells(lin, 5).Value & "<br>" _
                & "<u>Ramal</u>: " & Cells(lin, 6).Value & "<br>" _
                & "<u>E-mail</u>: " & Cells(lin, 7).Value & "<br></p>"
                
                
                texto3 = "<p style='font-family:calibri;'>Em caso de dúvidas na execução do procedimento ou em caso de mudança na assinatura enviada, " _
                & "retornar este e-mail clicando em <b>Responder a Todos</b>." _
                & "<br><br>Desde já, agradecemos sua colaboração!<br></p>"
                
                'texto3 = "<p style='font-family:calibri;'><br><b>Por gentileza, confirmar a alteração da Assinatura respondendo este e-mail com a assinatura alterada.<b><br>Em caso de dúvidas na execução do procedimento a equipe de HelpDesk estará à disposição. (Ramal 7819)<br></p>"
                
                'corpo html do email
                'Email.htmlbody = texto1 & texto2 & texto3
                Email.htmlbody = texto1 & texto2 & texto3 & ass
                'Email.htmlbody 'para retornar corpo do email
                
                'Incluindo anexo
                Email.attachments.Add (ThisWorkbook.Path & "\Manual de Inclusão de Assinatura(Outlook-Thunderbird-Webmail).pdf")
                    
                'Anexo personalizado
                Email.attachments.Add (ThisWorkbook.Path & "\Arquivos\" & Cells(lin, 8).Value)
                'ThisWorkbook.Path para pegar a pasta atual do arquivo excel
                 
                Email.display 'Abri outlook para carregar imagens
                'Email.send 'Para enviar email diretamente
                    
                Set Email = Nothing
                l_atual = lin
                lin = lin + 1
                
                'Barra de progresso
                ThisWorkbook.Sheets(1).PBar1 = ThisWorkbook.Sheets(1).PBar1.Value + 1
                Application.ScreenUpdating = True
                Application.ScreenUpdating = False
                
            Loop
            '------------------------------------------------------------------------------------------------------------'
            
            MsgBox "E-mails enviados com sucesso até a linha " _
            & l_atual & "."
            
            'Barra de progresso original
            ThisWorkbook.Sheets(1).PBar1.Value = 0
            Application.ScreenUpdating = True
            Application.ScreenUpdating = False
            
            Cells(l_atual, 9).Select
        Else
            MsgBox "Envios Cancelados"
        End If
End Sub

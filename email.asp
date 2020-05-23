<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title>Untitled Document</title>
</head>

<body>
<%
			strFrom = "frodrigues@conceitobrazil.com.br"
			strTo = "eduardo@respect.art.br"
			strCc = ""
			strSubject = "Convite evento:"
			strBody = "Convite evento"

	' Cria o objeto CDOSYS
	 Set objCDOSYSMail = Server.CreateObject("CDO.Message")
	 
	 'Cria o objeto para configuração do SMTP
	 Set objCDOSYSCon = Server.CreateObject ("CDO.Configuration")
	 
	 'SMTP 
	 objCDOSYSCon.Fields("http://schemas.microsoft.com/cdo/configuration/smtpserver") = "localhost"
	 
	 'Porta do SMTP
	 objCDOSYSCon.Fields("http://schemas.microsoft.com/cdo/configuration/smtpserverport")= 25
	 
	 'Porta do CDO
	 objCDOSYSCon.Fields("http://schemas.microsoft.com/cdo/configuration/sendusing") = 2
	 
	 'Timeout
	 objCDOSYSCon.Fields("http://schemas.microsoft.com/cdo/configuration/smtpconnectiontimeout") = 30
	 objCDOSYSCon.Fields.update
	 
	 'Atualiza a configuração do CDOSYS para envio do e-mail
	 Set objCDOSYSMail.Configuration = objCDOSYSCon
	 
	 ' #### CONFIGURAÇÕES DO CABEÇALHO DA MENSAGEM ####
	 'Configura o remetente(FROM)
	 objCDOSYSMail.From = trim(strFrom)
	 
	 'Configura o destinatário(TO)
	 objCDOSYSMail.To = trim(strTo)
	 objCDOSYSMail.Cc = trim(strCc)
	 
	 'Configura o Reply-To(Responder Para) 
	 objCDOSYSMail.ReplyTo = trim(strTo)
	 
	 'Configura Copia
	 If Trim(strToCopia) <> "" Then
		objCDOSYSMail.Cc = strToCopia
	 End if

	 'Configura o assunto(SUBJECT)
	 objCDOSYSMail.Subject = strSubject
	 
	 'Configura o conteúdo da mensagem 
	 'Para enviar mensagens no formato HTML, altere o TextBody para HtmlBody
	 objCDOSYSMail.HtmlBody = strBody 
	 
	 ' ### ENVIA O E-MAIL ###
		enviaEmail = False
		On Error Resume Next 
		  objCDOSYSMail.Send ' Enviando
		If Err <> 0 Then ' Erros
		  enviaEmail = Err.Description
		  response.Write(enviaEmail)
		else
		  enviaEmail = 1
		End If
		
		
	 
	 ' ### DESTRÓI OS OBJETOS ###
	 Set objCDOSYSMail = Nothing
	 Set objCDOSYSCon = Nothing

%>
</body>
</html>

<%
'Abre conex�o com o banco de dados.
dim adoConn
dim adoRS
dim counter
set adoConn = Server.CreateObject("ADODB.Connection")
adoConn.Open "Driver={MariaDB ODBC 3.1 Driver};SERVER=192.168.40.15;USER=conceitobrazil;PASSWORD=@9vRp7i5;DATABASE=conceitobrazil;PORT=3306"

'Local
'adoConn.Open "DSN=conceitobrazil"

Public Sub gravaLog(id_usu,tela,acao,obs)
	set atualiza = adoConn.execute("insert into log (id_usu,tela,acao,obs) values ("&id_usu&",'"&tela&"','"&acao&"','"&obs&"')")
	set atualiza = nothing
End Sub

	
Public function enviaEmail(strFrom,strTo,strCc,strToCopia,strSubject,strBody)	
	strSchema = "http://schemas.microsoft.com/cdo/configuration/" 
	Set objCDOConfig = Server.CreateObject("CDO.Configuration") 
	With objCDOConfig.Fields 
				.Item(strSchema & "smtpusessl") = True 
				.Item(strSchema & "smtpauthenticate") = 1 
				.Item(strSchema & "sendusername") = "operacional2@conceitobrazil.com.br"
				.Item(strSchema & "sendpassword") = "!@#Psystem"
				.Item(strSchema & "smtpserver") = "smtp.office365.com"
				.Item(strSchema & "smtpserverport") = 25
				.Item(strSchema & "sendusing") = 2 
				.Item(strSchema & "smtpconnectiontimeout") = 30 
				.Update 
	End With 
	Set objCDOMessage = Server.CreateObject("CDO.Message") 
	With objCDOMessage 			
				Set .Configuration = objCDOConfig 
				.From = strFrom
				.ReplyTo = strFrom
				'.Cc	= "marketing@conceitobrazil.com.br; frodrigues@conceitobrazil.com.br"
				.To = strTo
				.Subject = strSubject 
				.HtmlBody = strBody 
				.Send 
	End With
End function
	
%>
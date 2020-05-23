<%
'Abre conex�o com o banco de dados.
dim adoConn
dim adoRS
dim counter
set adoConn = Server.CreateObject("ADODB.Connection")
'adoConn.CursorLocation=3
'adoConn.Open "Driver=MariaDB ODBC 3.1.7 Driver; Server=http://banco.conceitobrazil.com.br/; Database=conceitobrazil Uid=conceitobrazil; Pwd=@9vRp7i5; Option=3"

'adoConn.Open("DRIVER={MariaDB ODBC 3.1.7 Driver};SERVER=localhost;PORT=3306;DATABASE=conceitobrazil;USER=conceitobrazil;PASSWORD=@9vRp7i5;OPTION=3;")
'adoConn.Open("DRIVER={MariaDB ODBC 3.1.7 Driver};SERVER=209.50.57.61;PORT=3306;DATABASE=conceitobrazil;USER=conceitobrazil;PASSWORD=@9vRp7i5;OPTION=3;")
'adoConn.Open("DRIVER={MariaDB ODBC 3.1.7 Driver};SERVER=http://banco.conceitobrazil.com.br/;PORT=3306;DATABASE=conceitobrazil;USER=conceitobrazil;PASSWORD=@9vRp7i5;OPTION=3;")


'adoConn.Open("DRIVER={MySQL ODBC 5.3 ANSI Driver};SERVER=localhost;PORT=3306;DATABASE=conceitobrazil;USER=conceitobrazil;PASSWORD=@9vRp7i5;OPTION=3;")
'adoConn.Open("DRIVER={MySQL ODBC 5.3 ANSI Driver};SERVER=209.50.57.61;PORT=3306;DATABASE=conceitobrazil;USER=conceitobrazil;PASSWORD=@9vRp7i5;OPTION=3;")
'adoConn.Open("DRIVER={MySQL ODBC 5.3 ANSI Driver};SERVER=http://banco.conceitobrazil.com.br;PORT=3306;DATABASE=conceitobrazil;USER=conceitobrazil;PASSWORD=@9vRp7i5;OPTION=3;")


'adoConn.Open "DRIVER={MySQL ODBC 5.3 **Unicode Driver**};SERVER=localhost; DATABASE=conceitobrazil; UID=conceitobrazil; PASSWORD=@9vRp7i5; Port=3306; OPTION=3"
'adoConn.Open "DRIVER={MySQL ODBC 5.3 **Unicode Driver**};SERVER=209.50.57.61; DATABASE=conceitobrazil; UID=conceitobrazil; PASSWORD=@9vRp7i5; Port=3306; OPTION=3"
'adoConn.Open "DRIVER={MySQL ODBC 5.3 **Unicode Driver**};SERVER=http://banco.conceitobrazil.com.br; DATABASE=conceitobrazil; UID=conceitobrazil; PASSWORD=@9vRp7i5; Port=3306; OPTION=3"

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
				.Item(strSchema & "sendusername") = "rodada@conceitobrazil.com.br"
				.Item(strSchema & "sendpassword") = "rdada2012"
				.Item(strSchema & "smtpserver") = "smtp.gmail.com"
				.Item(strSchema & "smtpserverport") = 465
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
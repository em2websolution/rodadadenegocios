<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!-- #Include File="includes/abre_conexao.asp" --> 

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
</head>

<body>
<%

'Envio automatico de notificação sobre a ação
strFrom = "rodada@conceitobrazil.com.br"
strTo = "eduardo@respect.art.br"
strSubject = "Convite para a Rodada de Negócios "
strBody = "Ola Edu"



call enviaEmail(strFrom,strTo,"","",strSubject,strBody)

%>
</body>
</html>

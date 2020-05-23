<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title>Conceito Brazil - Rodada de Negócios</title>
<link rel="stylesheet" type="text/css" href="css/style_new.css"/>
<link rel="stylesheet" type="text/css" href="css/style_form.css"/>
<!-- #Include File="includes/abre_conexao.asp" --> 
</head>

<body>
<%
	set adoRS = Server.CreateObject("ADODB.Recordset")
	adoRS.ActiveConnection = adoConn
	adoRS.Open("select * from tb_login where id_usu="&request("id")&" and stu=1")
	
	empresa = adors("empresa")
	emp_fantasia = adors("emp_fantasia")
	perfil = adors("perfil")
%>
<table width="100%" border="0" cellspacing="2" cellpadding="2">
  <tr>
    <th width="35%" scope="row"><div align="right"><strong>Razão Social<br />
    <span class="phone" style="font-size:10px">Company Name</span><span class="phone" style="font-size:10px">:</span></strong></div></th>
    <td width="65%"><%= empresa%></td>
  </tr>
  <tr>
    <th scope="row"><div align="right"><strong>Nome Fantasia<br />
    <span class="phone" style="font-size:10px">Fantasy Name</span><span class="phone" style="font-size:10px">:</span></strong></div></th>
    <td><%= emp_fantasia%></td>
  </tr>
  <tr>
    <th scope="row"><div align="right"><span style="font-family: Verdana, Arial, Helvetica, sans-serif; font-size: 12px; font-weight: bold">Informação da Empresa<br />
    <span class="phone" style="font-size:10px" id="result_box5" lang="en" xml:lang="en">Company Information:</span></span></div></th>
    <td><%= perfil%></td>
  </tr>
</table>
</body>
</html>

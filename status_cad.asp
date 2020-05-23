<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!-- #Include File="includes/abre_conexao.asp" --> 
<%
if request.QueryString("stu") = 1 then
	ativa = 0
else
	ativa = 1
end if

set atualizar = Server.CreateObject("ADODB.Recordset")
atualizar.ActiveConnection = adoConn
atualizar.Open("update tb_login set stu="&ativa&" where id_usu = "&request.QueryString("id")&"")

Set atualizar = Nothing

response.Redirect("cadastros.asp")
response.End()
%>
<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!-- #Include File="abre_conexao.asp" --> 
<%
	id_rod = trim(request.Form("id_rod"))
	id_usu = trim(request.Form("id_usu"))

	set executa = Server.CreateObject("ADODB.Recordset")
	executa.ActiveConnection = adoConn
	executa.Open("update tb_indicacao set id_usu = null, stu = 1, dta_indicado = '"&year(now()) &"-"& month(now()) &"-"& day(now())&" "&hour(now())&":"&minute(now())&"' where id_rod="&id_rod&" and id_usu="&id_usu&"")

	set executa = nothing
	
	'Grava log
	Call gravaLog(session("user_rd"),"Cancelamento de Participação","Cancelamento","Empresa: "&session("empresa")&" - Usuário: "&session("user_name")&"")

%>

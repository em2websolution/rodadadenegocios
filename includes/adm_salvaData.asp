<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!-- #Include File="abre_conexao.asp" --> 
<%
	id_rod = trim(request.Form("id_rod"))
	id_usu = session("user_rd")
	dias = trim(request.Form("dias"))
	dias = left(dias,len(dias)-1)


	set executa = Server.CreateObject("ADODB.Recordset")
	executa.ActiveConnection = adoConn

	set lista = Server.CreateObject("ADODB.Recordset")
	lista.ActiveConnection = adoConn
	lista.Open("update tb_indicacao set dias='"&dias&"' where id_rod="&id_rod&" and id_usu="&id_usu&"")	
	set lista = nothing
	set executa = nothing
%>

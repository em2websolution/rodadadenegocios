<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!-- #Include File="abre_conexao.asp" --> 
<%
	id_rod = trim(request.Form("id_rod"))
	id_usu = trim(request.Form("id_usu"))
	chk = trim(request.Form("chk"))
	

	set executa = Server.CreateObject("ADODB.Recordset")
	executa.ActiveConnection = adoConn

	set lista = Server.CreateObject("ADODB.Recordset")
	lista.ActiveConnection = adoConn
	lista.Open("select * from tb_selecao where id_rod="&id_rod&" and id_usu="&id_usu&"")
	
	if not lista.eof then
		executa.Open("delete from tb_selecao where id_rod="&id_rod&" and id_usu="&id_usu&"")
	else
		executa.Open("insert into tb_selecao (id_rod,id_usu) values ("&id_rod&","&id_usu&")")
	end if
	
	set lista = nothing
	set executa = nothing
%>

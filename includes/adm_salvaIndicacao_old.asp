<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!-- #Include File="abre_conexao.asp" --> 
<%
	id_rod = trim(request.Form("id_rod"))
	id_usu = session("user_rd")
	id_indicado = trim(request.Form("id_usu"))
	chk = trim(request.Form("chk"))
	dias = trim(request.Form("dias"))
	
	dias = left(dias,len(dias)-1)


	set executa = Server.CreateObject("ADODB.Recordset")
	executa.ActiveConnection = adoConn

	set lista = Server.CreateObject("ADODB.Recordset")
	lista.ActiveConnection = adoConn
	lista.Open("select * from tb_indicacao where id_rod="&id_rod&" and id_usu="&id_usu&" and id_indicado="&id_indicado&"")
	
	if not lista.eof then
		executa.Open("delete from tb_indicacao where id_rod="&id_rod&" and id_usu="&id_usu&" and id_indicado="&id_indicado&"")
	else
		executa.Open("insert into tb_indicacao (id_rod,id_usu,id_indicado,dias) values ("&id_rod&","&id_usu&","&id_indicado&",'"&dias&"')")
	end if
	
	set lista = nothing
	set executa = nothing
%>

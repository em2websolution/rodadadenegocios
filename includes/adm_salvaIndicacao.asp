<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!-- #Include File="abre_conexao.asp" --> 
<%
	id_ind = int(request.Form("id_ind"))
	id_usu = int(request.Form("id_usu"))
	id_ind_old = int(request.Form("id_ind_old"))
	id_indicado = int(request.Form("id_indicado"))
	id_rod = int(request.Form("id_rod"))
	
	set executa = Server.CreateObject("ADODB.Recordset")
	executa.ActiveConnection = adoConn
	
	if id_ind_old <> 0 then
		sql = ("update tb_indicacao set id_usu = null, stu = 1, dta_indicado = '"&year(now()) &"-"& month(now()) &"-"& day(now())&" "&hour(now())&":"&minute(now())&"' where id_ind="&id_ind_old&" ")
		response.Write(sql&"<br>")
		executa.Open(sql)
	end if
	
	set lista = Server.CreateObject("ADODB.Recordset")
	lista.ActiveConnection = adoConn
	lista.Open("select * from tb_indicacao where id_rod="&id_rod&" and id_usu="&id_usu&" and id_indicado="&id_indicado&"")
	
	if not lista.eof then
		executa.Open("update tb_indicacao set id_usu = null, stu = 1, dta_indicado = '"&year(now()) &"-"& month(now()) &"-"& day(now())&" "&hour(now())&":"&minute(now())&"' where id_ind="&lista("id_ind")&" ")
	end if

	
	sql = ("update tb_indicacao set id_usu = "&id_usu&", stu = 0, dta_indicado = '"&year(now()) &"-"& month(now()) &"-"& day(now())&" "&hour(now())&":"&minute(now())&"' where id_ind="&id_ind&" ")
	
	
	response.Write("Horario salvo com sucesso" &sql&"<br>")
		
	executa.Open(sql)
	set executa = nothing

%>

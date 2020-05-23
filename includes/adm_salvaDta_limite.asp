<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!-- #Include File="abre_conexao.asp" --> 
<%
	id_rod = trim(request("id_rod"))
	dta_limite = year(request.Form("dta_limite")) &"-"& month(request.Form("dta_limite")) &"-"& day(request.Form("dta_limite"))
		
	set executa = Server.CreateObject("ADODB.Recordset")
	executa.ActiveConnection = adoConn
	executa.Open("update tb_rodada set dta_limite = '"&dta_limite&"' where id_rod="&id_rod&" ")
	set executa = nothing
%>

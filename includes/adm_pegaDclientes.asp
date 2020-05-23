<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!-- #Include File="abre_conexao.asp" --> 
<%
	id = int(request("id"))

	if session("lv_user") = 1 or session("lv_user") = 2 then	
		set lista = Server.CreateObject("ADODB.Recordset")
		lista.ActiveConnection = adoConn
		lista.Open("select * from tb_login where id_usu="&id&"")
		if not lista.eof then
			response.Write(lista.fields("id_usu")&"|"&lista.fields(1)&"|"&lista.fields(2)&"|"&lista.fields(3)&"|"&lista.fields(4)&"|"&lista.fields(5)&"|"&lista.fields(6)&"|"&lista.fields(7)&"|"&lista.fields(8)&"|"&lista.fields(9)&"|"&lista.fields(10)&"|"&lista.fields(11)&"|"&lista.fields(12)&"|"&lista.fields(14)&"|"&lista.fields(15)&"|"&lista.fields(16)&"|"&lista.fields(17)&"|"&lista.fields(18)&"|"&lista.fields(19)&"|"&lista.fields(20)&"|"&lista.fields(21)&"|"&lista.fields(22)&"|"&lista.fields(23)&"|"&lista.fields(24)&"|"&lista.fields(25)&"|"&lista.fields(26)&"|"&lista.fields(27)&"|"&lista.fields(28)&"|"&lista.fields(29)&"")
		end if
	else 
		if session("lv_user") = 3 or session("lv_user") = 4 then
			if session("user_rd") = id then
				set lista = Server.CreateObject("ADODB.Recordset")
				lista.ActiveConnection = adoConn
				lista.Open("select * from tb_login where id_usu="&id&"")
				if not lista.eof then
					response.Write(lista.fields("id_usu")&"|"&lista.fields(1)&"|"&lista.fields(2)&"|"&lista.fields(3)&"|"&lista.fields(4)&"|"&lista.fields(5)&"|"&lista.fields(6)&"|"&lista.fields(7)&"|"&lista.fields(8)&"|"&lista.fields(9)&"|"&lista.fields(10)&"|"&lista.fields(11)&"|"&lista.fields(12)&"|"&lista.fields(14)&"|"&lista.fields(15)&"|"&lista.fields(16)&"|"&lista.fields(17)&"|"&lista.fields(18)&"|"&lista.fields(19)&"|"&lista.fields(20)&"|"&lista.fields(21)&"|"&lista.fields(22)&"|"&lista.fields(23)&"|"&lista.fields(24)&"|"&lista.fields(25)&"|"&lista.fields(26)&"|"&lista.fields(27)&"|"&lista.fields(28)&"|"&lista.fields(29)&"")
				end if
			else
				response.Write("false")
			end if
		end if
	end if
%>

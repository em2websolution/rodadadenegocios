<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!-- #Include File="abre_conexao.asp" --> 
<%
	id = trim(request.Form("id"))
	set lista = Server.CreateObject("ADODB.Recordset")
	lista.ActiveConnection = adoConn
	lista.Open("select * from tb_rodada where id_rod="&id&"")
	if not lista.eof then
		
		
		dta_i = year(lista.fields(3)) &"/"& month(lista.fields(3)) &"/"& day(lista.fields(3))
		dta_f = year(lista.fields(4)) &"/"& month(lista.fields(4)) &"/"& day(lista.fields(4))
		
		if hour(lista.fields(3)) < 10 then 
			h1 = "0" & hour(lista.fields(3))
		else
			h1 = hour(lista.fields(3))
		end if

		if minute(lista.fields(3)) < 10 then 
			m1 = "0" & minute(lista.fields(3))
		else
			m1 = minute(lista.fields(3))
		end if

		if hour(lista.fields(4)) < 10 then 
			h2 = "0" & hour(lista.fields(4))
		else
			h2 = hour(lista.fields(4))
		end if

		if minute(lista.fields(4)) < 10 then 
			m2 = "0" & minute(lista.fields(4))
		else
			m2 = minute(lista.fields(4))
		end if

		
		hs_i = h1 &":"& m1
		hs_f = h2 &":"& m2
		
		mensagem = lista.fields(13)
		
		if mensagem <> "" then
			mensagem = replace(mensagem,"<BR>",chr(13))
		else
			mensagem = ""
		end if
		
		response.Write(lista.fields(0)&"|"&lista.fields(1)&"|"&lista.fields(2)&"|"&dta_i&"|"&dta_f&"|"&hs_i&"|"&hs_f&"|"&lista.fields(5)&"|"&lista.fields(6)&"|"&lista.fields(7)&"|"&lista.fields(8)&"|"&lista.fields(9)&"|"&lista.fields(10)&"|"&lista.fields(12)&"|"&mensagem&"|"&lista.fields(14)&"|"&lista.fields(15)&"|"&lista.fields(16)&"")
	end if
%>

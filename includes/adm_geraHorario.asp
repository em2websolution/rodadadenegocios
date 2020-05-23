<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!-- #Include File="abre_conexao.asp" --> 
<%
	id_rod = trim(request.Form("id_rod"))
	id_indicado = trim(request.Form("id_indicado"))
	id_ind = trim(request.Form("id_ind"))
	id_usu = int(request.Form("id_usu"))

	if request.Form("dta") <> "12/12/2000" then
		if day(request.Form("dta")) < 10 then 
			dia = "0" & day(request.Form("dta"))
		else
			dia = day(request.Form("dta"))
		end if
		
		if month(request.Form("dta")) < 10 then 
			mes = "0" & month(request.Form("dta"))
		else
			mes = month(request.Form("dta"))
		end if
		
		dta = year(request.Form("dta")) &"-"& mes &"-"& dia
		sql = ("SELECT DATE_FORMAT(I.hr,'%H:%i') as Horario, i.id_ind FROM conceitobrazil.tb_indicacao as I inner join conceitobrazil.tb_login as L on I.id_indicado = L.id_usu where I.id_rod ="&id_rod&" and I.id_indicado="&id_indicado&" and I.stu = 1 and DATE_FORMAT(I.hr,'%Y-%m-%d') = '"&dta&"' order by Horario")
	
		set lista = Server.CreateObject("ADODB.Recordset")
		lista.ActiveConnection = adoConn
		lista.Open(sql)
		
		if not lista.eof then
	%>
			<select name="hr<%=id_indicado%>" id="hr<%=id_indicado%>" onChange="salva_selecao(this.value,<%=id_indicado%>,0,<%=id_usu%>)">
				<option value="0">Escolha o horário</option>
			  <% while not lista.eof
			  
					set verifica = Server.CreateObject("ADODB.Recordset")
					verifica.ActiveConnection = adoConn
					sql = ("SELECT i.id_ind FROM conceitobrazil.tb_indicacao as i inner join conceitobrazil.tb_login as L on I.id_indicado = L.id_usu where I.id_rod ="&id_rod&"  and i.stu = 0 and I.hr= '"&dta&" "&lista.fields(0)&"' and i.id_usu <> 'null' and i.id_usu="&id_usu&"")
					verifica.Open(sql)
					
					if verifica.eof then
			  %>
					<option value="<%=lista.fields(1)%>"><%=lista.fields(0)%></option>
			  <% 
					end if
					set verifica = nothing
			  lista.movenext
			  wend %>
		   </select>
	<%		
		else
			response.Write("Horários não disponíveis" & sql)
		end if
	else
		'Cancela a reunião do dia
		set executa = Server.CreateObject("ADODB.Recordset")
		executa.ActiveConnection = adoConn
		sql = ("update tb_indicacao set id_usu = null, stu = 1, dta_indicado = '"&year(now()) &"-"& month(now()) &"-"& day(now())&" "&hour(now())&":"&minute(now())&"' where id_ind="&id_ind&"")
		executa.Open(sql)
		
	end if
	
	set lista = nothing
%>

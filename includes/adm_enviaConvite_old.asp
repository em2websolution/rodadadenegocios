﻿<!-- #Include File="abre_conexao.asp" --> 
<%
	id_rod = trim(request("id_rod"))
	
	set executa = Server.CreateObject("ADODB.Recordset")
	executa.ActiveConnection = adoConn

	set lista = Server.CreateObject("ADODB.Recordset")
	lista.ActiveConnection = adoConn
	
	set registros = Server.CreateObject("ADODB.Recordset")
	registros.ActiveConnection = adoConn
	
	sql = ("SELECT id_sel FROM conceitobrazil.tb_selecao as S where S.id_rod = "&id_rod&" and S.convites = 0")
	registros.Open(sql)

	if not registros.eof then
		while not registros.eof	
			sql2 = ("SELECT * FROM conceitobrazil.tb_selecao as S inner join conceitobrazil.tb_login as L on S.id_usu = L.id_usu inner join conceitobrazil.tb_rodada as R on R.id_rod = S.id_rod where S.id_rod = "&id_rod&" and S.convites = 0 and L.email <> '' and id_sel="&registros.fields(0)&" order by L.nivel")
			lista.Open(sql2)
			if not lista.eof then
				mensagem = trim(lista("mensagem"))
				if mensagem <> "" then
					retorno = 0
					if lista("nivel") = 3 then
						emp_vendedores = "<b>" & lista("empresa") & "</b> - "& lista("site") &" <br> " &  lista("perfil") & "<br><br>" & emp_vendedores
					else			
						dta_i = day(lista("dta_ini")) &"/"& month(lista("dta_ini")) &"/"& year(lista("dta_ini"))
						dta_f = day(lista("dta_fim")) &"/"& month(lista("dta_fim")) &"/"& year(lista("dta_fim"))
						if hour(lista("dta_ini")) < 10 then 
							h1 = "0" & hour(lista("dta_ini"))
						else
							h1 = hour(lista("dta_ini"))
						end if
					
						if minute(lista("dta_ini")) < 10 then 
							m1 = "0" & minute(lista("dta_ini"))
						else
							m1 = minute(lista("dta_ini"))
						end if
					
						if hour(lista("dta_fim")) < 10 then 
							h2 = "0" & hour(lista("dta_fim"))
						else
							h2 = hour(lista("dta_fim"))
						end if
					
						if minute(lista("dta_fim")) < 10 then 
							m2 = "0" & minute(lista("dta_fim"))
						else
							m2 = minute(lista("dta_fim"))
						end if
						
						hs_i = h1 &":"& m1
						hs_f = h2 &":"& m2
						
						dta_limite = day(lista("dta_limite")) &"/"& month(lista("dta_limite")) &"/"& year(lista("dta_limite"))
			
						
					
						mensagem = replace(lista("mensagem"),"$$NOME$$",lista("nome"))
						mensagem = replace(mensagem,"$$EMPRESAS$$",emp_vendedores)
						mensagem = replace(mensagem,"$$USER$$",lista("email"))
						mensagem = replace(mensagem,"$$SENHA$$",lista("senha"))					
						mensagem = replace(mensagem,"$$DTA_LIMITE$$",dta_limite)
	
						
						mensagem = "<!DOCTYPE html PUBLIC '-//W3C//DTD XHTML 1.0 Transitional//EN' 'http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd'><html xmlns='http://www.w3.org/1999/xhtml'><head><meta http-equiv='Content-Type' content='text/html; charset=utf-8' /><body>Evento: Rodada de Negócios "&lista("local")&" - Assunto: "&lista("assunto")&" <br> Periodo: "& dta_i & " à " & dta_f &" - Horário: "& hs_i&" às "&hs_f&"<br><br> Prezado(a) Sr.(a), " & lista("nome") &"<br> Favor enviar a relação das empresas com as quais você tem interesse em reuniões durante a "&lista("local")&". <br> Ou <br> Favor registrar suas reuniões no link: <a href='http://rodadadenegocios.conceitobrazil.com.br'> http://rodadadenegocios.conceitobrazil.com.br</a> <br> Usuário: "&lista("email")&" <br> Senha: "&lista("senha")&"<br><font color=red>Data limite para confirmação de participação: "&dta_limite&"</font><br><br>"&lista("mensagem")&"<br>"&emp_vendedores&"<br><br>Atenciosamente, <br><br>Carolinna Souza<br>Responsável Rodada de Negócios<br>marketing@conceitobrazil.com.br<br>Tel: 55 11 3527-5000<br>Cel: 55 11 9 8144-7245</body></head></html>"
						
	
						'Envio automatico de notificação sobre a ação
						strFrom = "rodada@conceitobrazil.com.br"
						strTo = lista("email")
						strSubject = "Convite para a Rodada de Negócios " & lista("local")
						strBody = mensagem
						call enviaEmail(strFrom,strTo,"","",strSubject,strBody)
						
						response.Write(strSubject &"<br>"&strBody&"<br>---------------------------------<br><br>")
			
						if i = 0 then
							retorno = retorno
						else
							retorno = retorno &"|"& i
						end if
						dta_ = year(now()) &"-"& month(now()) &"-"& day(now()) &" "& hour(now()) & ":" & minute(now())
						executa.Open("update tb_selecao set convites = convites+1, dta_convites='"&dta_&"' where id_rod="&id_rod&" and id_usu="&lista("id_usu"))
						i = i + 1
						
						response.End()				
					end if
				else
					response.Write("Mensagem personalizada não cadastrada no projeto!")
				end if
			end if
			lista.close

		registros.movenext
		wend			
	else
		response.Write("Nenhuma nova empresa foi selecionada")
	end if
	
	response.Write(retorno)
	set executa = nothing
	
%>




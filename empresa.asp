<!-- #Include File="top.asp" -->
<script language="javascript" type="text/javascript">
function IniciaAjax(){
	var ajax;
	if (window.XMLHttpRequest) { //Mozila Safari
		ajax = new XMLHttpRequest();
	} else if (window.ActiveXObject) { //IE
		ajax = new ActiveXObject("Msxml2.XMLHTTP");
		if (!ajax){
			ajax = new ActiveXObject("Microsoft.XMLHTTP");
		}
	} else {
		alert("Seu navegador não possui suporte para esta aplicação!")	
	}
	return ajax;
}
</script>
<%

if request("id") = "" then
	response.Redirect("index.asp")
	response.End()
else
	id_rod = request("id")
end if


sql = ("select * from tb_rodada where id_rod="&id_rod&"")
'response.Write(sql)

set dados = Server.CreateObject("ADODB.Recordset")
dados.ActiveConnection = adoConn
dados.Open(sql)

%>
<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td scope="row"><table width="950" border="0" align="center" cellpadding="0" cellspacing="0">
      <tr>
        <td scope="row">
          <% if not dados.eof then
	
		dias = DateDiff("d",dados("dta_ini"),dados("dta_fim")+1)
		hs1=hour(dados("dta_ini"))&":"&minute(dados("dta_ini"))
		hs2=hour(dados("dta_fim"))&":"&minute(dados("dta_fim"))
		minutos = DateDiff("n",hs1,hs2)
		mesas = dados("mesas")
		almoco = Split(dados("almoco"),"-")
		i = 1
		for each x in almoco
		    if i = 1 then hs_ai = x
		    if i = 2 then hs_af = x
			i = i+1
		next
		
		min_a = DateDiff("n",hs_ai,hs_af)		
		
		tempo = dados("tempo")
		intervalo = dados("intervalo")
		
		'Calculo
		'response.Write("(tempo dia: " & minutos & " - tempo almoco: " & min_a & ") / (tempo reuniao: " & tempo & "+ tempo intervalo: " & intervalo & ")" )
		
		reunioes = int((minutos-min_a))/(tempo+intervalo) * mesas
		
		
		select case tempo
			case "60"
				tempo = "1hora"
			case else
				tempo = dados("tempo") & "min"
		end select
		
	dta_i = day(dados("dta_ini")) &"/"& month(dados("dta_ini")) &"/"& year(dados("dta_ini"))
	dta_f = day(dados("dta_fim")) &"/"& month(dados("dta_fim")) &"/"& year(dados("dta_fim"))
	if hour(dados("dta_ini")) < 10 then 
		h1 = "0" & hour(dados("dta_ini"))
	else
		h1 = hour(dados("dta_ini"))
	end if

	if minute(dados("dta_ini")) < 10 then 
		m1 = "0" & minute(dados("dta_ini"))
	else
		m1 = minute(dados("dta_ini"))
	end if

	if hour(dados("dta_fim")) < 10 then 
		h2 = "0" & hour(dados("dta_fim"))
	else
		h2 = hour(dados("dta_fim"))
	end if

	if minute(dados("dta_fim")) < 10 then 
		m2 = "0" & minute(dados("dta_fim"))
	else
		m2 = minute(dados("dta_fim"))
	end if

	
	hs_i = h1 &":"& m1
	hs_f = h2 &":"& m2

	manual = dados("manual")


select case idioma
	case "English"
		var1 = "Location: "
		var2 = "Date/Time: "
		var11 = " day(s) of the event."
		var3 = "Lunch time: "
		var9 = "DATE"
		var10 = "MEETING"
	case "Spanish"
		var1 = "Localizacion: "
		var2 = "Fecha/Hora: "
		var11 = " día(s) del evento."
		var3 = "La hora del almuerzo: "
		var9 = "DATA "
		var10 = "REUNIÓN"
	case else 'Portugues
		var1 = "Local: "
		var2 = "Data/Hora: "
		var11 = " dia(s) de evento."
		var3 = "Horário de almoço: "
		var9 = "DATA"
		var10 = "REUNIÃO"
end select

		banner = dados("banner")
	 %>
          <table width="950" border="0" align="center" cellpadding="2" cellspacing="2">
            <tr>
							<td colspan="5" style="text-align: center;">
								<%
								if banner <> "" then response.Write("<img src='arquivos/"&banner&"' border='0' />")
								%>
								<table width="100%" border="0" cellspacing="0" cellpadding="0">
               		 <tr>
                  	<td scope="row"><h1><%= var1 & dados("local") & " - " & dados("assunto")%></h1><h3><%= "<font color=red>"& mesas &"</font> Mesa(s)  - <font color=red>"& CInt(reunioes) & "</font> reuniões por dia - <font color=red>" & Cint(dias)*cint(reunioes) & "</font> Reuniões em <font color=red> "& dias & var11 &"</font>" %></h3>
											<h2>
												<strong><%=var2%></strong> 								
												<%= (dta_i &" "& hs_i & " à " & dta_f &" "& hs_f &"<br> <b>"&var3&"</b> ("&hs_ai&" às "&hs_af&")")%>
												<br />
										</h2>
										</td>
                  </tr>
								</table>
							</td>
              </tr>
            <tr>
              </tr>
            
        </table>
          <% else 
		response.Write("Registro não encontrado")
		end if	
	%>
          
          </td>
      </tr>
    </table></td>
  </tr>
</table>
<!-- #Include File="rodape.asp" --> 
</body>
</html>

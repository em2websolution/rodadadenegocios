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

function geraPDF(i,f,r1,r2){
	ajax = IniciaAjax();
	if (ajax) {
		ajax.onreadystatechange = function(){
			if(ajax.readyState<4)	
			{
				var carregando = "<div align='right'><img src='img/carregando.gif' /></div>";
				document.getElementById("iconPDF").innerHTML=carregando;
			} else if(ajax.readyState==4){
					document.getElementById('iconPDF').innerHTML=ajax.responseText;
				} 
		}
		dados = 'dta_i='+i+'&dta_f='+f+'&r1='+r1+'&r2='+r2+'&opc=0';	
		ajax.open('POST','pdf.asp',true);
		ajax.setRequestHeader('Content-Type','application/x-www-form-urlencoded');
		ajax.send(dados);
	}
}

function geraXLS() {	
	document.getElementById("icon").innerHTML="Arquivo gerado com sucesso!";
}


</script>
<%
if session("user_name") = "" then
	response.Redirect("index.asp")
	response.End()
end if

if request("id") = "" or session("lv_user") = 3 or session("lv_user") = 4 then
	response.Redirect("projetos.asp")
	response.End()
else
	id_rod = request("id")
end if

if request.QueryString("ordem") = "" then
	ordem = "asc"
else
	if request.QueryString("ordem") = "asc" then
		ordem = "desc"
		img = "images/ico_cinza-arrow_fat_up_g.gif"
	else
		ordem = "asc"
		img = "images/ico_cinza-arrow_fat_down_g.gif"
	end if
end if

if request.QueryString("filtro") = "" then
	filtro = "id_usu"
else
	filtro = request.QueryString("filtro") & ", id_usu"
end if


busca = ""
sql = ("select I.id_ind, I.id_rod, (select count(distinct S2.id_usu) FROM conceitobrazil.tb_indicacao as S2 where I.id_rod = S2.id_rod) as qtd_empresas, (select count(id_usu) from conceitobrazil.tb_indicacao as S2a where I.id_rod = S2a.id_rod) as qtd_reunioes, R.dta_ini, R.dta_fim, R.mesas, R.almoco, R.tempo, R.intervalo, R.local, R.assunto, R.dta_limite, R.manual, I.id_usu, case L.nivel when 1 then 'Administrador' when 2 then 'Gerente' when 3 then 'Vendedor' when 4 then 'Comprador' END as Nivel , L.empresa, L.cnpj, L.emp_fantasia, L.nome, L.cpf, L.email, L.telefone, L.celular, L.perfil, (select count(Si.id_usu) from conceitobrazil.tb_indicacao as Si where Si.id_usu = I.id_usu ) as qtd_indicado, I.dta_indicado, I.dias, (select Di.nivel from conceitobrazil.tb_login as Di where I.id_indicado = Di.id_usu) as nivel_indicada, I.id_indicado, (select Di.emp_fantasia from conceitobrazil.tb_login as Di where I.id_indicado = Di.id_usu) as emp_indicada, I.hr, R.ancora FROM conceitobrazil.tb_indicacao as I inner join conceitobrazil.tb_login as L on I.id_usu = L.id_usu inner join conceitobrazil.tb_rodada as R on R.id_rod = I.id_rod where R.id_rod="&id_rod&" group by I.id_usu, I.id_ind order by "&filtro&" "&ordem&"")

set dados = Server.CreateObject("ADODB.Recordset")
dados.ActiveConnection = adoConn
dados.Open(sql)
'response.Write(sql)

%>
<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <th scope="row">&nbsp;</th>
  </tr>
  <tr>
    <th scope="row">&nbsp;</th>
  </tr>
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
		
		qtd_empresas = cint(dados("qtd_empresas"))
		qtd_reunioes = cint(dados("qtd_reunioes"))
		saldo = ((dias*reunioes)-qtd_reunioes)
		
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
	
	
	 %>
          <table width="950" border="0" align="center" cellpadding="2" cellspacing="2">
            <tr>
              <td colspan="5"><table width="100%" border="0" cellspacing="0" cellpadding="0">
                <tr>
                  <td width="65%" scope="row"><h1>Indicações: <%=dados("local") & " - " & dados("assunto")%></h1></td>
                  <th width="24%"><div id="icon"></div>&nbsp;</th>
                  <td width="6%"><a href="includes/adm_excel.asp?id_rod=<%= id_rod%>&amp;tipo=1" target="_blank"><img src="images/icon_excel.png" alt="Exportar XLS" width="49" height="42" border="0" onclick="geraXLS()" style="cursor:pointer" /></a><span class="aviso">completo</span></td>
                  <td width="5%"><a href="includes/adm_excel.asp?id_rod=<%= id_rod%>&amp;tipo=2" target="_blank"><img src="images/icon_excel.png" alt="Exportar XLS" width="49" height="42" border="0" onclick="geraXLS()" style="cursor:pointer" /><span class="aviso">resumido</span></a></td>
                </tr>
              </table>
                <strong>Data/Hora:</strong> <%= (dta_i &" "& hs_i & " à " & dta_f &" "& hs_f & "<font color=red> - "& dias & " dia(s) de evento.</font>")&" - <b>Horário de almoço</b> ("&hs_ai&" às "&hs_af&")"%><br />
                <strong>Detalhes:</strong> <%= "<font color=red>"& mesas &"</font> Mesa(s)  - <font color=red>"& CInt(reunioes) & "</font> por dia <font color=red>" & Cint(dias)*cint(reunioes) & "</font> Reuniões - <font color=red>"&qtd_empresas&"</font> Emp. confirmada(s) - <font color=red>"& qtd_reunioes&"</font> Reuniões agendadas - <font color=red>"& saldo &"</font> reuniões disponíveis" %><br />
              	<strong>Data Limite:</strong> <%= dados("dta_limite")%><br /><br /></td>
              </tr>
            <tr>
              <td width="314" height="30" bgcolor="#CCCCCC" style="text-align:center"><a href="indicados.asp?id=<%=id_rod%>&filtro=empresa&ordem=<%=ordem%>">EMPRESA</a><%if request("filtro") = "empresa" then %> <img src="<%=img%>" width="10" height="10" alt="Crescente" border="0" /> <% end if %> </td>
              <td width="82" height="30" bgcolor="#CCCCCC" align="center"><a href="indicados.asp?id=<%=id_rod%>&filtro=nivel&ordem=<%=ordem%>">NÍVEL </a><%if request("filtro") = "nivel" then %> <img src="<%=img%>" width="10" height="10" alt="Crescente" border="0" /> <% end if %> </td>
              <td width="438" height="30" align="center" bgcolor="#CCCCCC"><a href="indicados.asp?id=<%=id_rod%>&filtro=qtd_indicado&ordem=<%=ordem%>"> QTD INDICADAS</a><%if request("filtro") = "qtd_indicado" then %> <img src="<%=img%>" width="10" height="10" alt="Crescente" border="0" /> <% end if %> </td>
              <td width="90" align="center" bgcolor="#CCCCCC"><a href="indicados.asp?id=<%=id_rod%>&filtro=dta_indicado&amp;ordem=<%=ordem%>"> DATA INDICAÇÃO</a>
                <%if request("filtro") = "dta_indicado" then %>
                <img src="<%=img%>" width="10" height="10" alt="Crescente" border="0" />
                <% end if %></td>
              <%if session("lv_user") = 1 or session("lv_user") = 2 then%>
              <%end if%>
              </tr>
            <%		
		
		y = 0
		x=0
		cont = 0
		while not dados.eof
		

		if y <> dados("id_usu") then
		cont = cont + 1
		
		'response.Write(cont & " - ")



			indicada = ""
			i = 1
			y = dados("id_usu")
			empresa = "<strong>Razão: </strong><font color=red>" & dados("empresa")
			fantasia = "<strong>Fantasia: </strong> " & dados("emp_fantasia")
			cnpj = "<strong>CNPJ: </strong> " & dados("cnpj")
			nivel = dados("nivel")
			participante = "<strong>Participante: </strong>" & dados("nome")
			cpf = "<strong>CPF: </strong>" & dados("cpf")
			email = "<strong>E-mail: </strong>" & dados("email")
			telefone = "<strong>Telefone: </strong>" & dados("telefone")
			celular = "<strong>Celular: </strong>" & dados("celular")
			dias = dados("dias")
			ancora = dados("ancora")
			if nivel = "Vendedor" then
				nivel_indicado = 3
			else
				nivel_indicado = 4
			end if
			select case nivel_indicado
				case 3
					n_nivel = 4
				case 4
					n_nivel = 3
			end select

			indicada = "<span align='center'><h2>"&dados("qtd_indicado")&"</h2></span>" & i &" - <a href='reunioes.asp?id="&id_rod&"&id_ind="&dados("id_indicado")&"&ancora=3'><img src='images/icon_print2.png' width='20' height='20' border='0' />&nbsp;"& left(dados("emp_indicada"),59) & "</a> - Data reunião: " & dados("hr")
			dta_indicado = "<span align='center'><h2>&nbsp;</h2></span>" & FormatDateTime(dados("dta_indicado"),2)
			
%>




            <tr id="linha<%=x%>" onMouseOver="Mudacor(0,'linha<%=x%>')" onMouseOut="Mudacor(1,'linha<%=x%>')">
              <td width="314" height="20" class="borda">
			    <%= "<span style='font-size:16px; font-weight:bold; color:#FF0000'>"&cont&" - </span>"& y & " - "& empresa &"</font><br>"& fantasia &"<br>"& cnpj &"<br><br>"& participante &"<br>"& cpf &"<br>"& email &"<br>"& telefone &"<br>"& celular &"<br><br>" %><br />
<br />
<br />
<a href="reunioes.asp?id=<%=id_rod%>&id_ind=<%=dados("id_usu")%>&ancora=4"><img src="images/icon_print2.png" width="20" height="20" border="0" /></a>&nbsp;<a href="projetos_indicacao.asp?id=<%=id_rod%>&id_usu=<%=dados("id_usu")%>"><img src="images/icon_agenda3.png" width="20" height="20" border="0" /></a></td>
              <td width="82" height="20" align="center" class="borda"><%= nivel%></td>
<%
		else
			i = i + 1
			select case nivel_indicado
				case 3
					n_nivel = 4
				case 4
					n_nivel = 3
			end select
			
			indicada =  indicada & "<br>" & i &" - <a href='reunioes.asp?id="&id_rod&"&id_ind="&dados("id_indicado")&"&ancora=3'><img src='images/icon_print2.png' width='20' height='20' border='0' />&nbsp;"& left(dados("emp_indicada"),59) & "</a> - Data reunião: " & dados("hr")
			dta_indicado = dta_indicado & "<br>" & FormatDateTime(dados("dta_indicado"),2)
		end if
		
		
		
		
	%>  
              
              
              <%
			  x=x+1
			  dados.movenext
		
		if not dados.eof then
			if y <> dados("id_usu")  then	%>
				  <td height="20" class="borda"><%= indicada%> </td>
				  <td height="20" class="borda" align="center"><%= dta_indicado%> </td>
				  </tr>	
			
		<%  end if 	
		else
		%>
              <td height="20" class="borda"><%= indicada%> </td>
              <td height="20" class="borda"><%= dta_indicado%> </td>
              </tr>	    
	
	<%      
		end if
		wend
		set dados = nothing
	%>
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

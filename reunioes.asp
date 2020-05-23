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

function detalhes(id){
	window.open("detalhes.asp?id="+id+"",'','width=480,height=200,status=yes,toolbar=no,scrollbars=yes,resizable=yes,navbar=no');
}


</script>
<%
if session("user_name") = "" then
	response.Redirect("index.asp")
	response.End()
end if

if request("id") = "" then
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
	filtro = "hr"
else
	filtro = request.QueryString("filtro")
end if


busca = ""
select case session("lv_user")
	case 1, 2, 3, 4
		select case int(request("ancora"))
			case 3
				sql = ("select I.id_ind, I.id_rod, (select count(distinct S2.id_usu) FROM conceitobrazil.tb_indicacao as S2 where I.id_rod = S2.id_rod) as qtd_empresas, (select count(id_usu) from conceitobrazil.tb_indicacao) as qtd_reunioes, R.dta_ini, R.dta_fim, R.mesas, R.almoco, R.tempo, R.intervalo, R.local, R.assunto, R.dta_limite, R.manual, R.banner, I.id_usu, case L.nivel when 1 then 'Administrador' when 2 then 'Gerente' when 3 then 'Vendedor' when 4 then 'Comprador' END as Nivel , L.empresa, L.cnpj, L.emp_fantasia, L.nome, L.cpf, L.email, L.telefone, L.celular, L.perfil, (select count(Si.id_usu) from conceitobrazil.tb_indicacao as Si where Si.id_usu = I.id_usu ) as qtd_indicado, I.dta_indicado, I.dias, I.id_indicado, (select Di.empresa from conceitobrazil.tb_login as Di where I.id_indicado = Di.id_usu) as emp_indicada, (select Di.nome from conceitobrazil.tb_login as Di where I.id_indicado = Di.id_usu) as nome_indicada, (select Di.emp_fantasia from conceitobrazil.tb_login as Di where I.id_indicado = Di.id_usu) as fantasia_indicada, (select Di.cnpj from conceitobrazil.tb_login as Di where I.id_indicado = Di.id_usu) as cnpj_indicada, (select if (Di.nivel = 3, 'Vendedor', 'Comprador') from conceitobrazil.tb_login as Di where I.id_indicado = Di.id_usu) as nivel_indicado, (select Di.cpf from conceitobrazil.tb_login as Di where I.id_indicado = Di.id_usu) as cpf_indicada, (select Di.email from conceitobrazil.tb_login as Di where I.id_indicado = Di.id_usu) as email_indicada, (select Di.telefone from conceitobrazil.tb_login as Di where I.id_indicado = Di.id_usu) as telefone_indicada, (select Di.celular from conceitobrazil.tb_login as Di where I.id_indicado = Di.id_usu) as celular_indicada, (select Di.perfil from conceitobrazil.tb_login as Di where I.id_indicado = Di.id_usu) as perfil_indicada, I.hr, L.idioma, (select Di.idioma from conceitobrazil.tb_login as Di where I.id_indicado = Di.id_usu) as idioma_indicada FROM conceitobrazil.tb_indicacao as I inner join conceitobrazil.tb_login as L on I.id_usu = L.id_usu inner join conceitobrazil.tb_rodada as R on R.id_rod = I.id_rod where R.id_rod="&id_rod&" and I.id_indicado = "&request("id_ind")&" group by I.id_usu, I.id_ind order by hr asc")
			case 4
			sql = ("select I.id_ind, I.id_rod, (select count(distinct S2.id_usu) FROM conceitobrazil.tb_indicacao as S2 where I.id_rod = S2.id_rod) as qtd_empresas, (select count(id_usu) from conceitobrazil.tb_indicacao) as qtd_reunioes, R.dta_ini, R.dta_fim, R.mesas, R.almoco, R.tempo, R.intervalo, R.local, R.assunto, R.dta_limite, R.manual, R.banner, I.id_usu, case L.nivel when 1 then 'Administrador' when 2 then 'Gerente' when 3 then 'Vendedor' when 4 then 'Comprador' END as Nivel , L.empresa, L.cnpj, L.emp_fantasia, L.nome, L.cpf, L.email, L.telefone, L.celular, L.perfil, (select count(Si.id_usu) from conceitobrazil.tb_indicacao as Si where Si.id_usu = I.id_usu ) as qtd_indicado, I.dta_indicado, I.dias, I.id_indicado, (select Di.emp_fantasia from conceitobrazil.tb_login as Di where I.id_indicado = Di.id_usu) as emp_indicada, I.hr, L.idioma FROM conceitobrazil.tb_indicacao as I inner join conceitobrazil.tb_login as L on I.id_usu = L.id_usu inner join conceitobrazil.tb_rodada as R on R.id_rod = I.id_rod where R.id_rod="&id_rod&" and I.id_usu = "&request("id_ind")&" group by I.id_usu, I.id_ind order by "&filtro&" "&ordem&"")
		end select
end select



'response.Write(sql)

set dados = Server.CreateObject("ADODB.Recordset")
dados.ActiveConnection = adoConn
dados.Open(sql)

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

select case int(request("ancora"))
	case 3
		idioma = dados("idioma_indicada")
	case 4
		idioma = dados("idioma")
end select


select case idioma
	case "English"
		var1 = "Indications: "
		var2 = "Date/Time: "
		var11 = " day(s) of the event."
		var3 = "Lunch time: "
		var4 = "Company: "
		var5 = "Fantasy company: "
		var6 = "Participant: "
		var7 = "Phone: "
		var8 = "Mobile: "
		var9 = "DATE"
		var10 = "MEETING"
	case "Spanish"
		var1 = "Indicaciones: "
		var2 = "Fecha/Hora: "
		var11 = " día(s) del evento."
		var3 = "La hora del almuerzo: "
		var4 = "Empresa: "
		var5 = "Fantasia: "
		var6 = "Partícipe: "
		var7 = "Teléfono: "
		var8 = "Celular: "
		var9 = "DATA "
		var10 = "REUNIÓN"
	case else 'Portugues
		var1 = "Indicações: "
		var2 = "Data/Hora: "
		var11 = " dia(s) de evento."
		var3 = "Horário de almoço: "
		var4 = "Razão: "
		var5 = "Fantasia: "
		var6 = "Participante: "
		var7 = "Telefone: "
		var8 = "Celular: "
		var9 = "DATA"
		var10 = "REUNIÃO"
end select

	
select case session("lv_user")
	case 1, 2, 3, 4
		select case int(request("ancora"))
			case 3
				empresa = "<strong>"&var4&"</strong><font color=red>" & dados("emp_indicada")
				fantasia = "<strong>"&var5&"</strong> " & dados("fantasia_indicada")
				cnpj = "<strong>CNPJ: </strong> " & dados("cnpj_indicada")
				nivel = dados("nivel_indicado")
				participante = "<strong>"&var6&"</strong>" & dados("nome_indicada")
				cpf = "<strong>CPF: </strong>" & dados("cpf_indicada")
				email = "<strong>E-mail: </strong>" & dados("email_indicada")
				telefone = "<strong>"&var7&"</strong>" & dados("telefone_indicada")
				celular = "<strong>"&var8&"</strong>" & dados("celular_indicada")
				indicada = i &" - "& left(dados("empresa"),59) 
			case 4
				empresa = "<strong>"&var4&"</strong><font color=red>" & dados("empresa")
				fantasia = "<strong>"&var5&"</strong> " & dados("emp_fantasia")
				cnpj = "<strong>CNPJ: </strong> " & dados("cnpj")
				nivel = dados("nivel")
				participante = "<strong>"&var6&"</strong>" & dados("nome")
				cpf = "<strong>CPF: </strong>" & dados("cpf")
				email = "<strong>E-mail: </strong>" & dados("email")
				telefone = "<strong>"&var7&"</strong>" & dados("telefone")
				celular = "<strong>"&var8&"</strong>" & dados("celular")
				indicada = i &" - "& left(dados("emp_indicada"),59) 
		end select
end select

		'dias = dados("dias")
		indicada = i &" - "& left(dados("emp_indicada"),59) 
		dta_indicado = dados("hr")
		banner = dados("banner")

	 %>
          <table width="950" border="0" align="center" cellpadding="2" cellspacing="2">
            <tr>
              <td colspan="5"><table width="100%" border="0" cellspacing="0" cellpadding="0">
                <tr>
                  <td scope="row"><h1><% if request("ancora") <> "" then response.Write("<a href='javascript:void(0);history.back(-1);'><img src='images/icon_voltar.png' width='27' height='30' /></a> ") & var1 & dados("local") & " - " & dados("assunto")%></h1></td>
                  <td scope="row"><a href="javascript:void(0)" onclick="window.print();"><img src="images/icon_print.png" alt="Imprimir" width="40" height="40" border="0" /></a></td>
                  </tr>
              </table>
                <strong><%=var2%></strong> <%= (dta_i &" "& hs_i & " à " & dta_f &" "& hs_f & "<font color=red> - "& dias & var11 &"</font>")&" - <b>"&var3&"</b> ("&hs_ai&" às "&hs_af&")"%><br />
                
   			  	<%response.Write(empresa &"</font><br>"& fantasia &"<br>"& cnpj &"<br><br>"& participante &"<br>"& cpf &"<br>"& email &"<br>"& telefone &"<br>"& celular &"<br><br>")%>

                </td>
              </tr>
            <tr>
              <td width="254" height="30" bgcolor="#CCCCCC" align="center"><a href="reunioes.asp?id=<%=id_rod%>&filtro=hr&ordem=<%=ordem%>"><%=var9%>  </a><%if request("filtro") = "hr" then %> <img src="<%=img%>" width="10" height="10" alt="Crescente" border="0" /> <% end if %> </td>
              <td width="682" height="30" align="center" bgcolor="#CCCCCC"><a href="reunioes.asp?id=<%=id_rod%>&filtro=emp_indicada&ordem=<%=ordem%>"><%=var10%></a><%if request("filtro") = "emp_indicada" then %> <img src="<%=img%>" width="10" height="10" alt="Crescente" border="0" /> <% end if %> </td>
              </tr>
            <%		
		x=1
		
		while not dados.eof
%>
            <tr id="linha<%=x%>" onMouseOver="Mudacor(0,'linha<%=x%>')" onMouseOut="Mudacor(1,'linha<%=x%>')">
              <td width="254" height="20" align="center" class="borda"><%= dados("hr")%></td>
              <td width="682" height="20" align="center" class="borda">
			  	<%
					select case session("lv_user")
						case 1, 2, 3, 4
							select case int(request("ancora"))
								case 3%>
									<a href="javascript:void(0)" onClick="javascript:detalhes(<%=ucase(dados("id_usu"))%>);"><%=ucase(dados("emp_fantasia"))%></a>
                                    <%
								case 4 									
									response.Write(ucase(dados("emp_indicada")))
							end select
					end select
				%>
              </td>
             </tr>
<%

	  x=x+1
	  dados.movenext
		wend
		set dados = nothing
	%>
        </table>
          <% else 
		response.Write("Registro não encontrado")
		end if


	if banner <> "" then response.Write("<img src='arquivos/"&banner&"' border='0' />")
	
	%>
          
          </td>
      </tr>
    </table></td>
  </tr>
</table>
<!-- #Include File="rodape.asp" --> 
</body>
</html>

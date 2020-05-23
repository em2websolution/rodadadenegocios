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

if request("id") = "" then
	response.Redirect("projetos.asp")
	response.End()
else
	id_rod = request("id")
end if

if request("id_ancora") = "" then
	sql = ("select I.id_sel, I.id_rod, R.dta_ini, R.dta_fim, R.mesas, R.almoco, R.tempo, R.intervalo, R.local, R.assunto, R.dta_limite, R.manual, I.id_usu, case L.nivel when 1 then 'Administrador' when 2 then 'Gerente' when 3 then 'Vendedor' when 4 then 'Comprador' END as Nivel , R.Ancora, (select Di.emp_fantasia from conceitobrazil.tb_login as Di where I.id_usu = Di.id_usu) as emp_indicada FROM conceitobrazil.tb_selecao as I inner join conceitobrazil.tb_login as L on I.id_usu = L.id_usu inner join conceitobrazil.tb_rodada as R on R.id_rod = I.id_rod where R.id_rod="&id_rod&" and L.nivel in (3,4) group by emp_indicada order by emp_indicada")
else
	sql = ("select I.id_sel, I.id_rod, R.dta_ini, R.dta_fim, R.mesas, R.almoco, R.tempo, R.intervalo, R.local, R.assunto, R.dta_limite, R.manual, I.id_usu, case L.nivel when 1 then 'Administrador' when 2 then 'Gerente' when 3 then 'Vendedor' when 4 then 'Comprador' END as Nivel , R.Ancora, (select Di.emp_fantasia from conceitobrazil.tb_login as Di where I.id_usu = Di.id_usu) as emp_indicada FROM conceitobrazil.tb_selecao as I inner join conceitobrazil.tb_login as L on I.id_usu = L.id_usu inner join conceitobrazil.tb_rodada as R on R.id_rod = I.id_rod where R.id_rod="&id_rod&" and L.nivel = "&request("id_ancora")&" group by emp_indicada order by emp_indicada")
end if


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
	

		set objRecordset = Server.CreateObject("adodb.recordset")
		with objRecordset 
			.Open sql,adoConn
			
			if not .eof then
				MyList = .GetRows()
			end if
			.Close
		end with
		set objRecordset = nothing

		intLin = cint(ubound(MyList,2))+1
		boolMountTable = true
		varTotalRecords = intLin 
		iList = 0

		'for iList = 0 to varTotalRecords-1
			'response.Write(MyList(29,iList) &"<br>")
		
		
			if request("id_ancora") = "" then
				ancora = MyList(14,iList)
			else
					ancora = request("id_ancora")
			end if
		
		dias = DateDiff("d",MyList(2,iList),MyList(3,iList)+1)
		'next
		'response.End()
		
		hs1=hour(MyList(2,iList))&":"&minute(MyList(2,iList))
		hs2=hour(MyList(3,iList))&":"&minute(MyList(3,iList))
		minutos = DateDiff("n",hs1,hs2)
		mesas = MyList(4,iList)
		almoco = Split(MyList(5,iList),"-")
		i = 1
		for each x in almoco
		    if i = 1 then hs_ai = x
		    if i = 2 then hs_af = x
			i = i+1
		next
		
		min_a = DateDiff("n",hs_ai,hs_af)		
		
		tempo = MyList(6,iList)
		intervalo = MyList(7,iList)
		
		duracao = (tempo+intervalo)
		
		'Calculo
		'response.Write("(tempo dia: " & minutos & " - tempo almoco: " & min_a & ") / (tempo reuniao: " & tempo & "+ tempo intervalo: " & intervalo & ")" )
		
		reunioes = int((minutos-min_a))/(duracao) * mesas
		
		qtd_empresas = varTotalRecords
		'qtd_reunioes = cint(MyList(3,iList))
		'saldo = ((dias*reunioes)-qtd_reunioes)
		
		select case tempo
			case "60"
				tempo = "1hora"
			case else
				tempo = MyList(6,iList) & "min"
		end select
		
	dta_i = day(MyList(2,iList)) &"/"& month(MyList(2,iList)) &"/"& year(MyList(2,iList))
	dta_f = day(MyList(3,iList)) &"/"& month(MyList(3,iList)) &"/"& year(MyList(3,iList))
	if hour(MyList(2,iList)) < 10 then 
		h1 = "0" & hour(MyList(2,iList))
	else
		h1 = hour(MyList(2,iList))
	end if

	if minute(MyList(2,iList)) < 10 then 
		m1 = "0" & minute(MyList(2,iList))
	else
		m1 = minute(MyList(2,iList))
	end if

	if hour(MyList(3,iList)) < 10 then 
		h2 = "0" & hour(MyList(3,iList))
	else
		h2 = hour(MyList(3,iList))
	end if

	if minute(MyList(3,iList)) < 10 then 
		m2 = "0" & minute(MyList(3,iList))
	else
		m2 = minute(MyList(3,iList))
	end if

	
	hs_i = h1 &":"& m1
	hs_f = h2 &":"& m2

	manual = MyList(11,iList)
	 %>
          <table width="950" border="0" align="center" cellpadding="2" cellspacing="2">
            <tr>
              <td colspan="5">
								<table width="100%" border="0" cellspacing="0" cellpadding="0">
                <tr>
                  <td width="72%" scope="row"><h1>Agenda Rodada: <%=MyList(8,iList) & " - " & MyList(9,iList)%></h1></td>
                  <th width="23%"><div id="icon"></div>&nbsp;</th>
                  <td width="5%"><a href="mapa_excel.asp?id=<%= id_rod%>" target="_blank"><img src="images/icon_excel.png" alt="Exportar XLS" width="49" height="42" border="0" onclick="geraXLS()" style="cursor:pointer" /></a></td>
                </tr>
              	</table>
                <strong>Data/Hora:</strong> <%= (dta_i &" "& hs_i & " à " & dta_f &" "& hs_f & "<font color=red> - "& dias & " dia(s) de evento.</font>")&" - <b>Horário de almoço</b> ("&hs_ai&" às "&hs_af&")"%><br />
                <strong>Detalhes:</strong> <%= "<font color=red>"& mesas &"</font> Mesa(s)  - <font color=red>"& CInt(reunioes) & "</font> por dia <font color=red>" & Cint(dias)*cint(reunioes) & "</font> Reuniões - <font color=red>"&qtd_empresas&"</font> Emp. confirmada(s) - <font color=red>"& qtd_reunioes&"</font> Reuniões agendadas - <font color=red>"& saldo &"</font> reuniões disponíveis" %><br />
              	<strong>Data Limite:</strong> <%= MyList(10,iList)%><br />
								<strong>Ancora do Projeto:</strong> <% if MyList(14,iList) = 3 then response.write("Vendedor") else response.write("Comprador") %> <br>
								<strong>Mostrar Agenda dos Ancoras</strong> <a href="mapa.asp?id=<%= id_rod%>&id_ancora=4" >Comprador</a> - <a href="mapa.asp?id=<%= id_rod%>&id_ancora=3" >Vendedor</a>
								<br /><br />
							</td>
						</tr>
            <tr>
              <td height="30" colspan="5">
                  <table width="942" border="0" cellspacing="4" cellpadding="0">

										<%
										if request("id_ancora") <> "" then
											if ancora = 3 then response.write("<tr><td><h1>Visão Vendedor</h1></td></tr><br>") else response.write("<tr><td><h1>Visão Comprador</h1></td></tr><br>")
										end if
										for x=0 to dias-1
                        dia_reuniao = DateAdd("d", x, dta_i)
						iList = 0
						 
    %>
                        <tr><th colspan="4" scope="row"><h1><%=dia_reuniao%></h1></th></tr>
    
    <%
                        'response.Write(dia_reuniao & "<br>")
                        
				 		
						quebra = 1
						dia_reuniao = year(dia_reuniao) &"-"& month(dia_reuniao) &"-"& day(dia_reuniao)

						for z=1 to qtd_empresas
                            
							if quebra = 1 then
							%> 
								<tr>
						 <% end if  %>
									<td width="226" scope="row" ><%= "<b>" & MyList(15,iList) & "</b><br>"%>
							   <% 'response.Write("Mesa " & z & "<br>")
							   		
                                   	o = 1
									for y=0 to minutos
                                        hora_reuniao = DateAdd("n", y, hs_i)
                                        if hour(hora_reuniao) = hour(hs_ai) then
                                            y = y + min_a -1
												response.Write(hora_reuniao & " - Almoço até "&hs_af&" - Indisponível<br> ")
                                        else
												response.Write(hora_reuniao & " - ")
												y = y + duracao-1
												
												'Verifica se já não tem o horario gravado
												set verifica = adoConn.execute("select I.stu, (select Di.emp_fantasia from conceitobrazil.tb_login as Di where I.id_usu = Di.id_usu) as empresa from conceitobrazil.tb_indicacao as I where hr = '"&dia_reuniao &" "& hora_reuniao&"' and id_indicado = "&MyList(12,iList)&" and id_rod = "&id_rod&"")
												
											'response.Write("select I.stu, (select Di.emp_fantasia from conceitobrazil.tb_login as Di where I.id_usu = Di.id_usu) as empresa from conceitobrazil.tb_indicacao as I where hr = '"&dia_reuniao &" "& hora_reuniao&"' and id_indicado = "&MyList(12,iList)&" and id_rod = "&id_rod&"")
												
												if not verifica.eof then
													nome = verifica("empresa")
													if verifica("stu") = 0 then
														if nome <> "" then
															response.Write("<font color='red'>" & nome & "</font><br>")
														else
															response.Write("<font color='red'>Indisponível</font><br>")
														end if
													else
														response.Write("Disponível" & "<br>")
													end if
													nome = ""
												else
													response.Write("Disponível"& "<br>")
												end if  
												set verifica = nothing
										  end if
										  o = o + 1
                                    next
									
									iList = iList + 1
							%>
									</td>
							<%
                            if quebra = 4 or z = mesas then
								quebra = 0
								%> 
                                </tr>
                                <tr><td width="226" scope="row" class="borda" colspan="4">&nbsp;</td></tr>
						 <% end if 
								'response.Write("<BR>")
						quebra = quebra + 1
						
						next
					next
                  
				  
				  'for i = 1  to mesas
					'response.Write(MeuArray(i,0) &" - "& MeuArray(i,1) &" - "& MeuArray(i,2) &" - "& MeuArray(i,3) &" - "& MeuArray(i,4) &"<br>" )
				   'next
				  
				  %>
                  </table>   
              </td>
              </tr>
        </table>
          <% else 
		response.Write("Registro não encontrado")
		

		end if

				adoConn.close
		set adoConn = nothing
		
	%>
          </td>
      </tr>
    </table></td>
  </tr>
</table>
<!-- #Include File="rodape.asp" --> 
</body>
</html>

<!-- #Include File="top.asp" -->
<script language="javascript" type="text/javascript">
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

</script>
<%
if session("user_name") = "" then
	response.Redirect("index.asp")
	response.End()
end if

if request.QueryString("c") = 1 then
%>
<script language="javascript" type="text/javascript">alert('Cadastro na rodada com sucesso participação efetivada');</script>
<%
end if

if request("id_cad") = "" then
	response.Redirect("index.asp")
else
	id_usu = request("id_cad")
	sql = ("SELECT S.*, L.*, R.*, (SELECT count(*) FROM conceitobrazil.tb_selecao as Sub inner join conceitobrazil.tb_indicacao as I on Sub.id_usu = I.id_usu where Sub.id_usu = "&id_usu&" and Sub.id_usu <> '') as qtd_indicado FROM conceitobrazil.tb_selecao as S inner join conceitobrazil.tb_login as L on S.id_usu = L.id_usu inner join conceitobrazil.tb_rodada as R on R.id_rod = S.id_rod where S.id_usu = "&id_usu&" ORDER BY dta_ini")
	set dados = Server.CreateObject("ADODB.Recordset")
	dados.ActiveConnection = adoConn
	dados.Open(sql)

'response.Write(sql)

if not dados.eof then

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
        <td scope="row" style="text-align:left"><h1>      
		<%if session("lv_user") = 1 or session("lv_user") = 2 then
	  			response.Write("<h1><a href='cadastros.asp'><img src='images/icon_voltar.png' width='27' height='30' /></a>&nbsp;Projeto Empresa: " & dados("empresa") & "</h1><br>Usuário: <font color=red>"&dados("nome") &" - "& dados("email")&"</font:")
			else
	  			response.Write("<h1>Projetos</h1>")
		end if
	  %>
</h1></td>
      </tr>
      <tr>
        <td scope="row">
    <% if not dados.eof then %>
        <table width="950" border="0" align="center" cellpadding="2" cellspacing="0">
      <tr>
        <td colspan="8" style="text-align:center">&nbsp;</td>
        </tr>
      <tr>
        <td width="106" height="30" bgcolor="#CCCCCC" style="text-align:center"><a href="projetos.asp?filtro=dta_ini&ordem=<%=ordem%>">DATA/HORA</a><%if request("filtro") = "dta_ini" then %> <img src="<%=img%>" width="10" height="10" alt="Crescente" border="0" /> <% end if %> </td>
        <td width="109" height="30" bgcolor="#CCCCCC" align="center"><a href="projetos.asp?filtro=local&ordem=<%=ordem%>">LOCAL </a><%if request("filtro") = "local" then %> <img src="<%=img%>" width="10" height="10" alt="Crescente" border="0" /> <% end if %> </td>
        <td width="97" height="30" bgcolor="#CCCCCC" align="center"><a href="projetos.asp?filtro=assunto&ordem=<%=ordem%>">ASSUNTO</a><%if request("filtro") = "assunto" then %> <img src="<%=img%>" width="10" height="10" alt="Crescente" border="0" /> <% end if %> </td>
        <td width="58" bgcolor="#CCCCCC" align="center"><a href="projetos.asp?filtro=tempo&ordem=<%=ordem%>"> REUNIÃO</a><%if request("filtro") = "tempo" then %> <img src="<%=img%>" width="10" height="10" alt="Crescente" border="0" /> <% end if %> </td>
        <td width="66" bgcolor="#CCCCCC" align="center"><a href="projetos.asp?filtro=intervalo&ordem=<%=ordem%>">INTERVALO</a><%if request("filtro") = "intervalo" then %> <img src="<%=img%>" width="10" height="10" alt="Crescente" border="0" /> <% end if %> </td>
                <td width="57" bgcolor="#CCCCCC" align="center">PRAZO</td>
        <td width="300" height="30" bgcolor="#CCCCCC" style="text-align:center">&nbsp;</td>
      </tr>
    <%		
		y = 0
		while not dados.eof
		y = y + 1
		
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
	%>  
          <tr id="linha<%=y%>" onMouseOver="Mudacor(0,'linha<%=y%>')" onMouseOut="Mudacor(1,'linha<%=y%>')" class="borda">
            <td width="106" height="20" align="center"><%= (dta_i &" "& hs_i & " <br>à<br> " & dta_f &" "& hs_f) &"<br> <b>Horário de almoço</b> ("&hs_ai&" às "&hs_af&") <br><font color=red>"& dias & " dia(s) de evento.</font>"%> </td>
            <td width="109" height="20" align="center"><%= dados("local")%></td>
            <td width="97" height="20" align="center"><%= dados("assunto")%></td>
            <td width="58" height="20" align="center"><%= tempo%></td>
            <td width="66" height="20" align="center"><%= dados("intervalo")&"min"%></td>
            <td width="57" height="20" align="center"><%= dados("dta_limite")%></td>
            <td width="300" height="20" align="center">
            	
                 
                <%  
						limite = DateDiff("d",now(),dados("dta_limite"))
						if dados("ancora") = dados("nivel") then
							ancora = 3
						else
							ancora = 4
						end if
						if limite > 0 then	
							if cint(dados("qtd_indicado")) = 0  then
				%>
                    			<a href="reunioes.asp?id=<%= dados.fields(1)%>&amp;id_ind=<%=id_usu%>&amp;ancora=<%=ancora%>"><img src="images/icon_print.png" alt="" width="40" height="40" border="0" /> &nbsp;</a><a href="projetos_indicacao.asp?id=<%= dados.fields(1)%>&id_usu=<%=id_usu%>"><img src="images/confirmar.png" alt="Confirmar participação" width="195" height="40" border="0" /></a><% 
							else %>       			  <a href="reunioes.asp?id=<%= dados.fields(1)%>&id_ind=<%=id_usu%>&ancora=4"><img src="images/icon_print.png" width="40" height="40" border="0" /> &nbsp;</a><a href="projetos_indicacao.asp?id=<%= dados.fields(1)%>&id_usu=<%=id_usu%>"><img src="images/alterar.png" alt="Alterar participação" width="195" height="40" border="0" /></a><% 
							end if
						else %>
							<a href="reunioes.asp?id=<%= dados.fields(1)%>&id_ind=<%=id_usu%>&ancora=3"><img src="images/icon_print.png" alt="" width="40" height="40" border="0" /> &nbsp; <img src="images/encerradas.png" alt="Inscrições encerradas!" width="195" height="40" border="0" /></a>
				<%
						end if
				%>
                </td>
          </tr>
    <%
		dados.movenext
		wend
		
		set dados = nothing
	%>
        </table>
    <% else 
		response.Write("Registro não encontrado")
		end if

else
	response.Write("Nenhum projeto selecionado para esse usuário")
end if 
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

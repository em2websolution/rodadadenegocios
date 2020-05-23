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

if request.QueryString("ordem") = "" then
	ordem = "desc"
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
	filtro = "dta_ini"
else
	filtro = request.QueryString("filtro")
end if

if request.Form("str") <> "" then
	select case session("lv_user")
		case 1
			txtStr = "where local like '%"&request.Form("str")&"%'"
		case else
			txtStr = "where S.id_usu = "&session("user_rd")&" and local like '%"&request.Form("str")&"%'"
	end select
else
	select case session("lv_user")
		case 1
			txtStr = ""
		case else
			txtStr = "where S.id_usu = "&session("user_rd")&""
	end select
end if

select case session("lv_user")
	case 1
		sql = ("SELECT distinct R.*, (select count(distinct S2.id_usu) from conceitobrazil.tb_indicacao as S2 where R.id_rod = S2.id_rod )as qtd_empresas, (select count(S3.id_indicado) from conceitobrazil.tb_indicacao as S3 where R.id_rod = S3.id_rod and S3.id_usu <> '') as qtd_indicado FROM conceitobrazil.tb_rodada as R "&txtStr&" group by R.id_rod ORDER BY "&filtro&" "&ordem&"")		
	case 2
		sql = ("SELECT R.*, S.*, L.*, (select count(distinct S2.id_usu) from conceitobrazil.tb_indicacao as S2 where R.id_rod = S2.id_rod )as qtd_empresas, (SELECT count(*) FROM conceitobrazil.tb_selecao as Sub inner join conceitobrazil.tb_indicacao as I on Sub.id_usu = I.id_usu where Sub.id_usu = "&session("user_rd")&" and Sub.id_usu <> '') as qtd_indicado FROM conceitobrazil.tb_selecao as S inner join conceitobrazil.tb_login as L on S.id_usu = L.id_usu inner join conceitobrazil.tb_rodada as R on R.id_rod = S.id_rod "&txtStr&" ORDER BY "&filtro&" "&ordem&"")
	case 3, 4
		sql = ("SELECT distinct R.*, L.nivel, (SELECT count(*) FROM conceitobrazil.tb_selecao as Sub inner join conceitobrazil.tb_indicacao as I on Sub.id_usu = I.id_usu where Sub.id_usu = "&session("user_rd")&" and Sub.id_usu <> '') as qtd_indicado FROM conceitobrazil.tb_selecao as S inner join conceitobrazil.tb_login as L on S.id_usu = L.id_usu inner join conceitobrazil.tb_rodada as R on R.id_rod = S.id_rod "&txtStr&" ORDER BY "&filtro&" "&ordem&"")
end select
set dados = Server.CreateObject("ADODB.Recordset")
dados.ActiveConnection = adoConn
dados.Open(sql)

'response.Write("<br> Projetos: "& sql)
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
        <td scope="row" style="text-align:left"><h1>Projetos</h1></td>
      </tr>
      <tr>
        <td scope="row">
    <% if not dados.eof then %>
        <table width="950" border="0" align="center" cellpadding="2" cellspacing="0">
      <%if session("lv_user") = 1 or session("lv_user") = 2 then%>
      <tr>
        <td colspan="8" class="borda"><strong>BUSCA:</strong></td>
      </tr>
      <tr>
        <td height="30" colspan="8" class="borda" style="text-align:center"><form id="form1" name="form1" method="post" action="">
          <table width="950" border="0" cellspacing="2" cellpadding="0">
            <tr>
              <td width="56"><strong>Local:</strong></td>
              <td width="357"><input name="str" type="text" id="str" value="<%=request.form("str")%>" style="width:350px" /></td>
              <td width="73"><input type="submit" name="button2" id="button2" value="Buscar"/></td>
              <td width="404" align="right"><% if session("lv_user") = 1 then %><a href="#">NOVO</a><% end if %>&nbsp;</td>
              <td width="48" align="center"><% if session("lv_user") = 1 then %><a href="projetos_novo.asp"><img src="images/icon_novo.png" alt="Novo registro" width="28" height="31" border="0" /></a><% end if %>&nbsp;</td>
            </tr>
          </table>
        </form></td>
        </tr>
      <% end if %>
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
		<%if session("lv_user") = 1 or session("lv_user") = 2 then%>
		<td width="125" height="30" align="center" bgcolor="#CCCCCC"><a href="projetos.asp?filtro=reunioes&ordem=<%=ordem%>">INFOS</a><%if request("filtro") = "reunioes" then %> <img src="<%=img%>" width="10" height="10" alt="Crescente" border="0" /> <% end if %> </td><%end if%>
        <td width="300" height="30" bgcolor="#CCCCCC" style="text-align:center">AÇÕES</td>
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
		
		if session("lv_user") = 1 or session("lv_user") = 2 then
			qtd_empresas = cint(dados("qtd_empresas"))
			qtd_indicado = cint(dados("qtd_indicado"))
			saldo = ((dias*reunioes)-qtd_indicado)
		end if
		
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
            <%if session("lv_user") = 1 or session("lv_user") = 2 then%><td width="125" height="20" align="right"><%= "<font color=red>"& mesas &"</font> Mesa(s) <br><font color=red>"& CInt(reunioes) & "</font> por dia <font color=red>" & Cint(dias)*cint(reunioes) & "</font> Reuniões <br> <font color=red>"&qtd_empresas&"</font> Emp. confirmada(s) <br> <font color=red>"& qtd_indicado&"</font> Reuniões agendadas <br> <font color=red>"& saldo &"</font> reuniões disponíveis" %></td><%end if%>
            <td width="300" height="20" align="center">
                 <table width="100%" border="0" cellspacing="4" cellpadding="0" style="font:Verdana, Geneva, sans-serif; font-size:9px;">
                   <tr>
                     <th scope="row" align="center"><% if manual <> "" then %><a href="arquivos/<%= manual%>" target="_blank"><img src="images/pdf.jpg" width="20" height="20" border="0" /></a><% end if %></th>
					 <% if session("lv_user") = 1  then %>
                     <td width="13%" align="center">
                     	<a href="projetos_novo.asp?id=<%= dados.fields(0)%>"><img src="images/icon_alterar.png" alt="Alterar" width="25" height="25" border="0" /><br />alterar </a>                       
                     </td>
                     <% end if %>
					 <% if session("lv_user") = 1 then %> 
                     <td align="center">
                     	<a href="javascript:void(0);" onClick="exclui('','projetos_novo.asp?id=<%= dados.fields(0)%>&empresa=<%=dados("local")%>&nome=<%=dados("assunto")%>&exc=1')"><img src="images/icon_excluir.png" alt="Alterar" width="26" height="25" border="0" /><br />excluir </a>
                     </td>
                     <% end if %>
					 <% if session("lv_user") = 1 then %>
                     <td width="15%" align="center"> 
                     	<a href="projetos_selecao.asp?id=<%= dados.fields(0)%>"><img src="images/icon_sales.png" alt="Vendedor" width="25" height="25" border="0"/><br />seleção </a> 
                     </td>
					 <% end if %>
					 <% if session("lv_user") = 1 then %>
                     <td align="center">
                     	<a href="agenda.asp?id=<%= dados.fields(0)%>"><img src="images/icon_agenda.png" width="25" height="25" border="0"/><br />Mesas </a>                    
                     </td>
				     <% end if %>
					 <% if session("lv_user") = 1 or session("lv_user") = 2 then %>
                     <td align="center">
                     	<a href="indicados.asp?id=<%= dados.fields(0)%>"><img src="images/icon_confirmados.png" width="17" height="25" border="0"/><br />Indicações </a>
                     </td>
					 <% end if %>
					 <% if session("lv_user") = 1 or session("lv_user") = 2 then %>
                     <td align="center">
                     	<a href="mapa.asp?id=<%= dados.fields(0)%>"><img src="images/icon_reunioes.png" width="26" height="25" border="0"/><br />Agenda<br /></a>
                     </td>
					 <% end if %>
                   </tr>
                 </table>
                <%  
					if session("lv_user") = 3 or session("lv_user") = 4 then
						if dados("ancora") = dados("nivel") then
							ancora = 3
						else
							ancora = 4
						end if

						limite = DateDiff("d",now(),dados("dta_limite"))
						if limite > 0 then	
							if cint(dados("qtd_indicado")) = 0  then
				%>
                    			<a href="reunioes.asp?id=<%= dados("id_rod")%>&amp;id_ind=<%=session("user_rd")%>&amp;ancora=<%=ancora%>"><img src="images/icon_print.png" alt="" width="40" height="40" border="0" /> &nbsp;</a><a href="projetos_indicacao.asp?id=<%= dados("id_rod")%>"><img src="images/confirmar.png" alt="Confirmar participação" width="195" height="40" border="0" /></a><% 
							else %>       			  <a href="reunioes.asp?id=<%= dados("id_rod")%>&id_ind=<%=session("user_rd")%>&ancora=<%=ancora%>"><img src="images/icon_print.png" width="40" height="40" border="0" /> &nbsp;</a><a href="projetos_indicacao.asp?id=<%= dados("id_rod")%>"><img src="images/alterar.png" alt="Alterar participação" width="195" height="40" border="0" /></a>
							<% 
							end if
						else %>
							<a href="reunioes.asp?id=<%= dados("id_rod")%>&id_ind=<%=session("user_rd")%>&ancora=<%=ancora%>"><img src="images/icon_print.png" alt="" width="40" height="40" border="0" /> &nbsp; <img src="images/encerradas.png" alt="Inscrições encerradas!" width="195" height="40" border="0" /></a>
				<%
                		end if
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
	%>
    </td>
      </tr>
    </table></td>
  </tr>
</table>
<!-- #Include File="rodape.asp" --> 
</body>
</html>

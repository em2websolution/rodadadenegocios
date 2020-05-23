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

	function upload(x){
		window.open('includes/uploadthumbnail.asp?tipo='+x+'','upload','width=480,height=200,status=yes,toolbar=no,scrollbars=yes,resizable=yes,navbar=no');
	}

</script>
<%
if session("user_name") = "" then
	response.Redirect("index.asp")
	response.End()
end if

if session("lv_user") = 3 or session("lv_user") = 4 then 
	response.Redirect("index.asp")
	response.End()
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
	filtro = "email"
else
	filtro = request.QueryString("filtro")
end if


busca = ""

		
if trim(request("str")) <> "" then
	str = "empresa like '%"&request("str")&"%' or email like '%"&request("str")&"%' or nome like '%"&request("str")&"%' "
	if request("polo") <> "" then
		str = str & "and polo = '"&request("polo")&"'"
	end if
	sql = "SELECT * FROM tb_login where nivel in (3,4) and stu=1 and "&str&" order by "&filtro&" "&ordem&""	
	Set dados  = adoConn.execute(sql)
else
	str = str & "and polo = '"&request("polo")&"'"
	sql = "SELECT * FROM tb_login where nivel in (3,4) and stu=1 "&str&" order by "&filtro&" "&ordem&""	
	Set dados  = adoConn.execute(sql)
end if


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
        <td scope="row" style="text-align:left"><h1>Cadastros</h1></td>
      </tr>
      <tr>
        <td scope="row"><form id="form1" name="form1" method="post" action="cadastros.asp">
          <table width="950" border="0" cellspacing="2" cellpadding="0">
            <tr>
              <td width="82"><strong>Informação:</strong></td>
              <td width="351"><input name="str" type="text" id="str" value="<%=request("str")%>" style="width:350px" /></td>
              <td width="59">&nbsp;</td>
              <td width="57" align="right">&nbsp;</td>
              <% if session("lv_user") = 1  then %>
              <td width="156" align="right">Importar</td>
              <td width="49" align="right"><img src="images/icon_excel.png" alt="Exportar XLS" width="49" height="42" border="0" onclick="javascript:upload(3);" style="cursor:pointer" /></td>
              <td width="141" align="right"><a href="#">NOVO</a></td>
              <td width="37" align="center"><a href="cadastros_novo.asp"><img src="images/icon_novo.png" alt="Novo registro" width="28" height="31" border="0" /></a></td>
            	<% end if %>
            </tr>
            <tr>
              <td><span class="vermelho"><strong>Segmento</strong>:</span></td>
              <td><span class="vermelho">
                <select name="polo" size="1" id="polo"  tabindex="27">
                  <option value="" selected="selected">Vazio</option>
                  <option value="Agroindústria em geral" <% if trim(request("polo")) = "Agroindústria em geral" then response.Write("selected")%>>Agroindústria em geral</option>
                  <option value="Alimentos" <% if trim(request("polo")) = "Alimentos" then response.Write("selected")%>>Alimentos</option>
                  <option value="Apicultura" <% if trim(request("polo")) = "Apicultura" then response.Write("selected")%>>Apicultura</option>
                  <option value="Aquicultura" <% if trim(request("polo")) = "Aquicultura" then response.Write("selected")%>>Aquicultura</option>
                  <option value="Arquitetura" <% if trim(request("polo")) = "Arquitetura" then response.Write("selected")%>>Arquitetura</option>
                  <option value="Artesanato" <% if trim(request("polo")) = "Artesanato" then response.Write("selected")%>>Artesanato</option>
                  <option value="Bebidas" <% if trim(request("polo")) = "Bebidas" then response.Write("selected")%>>Bebidas</option>
                  <option value="Calçados" <% if trim(request("polo")) = "Calçados" then response.Write("selected")%>>Calçados</option>
                  <option value="Cerâmica" <% if trim(request("polo")) = "Cerâmica" then response.Write("selected")%>>Cerâmica</option>
                  <option value="Confecções" <% if trim(request("polo")) = "Confecções" then response.Write("selected")%>>Confecções</option>
                  <option value="Eletro Eletrônico" <% if trim(request("polo")) = "Eletro Eletrônico" then response.Write("selected")%>>Eletro Eletrônico</option>
                  <option value="Gastronomia" <% if trim(request("polo")) = "Gastronomia" then response.Write("selected")%>>Gastronomia</option>
                  <option value="Gemas, Jóias e Assessórios" <% if trim(request("polo")) = "Gemas, Jóias e Assessórios" then response.Write("selected")%>>Gemas, Jóias e Assessórios</option>
                  <option value="Higiene Pessoal, Cosméticos e Perfumaria" <% if trim(request("polo")) = "Higiene Pessoal, Cosméticos e Perfumaria" then response.Write("selected")%>>Higiene Pessoal, Cosméticos e Perfumaria</option>
                  <option value="Investimento Imobiliário" <% if trim(request("polo")) = "Investimento Imobiliário" then response.Write("selected")%>>Investimento Imobiliário</option>
                  <option value="Leite e Derivados" <% if trim(request("polo")) = "Leite e Derivados" then response.Write("selected")%>>Leite e Derivados</option>
                  <option value="Madeireiro" <% if trim(request("polo")) = "Madeireiro" then response.Write("selected")%>>Madeireiro</option>
                  <option value="Malacocultura" <% if trim(request("polo")) = "Malacocultura" then response.Write("selected")%>>Malacocultura</option>
                  <option value="Metalmecânico" <% if trim(request("polo")) = "Metalmecânico" then response.Write("selected")%>>Metalmecânico</option>
                  <option value="Móveis" <% if trim(request("polo")) = "Móveis" then response.Write("selected")%>>Móveis</option>
                  <option value="Náutico" <% if trim(request("polo")) = "Náutico" then response.Write("selected")%>>Náutico</option>
                  <option value="Orgânicos" <% if trim(request("polo")) = "Orgânicos" then response.Write("selected")%>>Orgânicos</option>
                  <option value="TI" <% if trim(request("polo")) = "TI" then response.Write("selected")%>>TI</option>
                  <option value="Petróleo, Gás e Energia" <% if trim(request("polo")) = "Petróleo, Gás e Energia" then response.Write("selected")%>>Petróleo, Gás e Energia</option>
                  <option value="Plástico" <% if trim(request("polo")) = "Plástico" then response.Write("selected")%>>Plástico</option>
                  <option value="Real State" <% if trim(request("polo")) = "Real State" then response.Write("selected")%>>Real State</option>
                  <option value="Suinocultura" <% if trim(request("polo")) = "Suinocultura" then response.Write("selected")%>>Suinocultura</option>
                  <option value="Têxtil" <% if trim(request("polo")) = "Têxtil" then response.Write("selected")%>>Têxtil</option>
                  <option value="Turismo" <% if trim(request("polo")) = "Turismo" then response.Write("selected")%>>Turismo</option>
                  <option value="Vitivinicultura" <% if trim(request("polo")) = "Vitivinicultura" then response.Write("selected")%>>Vitivinicultura</option>
                </select>
              </span></td>
              <td><input type="submit" name="button2" id="button2" value="Buscar"/></td>
              <td align="right">&nbsp;</td>
              <td align="right">&nbsp;</td>
              <td align="right">&nbsp;</td>
              <td align="right">&nbsp;</td>
              <td align="center">&nbsp;</td>
            </tr>
          </table>
        </form></td>
      </tr>
      <tr>
        <td scope="row">
<% 
	if not dados.eof then 	

%>
        <table width="100%" border="0" align="center" cellpadding="2" cellspacing="0">
      <tr>
        <td height="20" colspan="8" class="borda"><strong>BUSCA:</strong></td>
      </tr>
      <tr>
        <td colspan="8" height="10"></td>
      </tr>
      <tr>
        <td width="46" height="30" bgcolor="#CCCCCC" style="text-align:center"><a href="cadastros.asp?filtro=stu&ordem=<%=ordem%>">STU</a><%if request("filtro") = "stu" then %> <img src="<%=img%>" width="10" height="10" alt="Crescente" border="0" /> <% end if %> </td>
        <td width="89" bgcolor="#CCCCCC" style="text-align:center"><a href="cadastros.asp?filtro=nivel&amp;ordem=<%=ordem%>">PERFIL</a>
          <%if request("filtro") = "nivel" then %>
          <img src="<%=img%>" width="10" height="10" alt="Crescente" border="0" />
          <% end if %></td>
        <td width="133" bgcolor="#CCCCCC" style="text-align:center"><a href="cadastros.asp?filtro=empresa&amp;ordem=<%=ordem%>">EMPRESA</a><%if request("filtro") = "empresa" then %> <img src="<%=img%>" width="10" height="10" alt="Crescente" border="0" /> <% end if %> </td>
        <td width="102" bgcolor="#CCCCCC" style="text-align:center"><a href="cadastros.asp?filtro=polo&amp;ordem=<%=ordem%>">SEGMENTO</a>
          <%if request("filtro") = "polo" then %>
          <img src="<%=img%>" width="10" height="10" alt="Crescente" border="0" />
          <% end if %></td>
        <td width="130" height="30" bgcolor="#CCCCCC" align="center"><a href="cadastros.asp?filtro=nome&ordem=<%=ordem%>">NOME </a><%if request("filtro") = "nome" then %> <img src="<%=img%>" width="10" height="10" alt="Crescente" border="0" /> <% end if %> </td>
        <td width="145" height="30" bgcolor="#CCCCCC" align="center"><a href="cadastros.asp?filtro=cargo&ordem=<%=ordem%>">CARGO</a><%if request("filtro") = "cargo" then %> <img src="<%=img%>" width="10" height="10" alt="Crescente" border="0" /> <% end if %> </td>
        <td width="184" height="30" align="center" bgcolor="#CCCCCC"><a href="cadastros.asp?filtro=email&ordem=<%=ordem%>">E-MAIL</a><%if request("filtro") = "email" then %> <img src="<%=img%>" width="10" height="10" alt="Crescente" border="0" /> <% end if %> </td>
        <td width="89" height="30" bgcolor="#CCCCCC" style="text-align:center">AÇÕES</td>
      </tr>
    <%		
		x = 0
		while not dados.eof
		x = x + 1
	%>  
          <tr id="linha<%=x%>" onMouseOver="Mudacor(0,'linha<%=x%>')" onMouseOut="Mudacor(1,'linha<%=x%>')" class="borda">
            <th height="20">
			<%
		
		select case dados("stu")
			case "0"
				response.Write("<a href='status_cad.asp?id="&dados.fields(0)&"&stu="&dados("stu")&"'><img src='images/icon_vermelho.png' border='0'></a>") 
			case "1"
				response.Write("<a href='status_cad.asp?id="&dados.fields(0)&"&stu="&dados("stu")&"'><img src='images/icon_verde.png' border='0'></a>") 
		end select 
		
		%></th>
            <td height="20"><%
				select case dados("nivel")
					case 1
						response.Write("Administrador")
					case 2
						response.Write("Gerente")
					case 3
						response.Write("Vendedor")
					case 4
						response.Write("Comprador")
				end select
			%></td>
            <td height="20"><%= dados("empresa")%></td>
            <td height="20"><%= dados("polo")%></td>
            <td height="20"><%= dados("nome")%></td>
            <td height="20" align="center"><%= dados("cargo")%></td>
            <td height="20" align="center"><%= dados("email")%></td>
            <td height="20" align="center">
            	
                <a href="cadastros_novo.asp?id_cad=<%= dados.fields(0)%>&str=<%=request("str")%>">
                	<img src="images/icon_alterar.png" alt="Alterar" width="25" height="25" border="0" />
                </a><a href="projetos_empresa.asp?id_cad=<%= dados.fields(0)%>"><img src="images/icon_projetos2.png" width="18" height="25" border="0" /></a> 
                <% if session("lv_user") = 1 then %>
                &nbsp;
                <a href="javascript:void(0);" onclick="exclui('','cadastros_novo.asp?id_cad=<%= dados.fields(0)%>&empresa=<%=dados("empresa")%>&nome=<%=dados("nome")%>&exc=1')">
                	<img src="images/icon_excluir.png" alt="Alterar" width="26" height="25" border="0" />
                </a>
                <% end if %>
            </td>
          </tr>
    <%
		dados.movenext
		wend
		set dados = nothing
	%>                 
        </table>
       </td>
      </tr>

      <% else 
      response.Write("Registro não encontrado")
      end if
    %>
</table></td>
  </tr>

</table>
<!-- #Include File="rodape.asp" --> 
</body>
</html>


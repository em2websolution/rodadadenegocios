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

function salva_selecao(x,y,w) {	

	if (document.getElementById("ids"+y).checked == true ) {
		z = 0
	} else {
		z = 1
	}
	
	ajax = IniciaAjax();
	if (ajax) {
		ajax.onreadystatechange = function(){
			if(ajax.readyState<4)	{
				var carregando = "<span align='center'><img src='images/carregando.gif' /></span>";
				document.getElementById("icon"+y).innerHTML=carregando;
			} else if(ajax.readyState==4){
					document.getElementById("icon"+y).innerHTML="<span align='center'><img src='images/checked.gif' /></span>";
					if (z==1){
						document.getElementById("msg"+y).innerHTML="Seleção excluida com sucesso!";
					} else {
						document.getElementById("msg"+y).innerHTML="Seleção salva com sucesso!";
					}
				} else {
					alert(ajax.responseText);
				}
		}
		dados = 'id_rod='+x+'&id_usu='+w+'&chk='+z;
		ajax.open('POST','includes/adm_salvaSelecao.asp',true);
		ajax.setRequestHeader('Content-Type','application/x-www-form-urlencoded');
		ajax.send(dados);
	}
	
}

function salva_dta_limite(x) {	
	if (document.getElementById("dta_i").value == "") {
		alert("Escolha a data limite");
	} else {
		ajax = IniciaAjax();
		if (ajax) {
			ajax.onreadystatechange = function(){
				if(ajax.readyState<4)	{
					var carregando = "<span align='center'><img src='images/carregando.gif' /></span>";
					document.getElementById("dta_limite").innerHTML=carregando;
				} else if(ajax.readyState==4){
						document.getElementById("dta_limite").innerHTML="<span align='center'><img src='images/checked.gif' /></span>";
						document.getElementById("dta_limite_msg").innerHTML="Data salva com sucesso!";
					} else {
						alert(ajax.responseText);
					}
			}
			dados = 'id_rod='+x+'&dta_limite='+document.getElementById("dta_i").value;		
			ajax.open('POST','includes/adm_salvaDta_limite.asp',true);
			ajax.setRequestHeader('Content-Type','application/x-www-form-urlencoded');
			ajax.send(dados);
		}
	}
	
}


function envia_convite(x) {	
	if (document.getElementById("dta_i").value == "") {
		alert("Escolha a data limite");
	} else {
		ajax = IniciaAjax();
		if (ajax) {
			ajax.onreadystatechange = function(){
				if(ajax.readyState<4)	{
					var carregando = "<span align='center'><img src='images/carregando.gif' /></span>";
					document.getElementById("dta_limite").innerHTML=carregando;
				} else if(ajax.readyState==4){
						MTX_campos = ajax.responseText.split('|');
						document.getElementById("dta_limite").innerHTML="<span align='center'><img src='images/checked.gif' /></span>";
						if (MTX_campos.length != 1 ){
							for(i = 0; i < MTX_campos.length; i++){
								//alert(MTX_campos[i]);
								document.getElementById("msg"+MTX_campos[i]).innerHTML=" Convites enviados com sucesso!";
							}
						} else {
								//alert(ajax.responseText);
								document.getElementById("dta_limite_msg").innerHTML=" Nenhuma nova empresa selecionada!";
						}
					} else {
						alert(ajax.responseText);
					}
			}
			dados = 'id_rod='+x;		
			ajax.open('POST','includes/adm_enviaConvite.asp',true);
			ajax.setRequestHeader('Content-Type','application/x-www-form-urlencoded');
			ajax.send(dados);
		}
	}
	
}

function checkAll(field)
{
for (i = 0; i < field.length; i++)
	field[i].checked = true;
	form_chekall.submit();
}

function checkOff(field)
{
for (i = 0; i < field.length; i++)
	field[i].checked = false;
	form_chekoff.submit();
}
function validaFiltro(){
	var form = document.form3;

	form.submit();

}

</script>
<%
if session("user_name") = "" then
	response.Redirect("index.asp")
	response.End()
end if



if request("id") <> "" then
	sql = ("SELECT * FROM tb_rodada where id_rod = "&request("id")&"")
else
	response.Redirect("projetos.asp")
	response.End()
end if

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
    <% if not dados.eof then %>
      <tr>
        <td scope="row" style="text-align:left"><h1>Assunto: <%= dados("assunto")%></h1></td>
      </tr>
      <tr>
        <td scope="row" valign="top" align="center">
    <%		
		
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
		
		reunioes = ((minutos-min_a))/(tempo+intervalo) * mesas
		
		
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

	if dados("dta_limite") <> "" then 
		dta_limite = dados("dta_limite")
	else
		dta_limite = ""
	end if
	
	%>    
        <table width="100%" border="0" align="center" cellpadding="2" cellspacing="0">
      <tr>
        <td height="2" colspan="7" class="borda"></td>
      </tr>
      <tr>
        <td colspan="7" height="10"></td>
      </tr>
      <tr>
        <th width="88" height="20" style="text-align:right">DATA:</th>
        <td width="247" height="20" bgcolor="#FFFFFF"><%= (dta_i & " à " & dta_f)%></td>
        <th width="149" height="20" align="right">DIAS DE EVENTO:</th>
        <td width="185" height="20" bgcolor="#FFFFFF"><%=dias%></td>
        <th height="20" colspan="3" bgcolor="#FFFFFF">ENVIAR CONVITE AOS PARTICIPANTES</th>
        </tr>
      <tr>
        <th height="20" style="text-align:right">HORÁRIO:</th>
        <td height="20" bgcolor="#FFFFFF"><%= hs_i&" às "&hs_f %></td>
        <th height="20" align="right">DURAÇÃO REUNIÃO:</th>
        <td height="20" bgcolor="#FFFFFF"><%= tempo%></td>
        <th width="105" height="20" bgcolor="#FFFFFF">Data limite:</th>
        <td width="76" height="20" bgcolor="#FFFFFF"><form id="form2" name="form2" method="post" action=""><span id="sprytextfield1">
          <input name="dta_i" type="text" class="campo" id="dta_i" onclick="displayCalendar(document.forms[0].dta_i,'dd/mm/yyyy',this)" size="10" tabindex="1" value="<%=dta_limite%>"/>
          <span class="textfieldRequiredMsg"><br />
          </span></span></form>
          </td>
        <td width="76" bgcolor="#FFFFFF"><img src="images/icon_ok.png" width="22" height="25" onclick="salva_dta_limite(<%= request("id")%>)" style="cursor:pointer"/></td>
      </tr>
      <tr>
        <th height="20" style="text-align:right">ALMOÇO:</th>
        <td height="20" bgcolor="#FFFFFF"><%= hs_ai&" às "&hs_af %></td>
        <th height="20" align="right">INTERVALO:</th>
        <td height="20" bgcolor="#FFFFFF"><%= dados("intervalo")&"min"%></td>
        <td height="20" bgcolor="#FFFFFF">&nbsp;</td>
        <td height="20" colspan="2" bgcolor="#FFFFFF"><div id="dta_limite"></div><div id="dta_limite_msg"></div></td>
      </tr>
      <tr>
        <th height="20" style="text-align:right">LOCAL:</th>
        <td height="20" bgcolor="#FFFFFF"><%= dados("local")%></td>
        <th height="20" align="right">TOTAL DE REUNIÕES:</th>
        <td height="20" bgcolor="#FFFFFF"><%= mesas &" Mesa(s) <br>"& CInt(reunioes) & " por dia. " & Cint(dias)*cint(reunioes) & " Total." %></td>
        <th height="20" bgcolor="#FFFFFF">Enviar convite:</th>
        <td height="20" colspan="2" bgcolor="#FFFFFF"><table width="100%" border="0" cellspacing="2" cellpadding="0">
          <tr>
              <td width="43%" align="right"><img src="images/icon_email.jpg" alt="" width="38" height="30" style="cursor:pointer" onclick="envia_convite(<%= request("id")%>)" /></td>
              <td width="51%" align="right">Exportar</td>
              <td width="6%" align="right"><a href="includes/adm_excel.asp?id_rod=<%= request("id")%>&amp;tipo=3" target="_blank"><img src="images/icon_excel.png" alt="Exportar XLS" width="49" height="42" border="0" onclick="geraXLS()" style="cursor:pointer" /></a></td>
            </tr>
        </table></td>
      </tr>

    <%
		set dados = nothing

var_selecao = request("selecao")
if var_selecao = "" then
	var_selecao = 0
else
	var_selecao = request("selecao")
end if

		
	%>
        </table>
        <div class="borda">&nbsp;</div>
        <form action="projetos_selecao.asp" method="post" accept-charset="utf-8" name="form3" id="form3" onsubmit="validaFiltro();return false;">
          <table width="950" border="0" cellspacing="2" cellpadding="0">
            <tr>
              <td width="171"><strong>Filtro (email, empresa):</strong></td>
              <td width="272"><input name="str" type="text" id="str" value="<%=request("str")%>" style="width:250px" /></td>
              <td width="499"><span class="vermelho">
                <strong>Segmento</strong>:
                <select name="polo" size="1" id="polo"  tabindex="27">
                  <option value="" selected="selected">Selecione</option>
                  <option value="Agroindústria em geral" <% if trim(request("polo")) = "Agroindústria em geral" then response.Write("selected")%>>Agroindústria em geral</option>
                  <option value="Alimentos e Bebidas" <% if trim(request("polo")) = "Alimentos e Bebidas" then response.Write("selected")%>>Alimentos e Bebidas</option>
                  <option value="Apicultura" <% if trim(request("polo")) = "Apicultura" then response.Write("selected")%>>Apicultura</option>
                  <option value="Aquicultura" <% if trim(request("polo")) = "Aquicultura" then response.Write("selected")%>>Aquicultura</option>
                  <option value="Arquitetura" <% if trim(request("polo")) = "Arquitetura" then response.Write("selected")%>>Arquitetura</option>
                  <option value="Artesanato" <% if trim(request("polo")) = "Artesanato" then response.Write("selected")%>>Artesanato</option>
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
              </span>                <input type="submit" name="button2" id="button2" value="Buscar"/></td>
              </tr>
            <tr>
              <td><strong>Apenas Selecionados</strong></td>
              <td><input name="selecao" type="checkbox" id="selecao" value="1" <% if var_selecao = 1 then response.Write("checked='checked'")%> onclick="form.submit();" /></td>
              <td>&nbsp;</td>
            </tr>
          </table>
          <input name="id" type="hidden" id="id" value="<%= request("id")%>" />
        </form>
        <div class="borda">&nbsp;</div>
        <table width="950" border="0" cellspacing="2" cellpadding="0">
          <form id="form1" name="form1" method="post" action="chekall.asp">
            <%
		
		
		
		
	'response.Write(request.Form("str") & " - " & request.Form("polo"))
	if var_selecao = 0 then
		if trim(request("str")) <> "" then
			str = "empresa like '%"&request("str")&"%' or email like '%"&request("str")&"%' "
			if request("polo") <> "" then
				str = str & "and polo = '"&request("polo")&"'"
			end if
			sql = "select id_usu, emp_fantasia, nivel, polo, email from conceitobrazil.tb_login where nivel in (3,4) and stu=1 and "&str&" order by nivel desc, emp_fantasia"		
		else
			if trim(request("polo")) <> "" then
				str = str & "and polo = '"&request("polo")&"'"
			else
				str = ""
			end if
			sql = "select id_usu, emp_fantasia, nivel, polo, email from conceitobrazil.tb_login where nivel in (3,4) and stu=1 "&str&" order by nivel desc, emp_fantasia"		
		end if
	else
		sql = "select L.id_usu, L.emp_fantasia, L.nivel, L.polo, L.email from conceitobrazil.tb_login as L inner join conceitobrazil.tb_selecao as S on L.id_usu = S.id_usu where id_rod = "&request("id")&" order by nivel desc, emp_fantasia"
	end if	
		
		'response.Write(sql)
		
		'número de registros por página
		intRecordsForPage = 100

		'mensagem 
		strMessageOne = "registro encontrado" 
		strMessageMore = "registros encontrados" 

		set objConnection=server.createobject("adodb.connection") 
		'objConnection.open "DSN=conceitobrazil"
		objConnection.open "Driver={MariaDB ODBC 3.1 Driver};SERVER=192.168.40.15;USER=conceitobrazil;PASSWORD=@9vRp7i5;DATABASE=conceitobrazil;PORT=3306"

		set objRecordset = Server.CreateObject("adodb.recordset")
		with objRecordset 
			.Open sql,objConnection
			
			if not .eof then
				MyList = .GetRows()
			else
				response.Write("Nenhum registro encontrado!")
				response.End()
			end if
			
			.Close
		end with
		set objRecordset = nothing
		objConnection.close
		set objConnection = nothing
		
		intCol = cint(ubound(MyList,1))+1: intLin = cint(ubound(MyList,2))+1
		boolMountTable = true
		strMessage = "<b>" & intLin & "</b> " & strMessageOne
		if intLin > 1 then strMessage = "<b>" & intLin & "</b> " & strMessageMore
		
		intNumberThisPage=Request.QueryString("page")
		
		if intNumberThisPage = "" then intNumberThisPage = 1
		if isnumeric(intNumberThisPage) = false then intNumberThisPage = 1
		varTotalRecords = intLin 
		intTotalPages = int(varTotalRecords/intRecordsForPage)
		if intTotalPages < varTotalRecords/intRecordsForPage then intTotalPages = intTotalPages + 1
		if int(intNumberThisPage) > int(intTotalPages) then 
		intNumberThisPage = 1
		end if
		intLastRecord = intNumberThisPage * intRecordsForPage
		intFirstRecord = intLastRecord - intRecordsForPage
		if int(intTotalPages) = int(intNumberThisPage) then 
		intLastRecord = varTotalRecords
		end if
		strTarget=Mid(Trim(Request.ServerVariables("PATH_INFO")),InstrRev(varLocal,"/")+1 )
		
		titulo2 = ""

		for iList = intFirstRecord to intLastRecord-1
		'while not empresas.eof
			set verifica = Server.CreateObject("ADODB.Recordset")
			verifica.ActiveConnection = adoConn
			sql2 = "select id_usu from tb_selecao where id_rod="&request("id")&" and id_usu="&MyList(0,iList)&""
			verifica.Open(sql2)

			titulo1 = MyList(2,iList)
			if titulo2 <> titulo1 then
				titulo2 = titulo1
	%>
            <tr>
              <td colspan="6" scope="row"><h1>
			  <% 
			  	select case titulo1
					case 1
						response.Write("Administrador:")
					case 2
						response.Write("Gerente:")
			    	case 3
				    	response.Write("Vendedores:")
					case 4 
						response.Write("Compradores:") 
				end select
				%>
                <input name="id_rod" type="hidden" id="id_rod" value="<%= request("id")%>" />
                </h1></td>
              </tr>
            <tr>
              <th colspan="2" scope="row">Seleções</th>
              <th width="249">Empresas Fantasia</th>
              <th width="207">E-mail</th>
              <th width="202">Segmento</th>
              <th width="203">&nbsp;</th>
              </tr>
            <% end if	%>
            <tr>
              <th width="43" scope="row"><input type="checkbox" <% if not verifica.eof then response.Write("checked='checked'") : set verifica = nothing %> name="ids" id="ids<%=iList%>" onclick="salva_selecao(<%= request("id")%>,<%=iList%>,<%=MyList(0,iList)%>)" /><input name="nivel" type="hidden" id="nivel" value="<%=MyList(2,iList)%>" /></th>
              <th width="32" scope="row"><div class="aviso" id="icon<%=iList%>"></div></th>
              <td><%= ucase(MyList(1,iList))%></td>
              <td><%= ucase(MyList(4,iList))%></td>
              <td align="left"><%= ucase(MyList(3,iList))%></td>
              <td align="left"><div class="aviso" id="msg<%=iList%>"></div></td>
              </tr>
            <% next %>
          </form>
          
        </table>
<table width="950" border="0" cellspacing="2" cellpadding="0">
			  <tr>
				<td scope="row">
                <table width="230" border="0" cellpadding="0" cellspacing="2">
                  <tr>
                    <th scope="row">&nbsp;</th>
                    <th>&nbsp;</th>
                  </tr>
                  <tr>
                    <th width="111" scope="row"><form id="form_chekall" name="form_chekall" method="post" action="chekall.asp">
                      <input type="button" name="CheckAll" value="Selecionar todos" onclick="checkAll(document.form1.ids)" />
                      <input name="page" type="hidden" id="page" value="<%= Request.QueryString("page")%>" />
                      <input name="id_rod" type="hidden" id="id_rod" value="<%= request("id")%>" />
                      <input name="str" type="hidden" id="str" value="<%=request("str")%>" style="width:350px" />
                	  <input name="polo" type="hidden" id="polo" value="<%= request("polo")%>" />

                    </form></th>
                    <th width="113"><form id="form_chekoff" name="form_chekoff" method="post" action="chekoff.asp">
                      <input type="button" name="CheckOff" value="Desmarcar todos" onclick="checkOff(document.form1.ids)" />
                      <input name="page" type="hidden" id="page" value="<%= Request.QueryString("page")%>" />
                      <input name="id_rod" type="hidden" id="id_rod" value="<%= request("id")%>" />
                      <input name="str" type="hidden" id="str" value="<%=request("str")%>" style="width:350px" />
                      <input name="polo" type="hidden" id="polo" value="<%= request("polo")%>" />
                    </form></th>
                  </tr>
                </table>
                <div class="borda">&nbsp;</div></td>
			  </tr>
		</table>
<table width="100%" border="1" align="center" cellpadding="2" cellspacing="2">
<tr> 
<td width="25%" align=left><%=strMessage%></td>
<td width="50%" align=center>
<%if intNumberThisPage <> 1 then%> 
<a href=<%=strTarget%>?page=1&id=<%=request("id")%>&str=<%=request("str")%>&polo=<%=request("polo")%>&selecao=<%=var_selecao%>>primeira</a> 
| <a href=<%=strTarget%>?page=<%=intNumberThisPage-1%>&id=<%=request("id")%>&str=<%=request("str")%>&polo=<%=request("polo")%>&selecao=<%=var_selecao%>>anterior</a> 
<%else%> 
primeira | anterior 
<%end if%><%if cLng(intNumberThisPage) = intTotalPages or boolMountTable = false then%> 
próxima | ultima 
<%else%> 
| <a href=<%=strTarget%>?page=<%=intNumberThisPage+1%>&id=<%=request("id")%>&str=<%=request("str")%>&polo=<%=request("polo")%>&selecao=<%=var_selecao%>>próxima</a> 
| <a href=<%=strTarget%>?page=<%=intTotalPages%>&id=<%=request("id")%>&str=<%=request("str")%>&polo=<%=request("polo")%>&selecao=<%=var_selecao%>>última</a> 
<%end if%> </td>

<td width="25%" align=right> 
<%if boolMountTable = true then%> 
página <b><%=intNumberThisPage%></b> de <b><%=intTotalPages%></b></td>
<%else%> 
&nbsp; 
<%end if%>

</tr>
</table>
      <% else 
		response.Write("Registro não encontrado")	
	%>
    </td>
      </tr>
      <%end if%>
    </table></td>
  </tr>
</table>
<!-- #Include File="rodape.asp" --> 
</body>
</html>

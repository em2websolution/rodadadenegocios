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



function Gera_Horario(x,y,z,w) {	
	ajax = IniciaAjax();
	if (ajax) {
		ajax.onreadystatechange = function(){
			if(ajax.readyState<4)	{
				var carregando = "<span align='center'><img src='images/carregando.gif' /></span>";
				document.getElementById("msg"+y).innerHTML=carregando;
			} else if(ajax.readyState==4){
					document.getElementById("horario"+y).innerHTML=ajax.responseText;
					document.getElementById("msg"+y).innerHTML="<span align='center'><img src='images/checked.gif' /></span>";
				} else {
					alert(ajax.responseText);
				}
		}
		
		if (document.getElementById("hr"+y)) {
			id_ind = document.getElementById("hr"+y).value
		} else {
			id_ind = 0
			}
		
		dados = 'id_rod='+x+'&id_indicado='+y+'&dta='+z+'&id_ind='+id_ind+'&id_usu='+w;
		ajax.open('POST','includes/adm_geraHorario.asp',true);
		ajax.setRequestHeader('Content-Type','application/x-www-form-urlencoded');
		ajax.send(dados);
		//alert(dados);
	}
}



function salva_selecao(x,y,w,z) {	
	ajax = IniciaAjax();
	if (ajax) {
		ajax.onreadystatechange = function(){
			if(ajax.readyState<4)	{
				var carregando = "<span align='center'><img src='images/carregando.gif' /></span>";
				document.getElementById("msg"+y).innerHTML=carregando;
			} else if(ajax.readyState==4){
					//alert(ajax.responseText);
					window.location = 'projetos_indicacao.asp?id=<%=request.QueryString("id")%>&id_usu='+z;
				} else {
					alert(ajax.responseText);
				}
		}
		dados = 'id_ind='+x+'&id_ind_old='+w+'&id_indicado='+y+'&id_rod='+<%=request.QueryString("id")%>+'&id_usu='+z;
		
		ajax.open('POST','includes/adm_salvaIndicacao.asp',true);
		ajax.setRequestHeader('Content-Type','application/x-www-form-urlencoded');
		ajax.send(dados);
		//alert(dados);
	}
}



function cancela(x,y) {	
	ajax = IniciaAjax();
	if (ajax) {
		ajax.onreadystatechange = function(){
			if(ajax.readyState<4)	{
				var carregando = "<span align='center'><img src='images/carregando.gif' /></span>";
				document.getElementById("msg_data").innerHTML=carregando;
			} else if(ajax.readyState==4){
					window.location = 'projetos.asp';
				} else {
					alert(ajax.responseText);
				}
		}
		dados = 'id_rod='+x+'&id_usu='+y;
		ajax.open('POST','includes/adm_cancelaIndicacao.asp',true);
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

if request.QueryString("id") <> "" then
	sql = ("SELECT * FROM tb_rodada where id_rod = "&request.QueryString("id")&"")
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
    <% if not dados.eof then 
	assunto = dados("assunto")
	
	%>
     <tr>
        <td scope="row" style="text-align:left"><h1>
		<%if session("lv_user") = 1 or session("lv_user") = 2 then
	  			response.Write("<h1><a href='javascript:void(0);history.back(-1);'><img src='images/icon_voltar.png' width='27' height='30' /></a>&nbsp;Assunto: " & assunto & "</h1>")
			else
	  			response.Write("<h1>Assunto: "&assunto&"</h1>")
		end if
	  %>
        </h1></td>
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
	
	ancora = dados("ancora")

	if request("id_usu") <> "" then
		id_usu_ind = request("id_usu")
	else
		id_usu_ind = session("user_rd")
	end if

set empresa = Server.CreateObject("ADODB.Recordset")
empresa.ActiveConnection = adoConn
empresa.Open("select * from tb_login where id_usu = "&id_usu_ind&"")

razao = empresa("empresa")
	
	%>    
        <table width="100%" border="0" align="center" cellpadding="2" cellspacing="0">
      <tr>
        <td height="2" colspan="6" class="borda"></td>
      </tr>
      <tr>
        <td colspan="6" height="10"></td>
      </tr>
      <tr>
        <th width="88" height="20" style="text-align:right">DATA:</th>
        <td width="247" height="20" bgcolor="#FFFFFF"><%= (dta_i & " à " & dta_f)%></td>
        <th width="149" height="20" align="right">DIAS DE EVENTO:</th>
        <td width="185" height="20" bgcolor="#FFFFFF"><%=dias%></td>
        <th height="20" bgcolor="#FFFFFF">DATA LIMITE:</th>
        <th height="20" bgcolor="#FFFFFF"><%= dados("dta_limite")%></th>
        </tr>
      <tr>
        <th height="20" style="text-align:right">HORÁRIO:</th>
        <td height="20" bgcolor="#FFFFFF"><%= hs_i&" às "&hs_f %></td>
        <th height="20" align="right">DURAÇÃO REUNIÃO:</th>
        <td height="20" bgcolor="#FFFFFF"><%= tempo%></td>
        <th width="105" height="20" bgcolor="#FFFFFF">&nbsp;</th>
        <td height="20" bgcolor="#FFFFFF">&nbsp;</td>
        </tr>
      <tr>
        <th height="20" style="text-align:right">ALMOÇO:</th>
        <td height="20" bgcolor="#FFFFFF"><%= hs_ai&" às "&hs_af %></td>
        <th height="20" align="right">INTERVALO:</th>
        <td height="20" bgcolor="#FFFFFF"><%= dados("intervalo")&"min"%></td>
        <td height="20" bgcolor="#FFFFFF">&nbsp;</td>
        <td height="20" bgcolor="#FFFFFF">&nbsp;</td>
      </tr>
      <tr>
        <th height="20" style="text-align:right">LOCAL:</th>
        <td height="20" bgcolor="#FFFFFF"><%= dados("local")%></td>
        <th height="20" align="right">TOTAL DE REUNIÕES:</th>
        <td height="20" bgcolor="#FFFFFF"><%= mesas &" Mesa(s) <br>"& CInt(reunioes) & " por dia. " & Cint(dias)*cint(reunioes) & " Total." %></td>
        <th height="20" bgcolor="#FFFFFF">&nbsp;</th>
        <td height="20" bgcolor="#FFFFFF">&nbsp;</td>
      </tr>

    <%
		set dados = nothing
		
	%>
        </table>
        <form id="form1" name="form1" method="post" action="">
        <table width="950" border="0" cellspacing="2" cellpadding="0">
<%
		if ancora <> empresa("nivel") then
			sql = "SELECT distinct I.id_indicado, L.emp_fantasia, L.perfil FROM conceitobrazil.tb_indicacao as I inner join conceitobrazil.tb_login as L on I.id_indicado = L.id_usu inner join conceitobrazil.tb_rodada as R on I.id_rod = R.id_rod where I.id_rod ="&request.QueryString("id")&"  and R.ancora = L.nivel order by L.empresa"		
		else
			sql = "SELECT distinct I.id_indicado, L.emp_fantasia, L.perfil FROM conceitobrazil.tb_indicacao as I inner join conceitobrazil.tb_login as L on I.id_indicado = L.id_usu inner join conceitobrazil.tb_rodada as R on I.id_rod = R.id_rod where I.id_rod ="&request.QueryString("id")&"  and R.ancora <> L.nivel order by L.empresa"	
		end if

		if empresa("nivel") = 3 then
			titulo_empresas = "Compradores"
		else
			titulo_empresas = "Vendedores"
		end if		
		

		'response.Write(sql)
		
		'número de registros por página
		intRecordsForPage = 500

		'mensagem 
		strMessageOne = "registro encontrado" 
		strMessageMore = "registros encontrados" 

		set objConnection=server.createobject("adodb.connection") 
		'objConnection.open "DSN=conceitobrazil"

		objConnection.Open "Driver={MariaDB ODBC 3.1 Driver};SERVER=192.168.40.15;USER=conceitobrazil;PASSWORD=@9vRp7i5;DATABASE=conceitobrazil;PORT=3306"
		set objRecordset = Server.CreateObject("adodb.recordset")
		fim = 0
		with objRecordset 
			.Open sql,objConnection
			if not .eof then
				MyList = .GetRows()
			else 
				fim = 1
			end if
			.Close
		end with
		set objRecordset = nothing
if fim = 0 then		
		intCol = cint(ubound(MyList,1))+1: intLin = cint(ubound(MyList,2))+1
		boolMountTable = true
		strMessage = "<b>" & intCol & "</b> - <b>" & intLin & "</b> " & strMessageOne

	'response.write(strMessage)
		
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
		
%>
			  <tr>
			    <td colspan="4" scope="row" style="color:#FF0000">EMPRESA: <%=ucase(razao)%>&nbsp;</td>
			    </tr>
			  <tr>
			    <td colspan="4" scope="row" align="center">Selecione abaixo as datas que sua empresa estará disponível e as empresas de seu interesse para a Rodade de Negócios na <%= assunto%></td>
			    </tr>

			  <tr>
				<td colspan="4" scope="row"><h1><%= titulo_empresas%>:</h1></td>
			  </tr>
			  <tr>
				<th width="557" scope="row">Empresas</th>
				<th width="168">DIA</th>
				<th width="135">HORÁRIO</th>
				<th width="80">&nbsp;</th>
			  </tr>
              		
<% 


for iList = intFirstRecord to intLastRecord-1

%>
  <tr>
    <td scope="row"><%= ucase(MyList(1,iList))%><div class="phone" style="font-size:10px"><%= (MyList(2,iList))%></div><br /></td>
    <th>
      <select name="dia<%=MyList(0,iList)%>" id="dia<%=MyList(0,iList)%>" onChange="Gera_Horario(<%=request.QueryString("id")%>,<%=MyList(0,iList)%>,this.value,<%=id_usu_ind%>)">
        <option value="12/12/2000">Não tenho interesse</option>
      <% 
	  horario = ""
	  var_stu = ""
	  var_stu2 = ""
	  dta = ""
	  id_ind = 0
	  for x=0 to dias-1
	  	  	dia_reuniao = DateAdd("d", x, dta_i)
	  	
			sql = "SELECT i.id_ind, DATE_FORMAT(I.hr,'%Y-%m-%d') as Dta, DATE_FORMAT(I.hr,'%H:%i') as Horario FROM conceitobrazil.tb_indicacao as I inner join conceitobrazil.tb_login as L on I.id_indicado = L.id_usu where I.id_rod ="&request.QueryString("id")&" and I.id_indicado="&MyList(0,iList)&" and i.id_usu = "&id_usu_ind&" order by Dta"
			set objRecordset = Server.CreateObject("adodb.recordset")
			with objRecordset 
				.Open sql,objConnection
				if not .eof then
					MyList2 = .GetRows()
					
					d1 = year(dia_reuniao) &"-"& month(dia_reuniao) &"-"& day(dia_reuniao)
					d2 = year(MyList2(1,0)) &"-"& month(MyList2(1,0)) &"-"& day(MyList2(1,0))
					
					if d1 = d2 then
						var_stu = " selected='selected'"
						var_stu2 = " selected='selected'"
						if day(d2) < 10 then 
							dia = "0" & day(d2)
						else
							dia = day(d2)
						end if
						
						if month(d2) < 10 then 
							mes = "0" & month(d2)
						else
							mes = month(d2)
						end if						
						dta = year(d2) &"-"& mes &"-"& dia
						horario = MyList2(2,0)
						id_ind = MyList2(0,0)
					else
						var_stu = ""
					end if
					
					.Close
				end if
			end with
			set objRecordset = nothing
	  %>
      	<option value="<%=dia_reuniao%>" <%=var_stu%>><%=dia_reuniao%></option>
      <% 
	  next 					
%>
      </select>
      </th>
    <th>
    <div id="horario<%=MyList(0,iList)%>">
    	<% if horario <> "" then 
		
			sql = ("SELECT DATE_FORMAT(I.hr,'%H:%i') as Horario, i.id_ind FROM conceitobrazil.tb_indicacao as I inner join conceitobrazil.tb_login as L on I.id_indicado = L.id_usu where I.id_rod ="&request.QueryString("id")&" and I.id_indicado="&MyList(0,iList)&" and I.stu = 1 and DATE_FORMAT(I.hr,'%Y-%m-%d') = '"&dta&"' order by Horario")
			set lista = Server.CreateObject("ADODB.Recordset")
			lista.ActiveConnection = objConnection
			lista.Open(sql)
			'response.Write(sql)
		%>

            <select name="hr<%=MyList(0,iList)%>" id="hr<%=MyList(0,iList)%>" onChange="salva_selecao(this.value,<%=MyList(0,iList)%>,<%=id_ind%>,<%=id_usu_ind%>)">
                <option value="0">Escolha o horário</option>
              	<option value="<%=id_ind%>" <%=var_stu2%>><%=horario%></option>
			  <% while not lista.eof
				set verifica = Server.CreateObject("ADODB.Recordset")
				verifica.ActiveConnection = objConnection
				sql = ("SELECT i.id_ind FROM conceitobrazil.tb_indicacao as i inner join conceitobrazil.tb_login as L on I.id_indicado = L.id_usu where I.id_rod ="&request.QueryString("id")&"  and i.stu = 0 and I.hr= '"&dta&" "&lista.fields(0)&"' and I.id_usu <> 'null' and i.id_usu="&id_usu_ind&"")
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
		<%end if%>
    
    </div>
    </th>
    <td align="left"><div class="aviso" id="msg<%=MyList(0,iList)%>"></div></td>
  </tr>
  <% next %>
  <tr>
    <th colspan="4" scope="row">&nbsp;</th>
  </tr>
  <tr>
    <th colspan="4" scope="row"><a href="javascript:void(0);" onClick="cancela(<%= request.QueryString("id")%>,<%=id_usu_ind%>)"><img src="images/cancelar.png" alt="Cancelar participação" width="195" height="40" border="0" /></a> 
    <%if session("lv_user") = 1 or session("lv_user") = 2 then%>
    	<a href="indicados.asp?id=<%=request.QueryString("id")%>"><img src="images/confirmar.png" alt="Confirmar participação" width="195" height="40" border="0" /></a>
    <% else %>
    	<a href="projetos.asp?c=1"><img src="images/confirmar.png" alt="Confirmar participação" width="195" height="40" border="0" /></a>
    <% end if %>
    </th>
    </tr>

</table>
 </form>
<% if varTotalRecords >= 500 then %>
<table width="100%" border="1" align="center" cellpadding="2" cellspacing="2">
<tr> 
<td width="25%" align=left><%=strMessage%></td>
<td width="50%" align=center>
<%if intNumberThisPage <> 1 then%> 
<a href=<%=strTarget%>?page=1&id=<%=request.QueryString("id")%>>primeira</a> 
| <a href=<%=strTarget%>?page=<%=intNumberThisPage-1%>&id=<%=request.QueryString("id")%>>anterior</a> 
<%else%> 
primeira | anterior 
<%end if%><%if cLng(intNumberThisPage) = intTotalPages or boolMountTable = false then%> 
próxima | ultima 
<%else%> 
| <a href=<%=strTarget%>?page=<%=intNumberThisPage+1%>&id=<%=request.QueryString("id")%>>próxima</a> 
| <a href=<%=strTarget%>?page=<%=intTotalPages%>&id=<%=request.QueryString("id")%>>última</a> 
<%end if%> </td>

<td width="25%" align=right> 
<%if boolMountTable = true then%> 
página <b><%=intNumberThisPage%></b> de <b><%=intTotalPages%></b></td>
<%else%> 
&nbsp; 
<%end if%>

</tr>
</table>
      <%
	  end if
	   else 
		response.Write("Agenda não disponível")	
end if
	   else 
		response.Write("Registro não encontrado")	
	%>
    </td>
      </tr>
      <%end if
	  
	  		objConnection.close
		set objConnection = nothing

	  %>
    </table></td>
  </tr>
</table>
<!-- #Include File="rodape.asp" --> 
</body>
</html>

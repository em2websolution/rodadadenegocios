<!-- #Include File="top.asp" -->
<link href="css/SpryValidationTextField.css" rel="stylesheet" type="text/css" />
<link href="css/SpryValidationSelect.css" rel="stylesheet" type="text/css" />
<script src="includes/SpryValidationTextField.js" type="text/javascript"></script>
<script src="includes/SpryValidationSelect.js" type="text/javascript"></script>
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

function altera(x) {	
	ajax = IniciaAjax();
	if (ajax) {
		ajax.onreadystatechange = function(){
			if(ajax.readyState==4)	{
				if(ajax.status==200){
					MTX_campos = ajax.responseText.split('|');
					document.form1.id_rod.value = MTX_campos[0];
					document.form1.local.value = MTX_campos[1];
					document.form1.assunto.value = MTX_campos[2];
					
					
					var date1 = MTX_campos[3];
					var date2 = MTX_campos[4];
					date1 = date1.split("/");
					date2 = date2.split("/");
					
					var sDate = new Date(date1[0]+"/"+date1[1]+"/"+date1[2]);	
					var curr_date1 = sDate.getDate();
					var curr_month1 = sDate.getMonth() + 1; //Months are zero based
					var curr_year1 = sDate.getFullYear();
					if (curr_date1<10) {
						curr_date1 = "0" + curr_date1
					}

					if (curr_month1<=10) {
						curr_month1 = "0" + curr_month1
					}
					

					var eDate = new Date(date2[0]+"/"+date2[1]+"/"+date2[2]);	
					var curr_date2 = eDate.getDate();
					var curr_month2 = eDate.getMonth() + 1; //Months are zero based
					var curr_year2 = eDate.getFullYear();
					if (curr_date2<10) {
						curr_date2 = "0" + curr_date2
					}

					if (curr_month2<=10) {
						curr_month2 = "0" + curr_month2
					}

					document.form1.dta_i.value = (curr_date1 + "/" + curr_month1 + "/" + curr_year1);
					document.form1.dta_f.value = (curr_date2 + "/" + curr_month2 + "/" + curr_year2);

					
					document.form1.hs_i.value = MTX_campos[5];
					document.form1.hs_f.value = MTX_campos[6];
					
					document.form1.tempo.value = MTX_campos[7];
					document.form1.intervalo.value = MTX_campos[8];
					
					var almoco = MTX_campos[9];
					almoco = almoco.split("-");
					document.form1.hs_ai.value = almoco[0];
					document.form1.hs_af.value = almoco[1];		
					document.form1.inpImgURL.value = MTX_campos[12];
					document.form1.mesas.value = MTX_campos[13];
					document.form1.mensagem.value = MTX_campos[14];
					document.form1.inpImgURL2.value = MTX_campos[15];
					document.form1.ancora.value = MTX_campos[16];
					document.form1.inpImgURL3.value = MTX_campos[17];

					document.form1.button.value = 'ALTERAR';
					
					<% if session("lv_user") = 1 then %>
						document.getElementById('iconExcluir').innerHTML="ou <input type='button' value='EXCLUIR' onclick=exclui('','projetos_novo.asp?id="+MTX_campos[0]+"&amp;exc=1')>";
					<% end if %>
					
				} else {
					alert(ajax.responseText);
				}
			}
		}
		dados = 'id='+x;
		ajax.open('POST','includes/adm_pegaDprojetos.asp',true);
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
	response.Write("<script language='javascript' type='text/javascript'>altera("&request.QueryString("id")&");</script>")
end if


'Grava
if trim(request.Form("button")) = "INCLUIR" then
	dta_i = year(request.Form("dta_i")) &"-"& month(request.Form("dta_i")) &"-"& day(request.Form("dta_i"))
	dta_f = year(request.Form("dta_f")) &"-"& month(request.Form("dta_f")) &"-"& day(request.Form("dta_f"))
	hs_i = request.Form("hs_i")	
	hs_f = request.Form("hs_f")
	hs_ai = request.Form("hs_ai")
	hs_af = request.Form("hs_af")
	almoco = hs_ai & "-"& hs_af
	mensagem = Replace(request.Form("mensagem"),Chr(13), "<BR>")
	
	set novo = adoConn.execute("insert into tb_rodada (dta_ini, dta_fim, tempo, almoco, assunto, intervalo, local, manual, mesas, mensagem, banner, ancora, logo_emp) values ('"&dta_i &" "& hs_i&"','"&dta_f &" "& hs_f&"',"& request.Form("tempo") &",'"& almoco &"','"& replace(request.Form("assunto"),"'","´") &"','"& replace(request.Form("intervalo"),"'","´")&"','"& replace(request.Form("local"),"'","´") &"','"& replace(request.Form("inpImgURL"),"'","´") &"',"& request.Form("mesas") &",'"& mensagem &"','"& replace(request.Form("inpImgURL2"),"'","´") &"',"& request.Form("ancora") &",'"& replace(request.Form("inpImgURL3"),"'","´") &"')")
	
	'Grava log
	Call gravaLog(session("user_rd"),"Projetos","Novo projeto","Rodada: "&replace(request.Form("local"),"'","´")&" - Assunto: "&replace(request.Form("assunto"),"'","´")&"")

%>
<script language="javascript" type="text/javascript">
        alert('Projeto salvo com sucesso!');
        window.location = "projetos.asp";
    </script> 
<%
	set novo = nothing
	response.End()
end if


'Alterar
if trim(request.Form("button")) = "ALTERAR" then
	
	dta_i = year(request.Form("dta_i")) &"-"& month(request.Form("dta_i")) &"-"& day(request.Form("dta_i"))
	dta_f = year(request.Form("dta_f")) &"-"& month(request.Form("dta_f")) &"-"& day(request.Form("dta_f"))
	hs_i = request.Form("hs_i")		
	hs_f = request.Form("hs_f")
	hs_ai = request.Form("hs_ai")
	hs_af = request.Form("hs_af")
	almoco = hs_ai & "-"& hs_af
	mensagem = Replace(request.Form("mensagem"),Chr(13), "<BR>")

	
	sql = "update tb_rodada set dta_ini='"&dta_i &" "& hs_i&"', dta_fim='"&dta_f &" "& hs_f&"', tempo="& request.Form("tempo") &", almoco='"& almoco &"', assunto='"& replace(request.Form("assunto"),"'","´") &"', intervalo='"& replace(request.Form("intervalo"),"'","´") &"', local='"& replace(request.Form("local"),"'","´") &"', manual='"& replace(request.Form("inpImgURL"),"'","´") &"', mesas="& request.Form("mesas") &", mensagem='"& mensagem &"', banner='"& replace(request.Form("inpImgURL2"),"'","´") &"', logo_emp='"& replace(request.Form("inpImgURL3"),"'","´") &"', ancora="& request.Form("ancora") &" where id_rod="&request.Form("id_rod")&" "
	set atualiza = adoConn.execute(sql)
	
	'Grava log
	Call gravaLog(session("user_rd"),"Projetos","Alteração de projetos","Rodada: "&replace(request.Form("local"),"'","´")&" - Assunto: "&replace(request.Form("assunto"),"'","´")&"")
	
	set atualiza = nothing
%>
	<script language="javascript" type="text/javascript">
        alert('Projeto atualizado com sucesso!');
        window.location = "projetos.asp";
    </script> 
<%
	response.End()
end if

'Deleta
if request.QueryString("exc") = 1 and request.QueryString("id") <> "" then
	set deleta = adoConn.execute("delete from tb_rodada where id_rod="&request.QueryString("id")&" ")
	set deleta = nothing
	set deleta = adoConn.execute("delete from tb_selecao where id_rod="&request.QueryString("id")&" ")
	set deleta = nothing
	'Grava log
	Call gravaLog(session("user_rd"),"Projetos","Exclusão de projetos","Rodada: "&replace(request.Form("local"),"'","´")&" - Assunto: "&replace(request.Form("assunto"),"'","´")&"")
%>
	<script language="javascript" type="text/javascript">
        alert('Projeto excluido com sucesso!');
        window.location = "projetos.asp";
    </script> 
<%
	response.End()
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
        <td scope="row" style="text-align:left"><h1><% if request("id") = "" then : response.Write("Novo ") : else response.Write("Alteração de ") : end if%>Projeto</h1></td>
        <td scope="row" style="text-align:left">&nbsp;</td>
      </tr>
      <tr>
        <td colspan="2" valign="top" scope="row">
          <form id="form1" name="form1" method="post" action="projetos_novo.asp">
            <table width="950" border="0" align="center" cellpadding="2" cellspacing="0">
              <tr>
                <td width="145" ><div align="right"><span style="font-family: Verdana, Arial, Helvetica, sans-serif; font-size: 12px; font-weight: bold">Data inicio:</span></div></td>
                <td colspan="2">
                  <table width="797" border="0" cellspacing="2" cellpadding="0">
                    <tr>
                      <td width="111" scope="row"><span id="sprytextfield1">
                  <input name="dta_i" type="text" class="campo" id="dta_i" onclick="displayCalendar(document.forms[0].dta_i,'dd/mm/yyyy',this)" onchange="dias()" size="10" tabindex="1"/>
                  <span class="textfieldRequiredMsg"><br />Campo obrigatorio.</span></span></td>
                      <td width="86"><div align="right"><span style="font-family: Verdana, Arial, Helvetica, sans-serif; font-size: 12px; font-weight: bold">Data fim:</span></div></td>
                      <td width="120"><span id="sprytextfield5">
                  <input name="dta_f" type="text" class="campo" id="dta_f" onclick="displayCalendar(document.forms[0].dta_f,'dd/mm/yyyy',this)" onchange="dias()" size="10" tabindex="2"/>
                  <span class="textfieldRequiredMsg"><br />Campo obrigatorio.</span></span></td>
                  <td width="480"><div id="n_dias"></div></td>
                    </tr>
                  </table></td>
                </tr>
              <tr>
                <td><div align="right"><span style="font-family: Verdana, Arial, Helvetica, sans-serif; font-size: 12px; font-weight: bold">Horário inicio:</span></div></td>
                <td colspan="2"><table width="797" border="0" cellspacing="2" cellpadding="0">
                    <tr>
                      <td width="111" scope="row"><span id="sprytextfield2">
                      <input name="hs_i" type="text" class="campo" id="hs_i" onchange="dias()" onKeyUp="maskIt(this,event,'##:##')" size="10" tabindex="3"/>
                      <span class="textfieldRequiredMsg"><br />
                      Campo obrigatorio.</span><span class="textfieldInvalidFormatMsg">Formato invalido.</span></span></td>
                      <td width="86"><div align="right"><span style="font-family: Verdana, Arial, Helvetica, sans-serif; font-size: 12px; font-weight: bold">Horário fim:</span></div></td>
                      <td width="120"><span id="sprytextfield4">
                      <input name="hs_f" type="text" class="campo" id="hs_f" onchange="dias()" onKeyUp="maskIt(this,event,'##:##')" size="10" tabindex="4"/>
                      <span class="textfieldRequiredMsg"><br />
                      Campo obrigatorio.</span><span class="textfieldInvalidFormatMsg">Formato invalido.</span></span></td>
                  <td width="480">&nbsp;</td>
                    </tr>
                  </table></td>
              </tr>
              <tr>
                <td><div align="right"><span style="font-family: Verdana, Arial, Helvetica, sans-serif; font-size: 12px; font-weight: bold">Local:</span></div></td>
                <td colspan="2"><span id="sprytextfield8">
                  <input type="text" name="local" id="local" style="width:350px;" tabindex="5" />
                  <span class="textfieldRequiredMsg">Campo obrigatorio.</span></span></td>
                </tr>
              <tr>
                <td><div align="right"><span style="font-family: Verdana, Arial, Helvetica, sans-serif; font-size: 12px; font-weight: bold">Assunto:</span></div></td>
                <td colspan="2"><span id="sprytextfield3">
                  <input type="text" name="assunto" id="assunto" style="width:350px;" tabindex="6"/>
                  <span class="textfieldRequiredMsg">Cargo Invalido.</span></span></td>
                </tr>
              <tr>
                <td><div align="right"><span style="font-family: Verdana, Arial, Helvetica, sans-serif; font-size: 12px; font-weight: bold">Duração da reunião:</span></div></td>
                <td colspan="2">
                    
				    <table width="100%" border="0" cellspacing="0" cellpadding="0">
				      <tr>
				        <td width="12%" scope="row"><span id="spryselect2"><select name="tempo" id="tempo" tabindex="7">
                      <option value="" selected="selected">Selecione</option>
                      <option value="10">10min</option>
                      <option value="20">20min</option>
                      <option value="30">30min</option>
                      <option value="40">40min</option>
                      <option value="50">50min</option>
                      <option value="60">60min</option>
                    </select>
                    <span class="selectRequiredMsg">Escolha a duração</span></span><span style="text-align:left">  </span></td>
				        <td width="19%"><div align="right"><span style="font-family: Verdana, Arial, Helvetica, sans-serif; font-size: 12px; font-weight: bold">Intervalo da  reunião:</span></div></td>
				        <td width="69%"><span id="spryselect3"><select name="intervalo" id="intervalo" tabindex="8">
                      <option value="" selected="selected">Selecione</option>
                      <option value="10">10min</option>
                      <option value="15">15min</option>
                      <option value="20">20min</option>
                      <option value="25">25min</option>
                      <option value="30">30min</option>
                    </select>
                    <span class="selectRequiredMsg">Escolha a duração</span></span><span style="text-align:left">  </span></td>
				        </tr>
				      </table></td>
                </tr>
              <tr>
                <td></td>
                <td colspan="2"></td>
                </tr>
              <tr>
                <td><div align="right"><span style="font-family: Verdana, Arial, Helvetica, sans-serif; font-size: 12px; font-weight: bold">Qtd de mesas:</span></div></td>
                <td colspan="2"><span id="spryselect4">
                  <select name="mesas" id="mesas" tabindex="9">
                    <option value="" selected="selected">Selecione</option>
                    <% for i = 1 to 50 %>
                    	<option value="<%=i%>"><%=i%></option>
                    <% next %>
                  </select>
                  <span class="selectRequiredMsg">Escolha a duração</span></span></td>
              </tr>
              <tr>
                <td><div align="right"><span style="font-family: Verdana, Arial, Helvetica, sans-serif; font-size: 12px; font-weight: bold">Duração do almoço:</span></div></td>
                <td colspan="2">
                <table width="797" border="0" cellspacing="2" cellpadding="0">
                    <tr>
                      <td width="111" scope="row"><span id="sprytextfield6">
                      <input name="hs_ai" type="text" class="campo" id="hs_ai" onchange="dias()" onKeyUp="maskIt(this,event,'##:##')" size="10" tabindex="10"/>
                      <span class="textfieldRequiredMsg"><br />
                      Campo obrigatorio.</span><span class="textfieldInvalidFormatMsg">Formato invalido.</span></span></td>
                      <td width="86"><div align="right"><span style="font-family: Verdana, Arial, Helvetica, sans-serif; font-size: 12px; font-weight: bold">Horário fim:</span></div></td>
                      <td width="120"><span id="sprytextfield7">
                      <input name="hs_af" type="text" class="campo" id="hs_af" onchange="dias()" onKeyUp="maskIt(this,event,'##:##')" size="10" tabindex="11"/>
                      <span class="textfieldRequiredMsg"><br />
                      Campo obrigatorio.</span><span class="textfieldInvalidFormatMsg">Formato invalido.</span></span></td>
                  <td width="480">&nbsp;</td>
                    </tr>
                  </table>
                </td>
              </tr>
              <tr>
                <td><div align="right"><span style="font-family: Verdana, Arial, Helvetica, sans-serif; font-size: 12px; font-weight: bold">Âncora:</span></div></td>
                <td>
                <span id="spryselect5">
                <select name="ancora" id="ancora" tabindex="9">
                  <option value="" selected="selected">Selecione</option>
                  <option value="4">Comprador</option>
                  <option value="3">Vendedor</option>
                </select><span class="selectRequiredMsg">Escolha o Âncora</span></span></td>
              </tr>
              <tr>
                <td><div align="right"><span style="font-family: Verdana, Arial, Helvetica, sans-serif; font-size: 12px; font-weight: bold">Anexar logo empresa:</span></div></td>
                <td colspan="2" valign="top"><table width="100%" border="0" cellspacing="2" cellpadding="0">
                  <tr>
                    <td width="24%"><input type="text" name="inpImgURL3" id="inpImgURL3" style="width:200px;"/> <br><span class="aviso"> 100px × 100px </span></td>
                    <td width="76%"><img src="images/icon_upload.png" alt="" width="25" height="25" style="cursor:pointer" onclick="javascript:upload(4);" style="cursor:pointer" /></td>
                  </tr>
                </table>
              </td>
                </tr>
              <tr>
                <td><div align="right"><span style="font-family: Verdana, Arial, Helvetica, sans-serif; font-size: 12px; font-weight: bold">Anexar banner:</span></div></td>
                <td colspan="2" valign="top"><table width="100%" border="0" cellspacing="2" cellpadding="0">
                  <tr>
                    <td width="24%"><input type="text" name="inpImgURL2" id="inpImgURL2" style="width:200px;"/><br><span class="aviso"> 733px × 238px </span></td>
                    <td width="76%"><img src="images/icon_upload.png" alt="" width="25" height="25" style="cursor:pointer" onclick="javascript:upload(1);" /></td>
                  </tr>
                </table></td>
              </tr>
              <tr>
                <td><div align="right"><span style="font-family: Verdana, Arial, Helvetica, sans-serif; font-size: 12px; font-weight: bold">Anexar manual:</span></div></td>
                <td colspan="2" valign="top"><table width="100%" border="0" cellspacing="2" cellpadding="0">
                  <tr>
                    <td width="24%"><input type="text" name="inpImgURL" id="inpImgURL" style="width:200px;"/></td>
                    <td width="76%"><img src="images/icon_upload.png" alt="" width="25" height="25" style="cursor:pointer" onclick="javascript:upload(2);" style="cursor:pointer" /></td>
                  </tr>
                </table>
              </td>
                </tr>
              <tr>
                <td>&nbsp;</td>
                <td valign="top">&nbsp;</td>
                <td valign="bottom">&nbsp;</td>
              </tr>
              <tr>
                <td><div align="right"><span style="font-family: Verdana, Arial, Helvetica, sans-serif; font-size: 12px; font-weight: bold">Mensagem convite:</span></div></td>
                <td colspan="2" valign="top"><textarea name="mensagem" id="mensagem" cols="90" rows="10"></textarea></td>
              </tr>
              <tr>
                <td>&nbsp;</td>
                <td width="193" valign="top">
                  <input type="submit" name="button" id="button" value="INCLUIR" tabindex="12"/>
                  <input type="hidden" name="id_rod" id="id_rod" />
                  </td>
                <td width="600" valign="bottom"><div id="iconExcluir"></div></td>
                </tr>
              <tr>
                <td colspan="3">&nbsp;</td>
              </tr>
            </table>
            <table width="100%" border="0" align="center" cellpadding="2" cellspacing="0">
              <tr> </tr>
              <tr> </tr>
              <tr> </tr>
              <tr> </tr>
            </table>
          </form></td>
      </tr>
    </table></td>
  </tr>
</table>
<script language="JavaScript" type="text/javascript">
  function upload(x){
    window.open('includes/uploadthumbnail.asp?tipo='+x+'','upload','width=480,height=200,status=yes,toolbar=no,scrollbars=yes,resizable=yes,navbar=no');
  }

<!--
var sprytextfield1 = new Spry.Widget.ValidationTextField("sprytextfield1", "none", {validateOn:["onchange"]});
var sprytextfield2 = new Spry.Widget.ValidationTextField("sprytextfield2", "time", {validateOn:["onchange"]});
var sprytextfield3 = new Spry.Widget.ValidationTextField("sprytextfield3", "none", {validateOn:["blur"]});
var sprytextfield4 = new Spry.Widget.ValidationTextField("sprytextfield4", "time", {validateOn:["onchange"]});
var sprytextfield5 = new Spry.Widget.ValidationTextField("sprytextfield5", "none", {validateOn:["onchange"]});
var sprytextfield6 = new Spry.Widget.ValidationTextField("sprytextfield6", "time", {validateOn:["onchange"]});
var sprytextfield7 = new Spry.Widget.ValidationTextField("sprytextfield7", "time", {validateOn:["onchange"]});

var sprytextfield8 = new Spry.Widget.ValidationTextField("sprytextfield8", "none", {validateOn:["blur"]});
var spryselect2 = new Spry.Widget.ValidationSelect("spryselect2", {validateOn:["blur"], isRequired:false});
var spryselect3 = new Spry.Widget.ValidationSelect("spryselect3", {validateOn:["blur"], isRequired:false});
var spryselect4 = new Spry.Widget.ValidationSelect("spryselect4", {validateOn:["blur"], isRequired:false});
var spryselect5 = new Spry.Widget.ValidationSelect("spryselect5", {validateOn:["blur"]});
//-->
</script>
<!-- #Include File="rodape.asp" --> 
</body>
</html>

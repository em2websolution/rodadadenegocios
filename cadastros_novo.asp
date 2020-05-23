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
					if (MTX_campos[0] == 'false') {
						altera(<%=session("user_rd")%>);
					} else {
						document.form1.id_usu.value = MTX_campos[0];
						document.form1.nome.value = MTX_campos[1];
						document.form1.email.value = MTX_campos[2];
						document.form1.senha.value = MTX_campos[3];
						document.form1.cargo.value = MTX_campos[4];
						document.form1.empresa.value = MTX_campos[5];
						document.form1.pais.value = MTX_campos[6];
						document.form1.telefone.value = MTX_campos[7];
						document.form1.celular.value = MTX_campos[8];
						document.form1.site.value = MTX_campos[9];
						document.form1.produtos.value = MTX_campos[10];
						document.form1.certificacao.value = MTX_campos[11];
						document.form1.nivel.value = MTX_campos[12];
						document.form1.stu.value = MTX_campos[13];
						document.form1.cpf.value = MTX_campos[14];					
						
						if (MTX_campos[15] == "Masculino") {
							document.form1.sexo[0].checked = true
						} else {
							document.form1.sexo[1].checked = true
						}
						document.form1.sexo.value = MTX_campos[15];					
						
						document.form1.emp_fantasia.value = MTX_campos[16];					
						document.form1.cnpj.value = MTX_campos[17];					
						document.form1.idioma.value = MTX_campos[18];					
						document.form1.emp_porte.value = MTX_campos[19];					
						document.form1.cep.value = MTX_campos[20];					
						document.form1.endereco.value = MTX_campos[21];					
						document.form1.complemento.value = MTX_campos[22];					
						document.form1.bairro.value = MTX_campos[23];
						document.form1.cidade.value = MTX_campos[24];					
						document.form1.uf.value = MTX_campos[25];										
						document.form1.obs.value = MTX_campos[26];										
						document.form1.polo.value = MTX_campos[27];					
						<% if session("lv_user") <> 1 then %>
							document.getElementById('txtpolo').innerHTML=MTX_campos[27];
						<% end if %>
						document.form1.perfil.value = MTX_campos[28];					
						
						document.form1.button.value = 'CONFIRMAR';
						
						<% if session("lv_user") <> 1 then %>
							switch(MTX_campos[12])
							{
							case "1":
							  document.getElementById('txtnivel').innerHTML="Administrador";
							  break;
							case "2":
							  document.getElementById('txtnivel').innerHTML="Gerente";
							  break;
							case "3":
							  document.getElementById('txtnivel').innerHTML="Vendedor";
							  break;
							case "4":
							  document.getElementById('txtnivel').innerHTML="Comprador";
							  break;
							}
						<% end if %>
					}
				} else {
					alert(ajax.responseText);
				}
			}
		}
		dados = 'id='+x;
		ajax.open('POST','includes/adm_pegaDclientes.asp',true);
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

if session("lv_user") = 1 or session("lv_user") = 2 then
	if request.QueryString("id_cad") <> "" then
		response.Write("<script language='javascript' type='text/javascript'>altera("&request.QueryString("id_cad")&");</script>")
	end if
	
end if

if session("lv_user") = 3 or session("lv_user") = 4 then
	response.Write("<script language='javascript' type='text/javascript'>altera("&session("user_rd")&");</script>")
end if

'Grava
if trim(request.Form("button")) = "CONCLUIR" then
	set novo = adoConn.execute("insert into tb_login (empresa, nivel, telefone, email, site, certificacao, pais, cargo, senha, produtos, nome, celular, stu, cpf, sexo, emp_fantasia, cnpj, idioma, emp_porte, cep, endereco, complemento, bairro, cidade, uf, obs, polo, perfil) values ('"& replace(request.Form("empresa"),"'","´") &"',"& request.Form("nivel") &",'"& replace(request.Form("telefone"),"'","´") &"','"& replace(request.Form("email"),"'","´") &"','"& replace(request.Form("site"),"'","´") &"','"& replace(request.Form("certificacao"),"'","´") &"','"& request.Form("pais") &"','"& replace(request.Form("cargo"),"'","´") &"','"& replace(request.Form("senha"),"'","´")&"','"& replace(request.Form("produtos"),"'","´") &"','"& replace(request.Form("nome"),"'","´") &"','"& replace(request.Form("celular"),"'","´") &"',"& request.Form("stu") &",'"& replace(request.Form("cpf"),"'","´") &"','"& request.Form("sexo") &"','"& replace(request.Form("emp_fantasia"),"'","´") &"','"& request.Form("cnpj") &"','"& request.Form("idioma") &"','"& replace(request.Form("emp_porte"),"'","´") &"','"& request.Form("cep") &"','"& replace(request.Form("endereco"),"'","´") &"','"& replace(request.Form("complemento"),"'","´") &"','"& replace(request.Form("bairro"),"'","´") &"','"& replace(request.Form("cidade"),"'","´") &"','"& request.Form("uf") &"','"& replace(request.Form("obs"),"'","´") &"','"& replace(request.Form("polo"),"'","´") &"','"& request.Form("perfil") &"' )")
	
	'Grava log
	Call gravaLog(session("user_rd"),"Cadastros","Novo cadastro","Empresa: "&replace(request.Form("empresa"),"'","´")&" - Nome: "&replace(request.Form("nome"),"'","´")&"")

%>
	<script language="javascript" type="text/javascript">
        alert('Cadastro salvo com sucesso!');
        window.location = "cadastros.asp?str=<%=request("str")%>";
    </script> 
<%
	set novo = nothing
	response.End()
end if


'Alterar
if trim(request.Form("button")) = "CONFIRMAR" then
	set atualiza = adoConn.execute("update tb_login set empresa='"& replace(request.Form("empresa"),"'","´") &"', nivel="& request.Form("nivel") &", telefone='"& replace(request.Form("telefone"),"'","´") &"', email='"& replace(request.Form("email"),"'","´") &"', site='"& replace(request.Form("site"),"'","´") &"', certificacao='"& replace(request.Form("certificacao"),"'","´") &"', pais='"& request.Form("pais") &"', cargo='"& replace(request.Form("cargo"),"'","´") &"', telefone='"& replace(request.Form("telefone"),"'","´") &"', senha='"& replace(request.Form("senha"),"'","´") &"', produtos='"& replace(request.Form("produtos"),"'","´") &"', nome='"& replace(request.Form("nome"),"'","´") &"', celular='"& replace(request.Form("celular"),"'","´") &"', stu="& request.Form("stu") &" , cpf='"& request.Form("cpf") &"', sexo='"& request.Form("sexo") &"', emp_fantasia='"& replace(request.Form("emp_fantasia"),"'","´") &"', cnpj='"& request.Form("cnpj") &"', idioma='"& replace(request.Form("idioma"),"'","´") &"', emp_porte='"& replace(request.Form("emp_porte"),"'","´") &"', cep='"& request.Form("cep") &"', endereco='"& replace(request.Form("endereco"),"'","´") &"', complemento='"& replace(request.Form("complemento"),"'","´") &"', bairro='"& replace(request.Form("bairro"),"'","´") &"', cidade='"& replace(request.Form("cidade"),"'","´") &"', uf='"& request.Form("uf") &"', obs='"& replace(request.Form("obs"),"'","´") &"', polo='"& replace(request.Form("polo"),"'","´") &"', perfil='"& replace(request.Form("perfil"),"'","´") &"' where id_usu="&request.Form("id_usu")&" ")


'response.Write("update tb_login set empresa='"& replace(request.Form("empresa"),"'","´") &"', nivel="& request.Form("nivel") &", telefone='"& replace(request.Form("telefone"),"'","´") &"', email='"& replace(request.Form("email"),"'","´") &"', site='"& replace(request.Form("site"),"'","´") &"', certificacao='"& replace(request.Form("certificacao"),"'","´") &"', pais='"& request.Form("pais") &"', cargo='"& replace(request.Form("cargo"),"'","´") &"', telefone='"& replace(request.Form("telefone"),"'","´") &"', senha='"& replace(request.Form("senha"),"'","´") &"', produtos='"& replace(request.Form("produtos"),"'","´") &"', nome='"& replace(request.Form("nome"),"'","´") &"', celular='"& replace(request.Form("celular"),"'","´") &"', stu="& request.Form("stu") &" , cpf='"& request.Form("cpf") &"', sexo='"& request.Form("sexo") &"', emp_fantasia='"& replace(request.Form("emp_fantasia"),"'","´") &"', cnpj='"& request.Form("cnpj") &"', idioma='"& replace(request.Form("idioma"),"'","´") &"', emp_porte='"& replace(request.Form("emp_porte"),"'","´") &"', cep='"& request.Form("cep") &"', endereco='"& replace(request.Form("endereco"),"'","´") &"', complemento='"& replace(request.Form("complemento"),"'","´") &"', bairro='"& replace(request.Form("bairro"),"'","´") &"', cidade='"& replace(request.Form("cidade"),"'","´") &"', uf='"& request.Form("uf") &"', obs='"& replace(request.Form("obs"),"'","´") &"', polo='"& replace(request.Form("polo"),"'","´") &"', perfil='"& replace(request.Form("perfil"),"'","´") &"' where id_usu="&request.Form("id_usu")&" ")


	'Grava log
	Call gravaLog(session("user_rd"),"Cadastros","Alteração de cadastro","Empresa: "&replace(request.Form("empresa"),"'","´")&" - Nome: "&replace(request.Form("nome"),"'","´")&"")
	
	set atualiza = nothing

if session("lv_user") = 1 or session("lv_user") = 2 then

%>
	<script language="javascript" type="text/javascript">
        alert('Cadastro atualizado com sucesso!');
        window.location = "cadastros.asp?str=<%=request("str")%>";
    </script> 
<% else %>
	<script language="javascript" type="text/javascript">
        alert('Cadastro atualizado com sucesso!');
        window.location = "projetos.asp";
    </script> 
<%
end if
	response.End()
end if

'Deleta
if request.QueryString("exc") = 1 and request.QueryString("id_cad") <> "" then
	set deleta = adoConn.execute("delete from tb_login where id_usu="&request.QueryString("id_cad")&" ")

	'Grava log
	Call gravaLog(session("user_rd"),"Cadastros","Exclusão de cadastro","Empresa: "&replace(request("empresa"),"'","´")&" - Nome: "&replace(request("nome"),"'","´")&"")

	set deleta = nothing
%>
	<script language="javascript" type="text/javascript">
        alert('Cadastro excluido com sucesso!');
        window.location = "cadastros.asp?str=<%=request("str")%>";
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
        <td scope="row" style="text-align:left"><h1>
			<% 
			if session("lv_user") = 1 or session("lv_user") = 2 then
				if request("id") = "" then : response.Write("Novo ") : else response.Write("Alteração de ") : end if
			else
				response.Write("Alteração de ") 
			end if			
			%>cadastro</h1></td>
        <td scope="row" style="text-align:left">&nbsp;</td>
      </tr>
      <tr>
        <td colspan="2" valign="top" scope="row">
          <form id="form1" name="form1" method="post" action="cadastros_novo.asp">
            <table width="979" border="0" align="center" cellpadding="2" cellspacing="0">
              <tr>
                <td width="169" ><div align="right"><strong>Razão Social<br />
                  <span class="phone" style="font-size:10px">Company Name</span><span class="phone" style="font-size:10px">:</span></strong></div></td>
                <td colspan="2"><span id="sprytextfield1">
                  <input type="text" name="empresa" id="empresa" style="width:350px;" tabindex="1"/>
                  <span class="textfieldRequiredMsg">Campo obrigatorio.</span></span></td>
                </tr>
              <tr>
                <td><div align="right"><strong>Nome Fantasia<br />
                  <span class="phone" style="font-size:10px">Fantasy Name</span><span class="phone" style="font-size:10px">:</span></strong></div></td>
                <td colspan="2"><input type="text" name="emp_fantasia" id="emp_fantasia" style="width:350px;" tabindex="2"/></td>
              </tr>
              <tr>
                <td><div align="right"><span style="font-family: Verdana, Arial, Helvetica, sans-serif; font-size: 12px; font-weight: bold">CNPJ:</span></div></td>
                <td colspan="2"><table width="97%" border="0" cellspacing="2" cellpadding="0">
                  <tr>
                    <td width="18%" scope="row"><span id="sprytextfield2">
                  <input type="text" name="cnpj" id="cnpj" style="width:140px;" tabindex="3" onKeyUp="maskIt(this,event,'##.###.###/####-##')"/>
                  <span class="textfieldRequiredMsg">Campo obrigatorio.</span></span></td>
                    <td width="7%">&nbsp;</td>
                    <td width="75%">&nbsp;</td>
                  </tr>
                </table></td>
              </tr>
              <tr>
                <td><div align="right"><strong>Endereço<br />
                  <span class="phone" style="font-size:10px">Address</span><span class="phone" style="font-size:10px">:</span></strong></div></td>
                <td colspan="2"><input type="text" name="endereco" id="endereco" style="width:350px;" tabindex="7"/></td>
              </tr>
              <tr>
                <td><div align="right"><strong>Complemento<br />
                  <span class="phone" style="font-size:10px">Complement</span><span class="phone" style="font-size:10px">:</span></strong></div></td>
                <td colspan="2"><table width="100%" border="0" cellspacing="2" cellpadding="0">
                  <tr>
                    <td width="17%" height="22" scope="row"><input type="text" name="complemento" id="complemento" style="width:140px;" tabindex="8"/></td>
                    <td width="6%"><div align="right"><span style="font-family: Verdana, Arial, Helvetica, sans-serif; font-size: 12px; font-weight: bold">UF:</span></div></td>
                    <td width="77%"><select name="uf" size="1"  tabindex="9">
                      <option selected="selected">Selecione</option>
                      <option value="AC"><b>AC</b></option>
                      <option value="AL"><b>AL</b></option>
                      <option value="AM"><b>AM</b></option>
                      <option value="AP"><b>AP</b></option>
                      <option value="BA"><b>BA</b></option>
                      <option value="CE"><b>CE</b></option>
                      <option value="DF"><b>DF</b></option>
                      <option value="ES"><b>ES</b></option>
                      <option value="GO"><b>GO</b></option>
                      <option value="MA"><b>MA</b></option>
                      <option value="MG"><b>MG</b></option>
                      <option value="MS"><b>MS</b></option>
                      <option value="MT"><b>MT</b></option>
                      <option value="PA"><b>PA</b></option>
                      <option value="PB"><b>PB</b></option>
                      <option value="PE"><b>PE</b></option>
                      <option value="PI"><b>PI</b></option>
                      <option value="PR"><b>PR</b></option>
                      <option value="RJ"><b>RJ</b></option>
                      <option value="RN"><b>RN</b></option>
                      <option value="RO"><b>RO</b></option>
                      <option value="RR"><b>RR</b></option>
                      <option value="RS"><b>RS</b></option>
                      <option value="SC"><b>SC</b></option>
                      <option value="SE"><b>SE</b></option>
                      <option value="SP"><b>SP</b></option>
                      <option value="TO"><b>TO</b></option>
                      </select></td>
                    </tr>
                  </table></td>
              </tr>
              <tr>
                <td><div align="right"><strong>Cidade<br />
                  <span class="phone" style="font-size:10px">City:</span></strong></div></td>
                <td colspan="2"><table width="97%" border="0" cellspacing="2" cellpadding="0">
                  <tr>
                    <td width="18%" scope="row"><input type="text" name="cidade" id="cidade" style="width:140px;" tabindex="10"/></td>
                    <td width="7%"><div align="right"><strong>Bairro/<span class="phone" style="font-size:10px">District</span>:</strong></div></td>
                    <td width="75%"><input type="text" name="bairro" id="bairro" style="width:140px;" tabindex="11"/></td>
                  </tr>
                </table></td>
              </tr>
              <tr>
                <td><div align="right"><strong>CEP<br />
                  <span class="phone" style="font-size:10px">ZIP:</span></strong></div></td>
                <td colspan="2"><input type="text" name="cep" id="cep" style="width:140px;" tabindex="6" onKeyUp="maskIt(this,event,'#####-###')"/></td>
              </tr>
              <tr>
                <td><div align="right"><strong>País<br />
                  <span class="phone" style="font-size:10px">Country:</span></strong></div></td>
                <td colspan="2"><span id="spryselect1">
                  <select name="pais" id="pais"  tabindex="12">
                    <option value="">Selecione</option>
                    
                    <option value="1">Afeganistão</option>
                    
                    <option value="2">África do Sul</option>
                    
                    <option value="3">Albânia</option>
                    
                    <option value="4">Alemanha</option>
                    
                    <option value="5">Andorra</option>
                    
                    <option value="6">Angola</option>
                    
                    <option value="9">Antígua e Barbuda</option>
                    
                    <option value="11">Arábia Saudita</option>
                    
                    <option value="12">Argélia</option>
                    
                    <option value="13">Argentina</option>
                    
                    <option value="14">Armênia</option>
                    
                    <option value="16">Austrália</option>
                    
                    <option value="17">Áustria</option>
                    
                    <option value="18">Azerbaijão</option>
                    
                    <option value="19">Bahamas</option>
                    
                    <option value="20">Bangladesh</option>
                    
                    <option value="21">Barbados</option>
                    
                    <option value="22">Barein</option>
                    
                    <option value="23">Bélgica</option>
                    
                    <option value="24">Belize</option>
                    
                    <option value="25">Benin</option>
                    
                    <option value="28">Bolívia</option>
                    
                    <option value="29">Bósnia-Herzegovina</option>
                    
                    <option value="30">Botsuana</option>
                    
                    <option value="31">Brasil</option>
                    
                    <option value="32">Brunei</option>
                    
                    <option value="33">Bulgária</option>
                    
                    <option value="34">Burkina Fasso</option>
                    
                    <option value="35">Burundi</option>
                    
                    <option value="36">Butão</option>
                    
                    <option value="37">Cabo Verde</option>
                    
                    <option value="38">Camarões</option>
                    
                    <option value="39">Camboja</option>
                    
                    <option value="40">Canadá</option>
                    
                    <option value="41">Catar</option>
                    
                    <option value="42">Cazaquistão</option>
                    
                    <option value="43">Chade</option>
                    
                    <option value="44">Chile</option>
                    
                    <option value="45">China</option>
                    
                    <option value="46">Chipre</option>
                    
                    <option value="48">Cingapura</option>
                    
                    <option value="49">Colômbia</option>
                    
                    <option value="50">Congo</option>
                    
                    <option value="51">Coréia do Norte</option>
                    
                    <option value="52">Coréia do Sul</option>
                    
                    <option value="53">Costa do Marfim</option>
                    
                    <option value="54">Costa Rica</option>
                    
                    <option value="55">Croácia</option>
                    
                    <option value="56">Cuba</option>
                    
                    <option value="57">Dinamarca</option>
                    
                    <option value="58">Djibuti</option>
                    
                    <option value="59">Dominica</option>
                    
                    <option value="60">Egito</option>
                    
                    <option value="61">El Salvador</option>
                    
                    <option value="62">Emirados Árabes</option>
                    
                    <option value="63">Equador</option>
                    
                    <option value="64">Eritréia</option>
                    
                    <option value="65">Eslováquia</option>
                    
                    <option value="66">Eslovênia</option>
                    
                    <option value="67">Espanha</option>
                    
                    <option value="68">Estados Unidos</option>
                    
                    <option value="69">Estônia</option>
                    
                    <option value="70">Etiópia</option>
                    
                    <option value="71">Fiji</option>
                    
                    <option value="72">Filipinas</option>
                    
                    <option value="73">Finlândia</option>
                    
                    <option value="74">França</option>
                    
                    <option value="75">Gabão</option>
                    
                    <option value="76">Gâmbia</option>
                    
                    <option value="77">Gana</option>
                    
                    <option value="78">Geórgia</option>
                    
                    <option value="80">Granada</option>
                    
                    <option value="81">Grécia</option>
                    
                    <option value="85">Guatemala</option>
                    
                    <option value="86">Guernsey</option>
                    
                    <option value="87">Guiana</option>
                    
                    <option value="88">Guiana Francesa</option>
                    
                    <option value="89">Guiné</option>
                    
                    <option value="91">Guiné Equatorial</option>
                    
                    <option value="90">Guiné-Bissau</option>
                    
                    <option value="92">Haiti</option>
                    
                    <option value="93">Holanda</option>
                    
                    <option value="94">Honduras</option>
                    
                    <option value="96">Hungria</option>
                    
                    <option value="97">Iêmen</option>
                    
                    <option value="114">Ilhas Marshall</option>
                    
                    <option value="116">Ilhas Salomão</option>
                    
                    <option value="123">Índia</option>
                    
                    <option value="124">Indonésia</option>
                    
                    <option value="125">Irã</option>
                    
                    <option value="126">Iraque</option>
                    
                    <option value="127">Irland</option>
                    
                    <option value="128">Islândia</option>
                    
                    <option value="129">Israel</option>
                    
                    <option value="130">Itália</option>
                    
                    <option value="131">Jamaica</option>
                    
                    <option value="132">Japão</option>
                    
                    <option value="134">Jordânia</option>
                    
                    <option value="135">Kiribati</option>
                    
                    <option value="136">Kuwait</option>
                    
                    <option value="137">Laos</option>
                    
                    <option value="138">Lesoto</option>
                    
                    <option value="139">Letônia</option>
                    
                    <option value="140">Líbano</option>
                    
                    <option value="141">Libéria</option>
                    
                    <option value="142">Líbia"</option>
                    
                    <option value="143">Liechtenstein</option>
                    
                    <option value="144">Lituânia</option>
                    
                    <option value="145">Luxemburgo</option>
                    
                    <option value="146">Macau</option>
                    
                    <option value="147">Macedônia</option>
                    
                    <option value="148">Madagascar</option>
                    
                    <option value="149">Malásia</option>
                    
                    <option value="150">Malauí</option>
                    
                    <option value="151">Maldivas</option>
                    
                    <option value="152">Mali</option>
                    
                    <option value="153">Malta</option>
                    
                    <option value="154">Marrocos</option>
                    
                    <option value="155">Martinica</option>
                    
                    <option value="156">Mauritânia</option>
                    
                    <option value="157">México</option>
                    
                    <option value="158">Mianmar</option>
                    
                    <option value="159">Micronésia</option>
                    
                    <option value="160">Moçambique</option>
                    
                    <option value="161">Moldávia</option>
                    
                    <option value="162">Mônaco</option>
                    
                    <option value="163">Mongólia</option>
                    
                    <option value="164">Montenegro</option>
                    
                    <option value="165">Montserrat</option>
                    
                    <option value="166">Namíbia</option>
                    
                    <option value="167">Nauru</option>
                    
                    <option value="168">Nepal</option>
                    
                    <option value="169">Nicarágua</option>
                    
                    <option value="170">Níger</option>
                    
                    <option value="171">Nigéria</option>
                    
                    <option value="172">Niue</option>
                    
                    <option value="173">Noruega</option>
                    
                    <option value="174">Nova Caledônia</option>
                    
                    <option value="175">Nova Zelândia</option>
                    
                    <option value="176">Omã</option>
                    
                    <option value="177">Palau</option>
                    
                    <option value="178">Palestinian Territories</option>
                    
                    <option value="179">Panamá</option>
                    
                    <option value="180">Papua-Nova Guiné</option>
                    
                    <option value="181">Paquistão</option>
                    
                    <option value="182">Paraguai</option>
                    
                    <option value="183">Peru</option>
                    
                    <option value="184">Pitcairn</option>
                    
                    <option value="185">Polinésia Francesa</option>
                    
                    <option value="186">Polônia</option>
                    
                    <option value="187">Porto Rico</option>
                    
                    <option value="188">Portugal</option>
                    
                    <option value="189">Quênia</option>
                    
                    <option value="190">Quirguistão</option>
                    
                    <option value="191">Reino Unido</option>
                    
                    <option value="192">República Centro-Africana</option>
                    
                    <option value="193">República Dominicana</option>
                    
                    <option value="194">República Tcheca</option>
                    
                    <option value="195">Romênia</option>
                    
                    <option value="196">Ruanda</option>
                    
                    <option value="27">Rússia</option>
                    
                    <option value="197">Rússia</option>
                    
                    <option value="198">Saara Ocidental</option>
                    
                    <option value="199">Saint-Pierre e Miquelon</option>
                    
                    <option value="200">Samoa</option>
                    
                    <option value="201">Samoa Americana</option>
                    
                    <option value="202">San Marino</option>
                    
                    <option value="203">Santa Helena</option>
                    
                    <option value="204">Santa Lúcia</option>
                    
                    <option value="205">São Cristóvão e Névis</option>
                    
                    <option value="206">São Tomé e Príncipe</option>
                    
                    <option value="207">São Vincente e Granadinas</option>
                    
                    <option value="208">Senegal</option>
                    
                    <option value="209">Serra Leoa</option>
                    
                    <option value="210">Sérvia</option>
                    
                    <option value="211">Sérvia e Montenegro</option>
                    
                    <option value="212">Síria</option>
                    
                    <option value="213">Somália</option>
                    
                    <option value="214">Sri Lanka</option>
                    
                    <option value="215">Suazilândia</option>
                    
                    <option value="216">Sudão</option>
                    
                    <option value="217">Suécia</option>
                    
                    <option value="218">Suíça</option>
                    
                    <option value="219">Suriname</option>
                    
                    <option value="220">Tailândia</option>
                    
                    <option value="221">Taiwan</option>
                    
                    <option value="222">Tajiquistão</option>
                    
                    <option value="223">Tanzânia</option>
                    
                    <option value="224">Timor-Leste</option>
                    
                    <option value="225">Togo</option>
                    
                    <option value="226">Tonga</option>
                    
                    <option value="227">Toquelau</option>
                    
                    <option value="228">Trinidad e Tobago</option>
                    
                    <option value="229">Tunísia</option>
                    
                    <option value="230">Turcomenistão</option>
                    
                    <option value="231">Turquia</option>
                    
                    <option value="232">Tuvalu</option>
                    
                    <option value="233">Ucrânia</option>
                    
                    <option value="234">Uganda</option>
                    
                    <option value="235">Uruguai</option>
                    
                    <option value="236">Uzbequistão</option>
                    
                    <option value="237">Vanuatu</option>
                    
                    <option value="47">Vaticano</option>
                    
                    <option value="238">Venezuela</option>
                    
                    <option value="239">Vietnã</option>
                    
                    <option value="240">Zâmbia</option>
                    
                    <option value="241">Zimbábue</option>
                    
                    </select>
                  <span class="selectRequiredMsg">Escolha o pais</span></span><span style="text-align:left">  </span>
                  
                    Language: 
                    <select name="idioma" id="idioma"  tabindex="12">
                    <option value="">Selecione</option>
                    
                    <option value="Português">Português</option>
                    
                    <option value="English">English</option>
                    
                    <option value="Spanish">Spanish</option>
                    </select>
                  </td>
              </tr>
              <tr>
                <td><div align="right"><strong>Nome<br />
                  <span class="phone" style="font-size:10px">Name:</span></strong></div></td>
                <td colspan="2"><span id="sprytextfield8">
                  <input type="text" name="nome" id="nome" style="width:350px;" tabindex="13"/>
                  <span class="textfieldRequiredMsg">Campo obrigatorio.</span></span></td>
              </tr>
              <tr>
                <td><div align="right"><strong>CPF<br />
                  <span class="phone" style="font-size:10px">Document:</span></strong></div></td>
                <td colspan="2"><table width="97%" border="0" cellspacing="2" cellpadding="0">
                  <tr>
                      <td width="13%" height="22" scope="row"><span id="sprytextfield3">
					<input type="text" name="cpf" id="cpf" style="width:110px;" tabindex="14" onKeyUp="maskIt(this,event,'###.###.###-##')"/>
                  <span class="textfieldRequiredMsg">Campo obrigatorio.</span></span></td>
                      <td width="6%"><div align="right"><span style="font-family: Verdana, Arial, Helvetica, sans-serif; font-size: 12px; font-weight: bold">Sexo:</span></div></td>
                      <td width="81%"><input type="radio" name="sexo" id="sexo" value="Masculino"  tabindex="15"/>
                        Masculino 
                        <input type="radio" name="sexo" id="sexo" value="Feminino"  tabindex="16"/>
                        Feminino</td>
                    </tr>
                </table></td>
              </tr>
              <tr>
                <td><div align="right"><strong>Cargo<br />
                  <span class="phone" style="font-size:10px">Title:</span></strong></div></td>
                <td colspan="2"><table width="100%" border="0" cellspacing="2" cellpadding="0">
                  <tr>
                      <td width="20%" scope="row"><span id="sprytextfield5">
                      <input type="text" name="cargo" id="cargo" style="width:150px;" tabindex="17"/>
                      <span class="textfieldRequiredMsg">Cargo Invalido.</span></span></td>
                      <td width="8%">&nbsp;</td>
                      <td width="72%">&nbsp;</td>
                    </tr>
                </table></td>
                </tr>
              <tr>
                <td><div align="right"><span style="font-family: Verdana, Arial, Helvetica, sans-serif; font-size: 12px; font-weight: bold">e-mail:</span></div></td>
                <td colspan="2"><span id="sprytextfield4">
                  <input type="text" name="email" id="email" style="width:350px;" tabindex="19"/>
                  <span class="textfieldRequiredMsg">e-mail invalido</span></span></td>
                </tr>
              <tr>
                <td><div align="right"><strong>Senha<br />
                  <span class="phone" style="font-size:10px">Password:</span></strong></div></td>
                <td colspan="2"><table width="100%" border="0" cellspacing="2" cellpadding="0">
                  <tr>
                      <td width="26%" scope="row">
                        <input type="text" name="senha" id="senha" style="width:150px;" tabindex="20"/>
                        <span class="aviso" style="font-size:10px">Recomendado alterar a senha padrão.</span> <span class="textfieldRequiredMsg">Campo obrigatorio.</span></td>
                      <td width="6%"><div align="right"><strong>Nivel<br />
                        <span class="phone" style="font-size:10px">Level:</span></strong></div></td>
                      <td width="68%"><% if session("lv_user") = 1 then %>  
                    <span id="spryselect2"><select name="nivel" id="nivel" tabindex="21">
                      <option value="" selected="selected">Selecione</option>
                      <option value="1">Administrador</option>
                      <option value="2">Gerente</option>
                      <option value="3">Vendedor</option>
                      <option value="4">Comprador</option>
                    </select><span class="selectRequiredMsg">Escolha o nível</span></span><span style="text-align:left">  </span>
				 <% else 
				 	
						if session("lv_user") = 2 then
							if request.QueryString("id_cad") = "" then
				 %>
                                <span id="spryselect2"><select name="nivel" id="nivel" tabindex="21">
                                  <option value="" selected="selected">Selecione</option>
                                  <option value="3">Vendedor</option>
                                  <option value="4">Comprador</option>
                                </select><span class="selectRequiredMsg">Escolha o nível</span></span><span style="text-align:left">  </span>
                     		<% else %>
                    			<input name="nivel" type="hidden" value="" /><div id="txtnivel"></div>
						<%   end if
						else %>
                            <input name="nivel" type="hidden" value="" /><div id="txtnivel"></div>
                 <% end if 
				 end if
				 %></td>
                    </tr>
                </table></td>
                </tr>
              <tr>
                <td><div align="right"><strong>Telefone<br />
                  <span class="phone" style="font-size:10px">Phone:</span></strong></div></td>
                <td colspan="2"><input type="text" name="telefone" id="telefone" style="width:350px;" onKeyUp="maskIt(this,event,'## (##) #####-####')" tabindex="22"/>                  
                  55 (11) XXXXX-XXXX</td>
              </tr>
              <tr>
                <td><div align="right"><strong>Celular<br />
                  <span class="phone" style="font-size:10px">Mobile:</span></strong></div></td>
                <td colspan="2"><input type="text" name="celular" id="celular" style="width:350px;" onKeyUp="maskIt(this,event,'## (##) #####-####')" tabindex="23"/>                  
                   55 (11) XXXXX-XXXX</td>
                </tr>
              <tr>
                <td><div align="right"><span style="font-family: Verdana, Arial, Helvetica, sans-serif; font-size: 12px; font-weight: bold">Site:</span></div></td>
                <td colspan="2"><input type="text" name="site" id="site" style="width:350px;" tabindex="24"/></td>
                </tr>
              <tr>
                <td><div align="right"><strong>Contato<br />
                  <span class="phone" style="font-size:10px">Contact:</span></strong></div></td>
                <td colspan="2" class="vermelho"><input type="text" name="obs" id="obs" style="width:350px;" tabindex="25"/></td>
              </tr>
              <tr>
                <td><div align="right"><span style="font-family: Verdana, Arial, Helvetica, sans-serif; font-size: 12px; font-weight: bold">Informação da Empresa<br />
                  <span class="phone" style="font-size:10px" id="result_box5" lang="en" xml:lang="en">Company Information:</span></span></div></td>
                <td colspan="2" class="vermelho"><textarea name="perfil" id="perfil" style="width:350px;" tabindex="26"></textarea></td>
              </tr>
              <tr>
                <td><div align="right"><strong>Segmento<br />
                  <span class="phone" style="font-size:10px">Segment:</span></strong></div></td>
                <td colspan="2" class="vermelho">
                <% if session("lv_user") = 1 then %>
                <select name="polo" id="polo"  tabindex="27">
                    <option value="" selected="selected">Selecione</option>
                    <option value="Agroindústria em geral">Agroindústria em geral</option>
                    <option value="Alimentos e Bebidas">Alimentos e Bebidas</option>
                    <option value="Apicultura">Apicultura</option>
                    <option value="Aquicultura">Aquicultura</option>
                    <option value="Arquitetura">Arquitetura</option>
                    <option value="Artesanato">Artesanato</option>
                    <option value="Calçados">Calçados</option>
                    <option value="Cerâmica">Cerâmica</option>
                    <option value="Confecções">Confecções</option>
                    <option value="Eletro Eletrônico">Eletro Eletrônico</option>
                    <option value="Gastronomia">Gastronomia</option> 
                    <option value="Gemas, Jóias e Assessórios">Gemas, Jóias e Assessórios</option>
                    <option value="Higiene Pessoal, Cosméticos e Perfumaria">Higiene Pessoal, Cosméticos e Perfumaria</option>
                    <option value="Investimento Imobiliário">Investimento Imobiliário</option>
                    <option value="Leite e Derivados">Leite e Derivados</option>
                    <option value="Madeireiro">Madeireiro</option>
                    <option value="Malacocultura">Malacocultura</option>
                    <option value="Metalmecânico">Metalmecânico</option>
                    <option value="Móveis">Móveis</option>
                    <option value="Náutico">Náutico</option>
                    <option value="Orgânicos">Orgânicos</option>
                    <option value="TI">TI</option>
                    <option value="Petróleo, Gás e Energia">Petróleo, Gás e Energia</option>
                    <option value="Plástico">Plástico</option>
                    <option value="Real State">Real State</option>
                    <option value="Suinocultura">Suinocultura</option>
                    <option value="Têxtil">Têxtil</option>
                    <option value="Turismo">Turismo</option>
                    <option value="Vitivinicultura">Vitivinicultura</option>
                </select>
				<% else %>
                <input name="polo" type="hidden" value="" /><div id="txtpolo"></div>
                <% end if %>
				</td>
                </tr>
              <tr>
                <td><div align="right"><span style="font-family: Verdana, Arial, Helvetica, sans-serif; font-size: 12px; font-weight: bold">Produtos de Interesse<br />
                  <span class="phone" style="font-size:10px" id="result_box6" lang="en" xml:lang="en">Products of Interest:</span></span></div></td>
                <td colspan="2"><textarea name="produtos" id="produtos" style="width:350px;" tabindex="26"></textarea></td>
                </tr>
              <tr>
                <td><div align="right"><strong>Certificação<br />
                  <span class="phone" style="font-size:10px" id="result_box7" lang="en" xml:lang="en">Certification</span><span class="phone" style="font-size:10px">:</span></strong></div></td>
                <td width="361" valign="top"><textarea name="certificacao" id="certificacao" style="width:350px;" tabindex="27"></textarea></td>
                <td width="408" valign="bottom">&nbsp;</td>
                </tr>
			<%if session("lv_user") = 1 or session("lv_user") = 2 then%>                
                  <tr>
                    <td><div align="right"><span style="font-family: Verdana, Arial, Helvetica, sans-serif; font-size: 12px; font-weight: bold">Status:</span></div></td>
                    <td><select name="stu" id="stu" tabindex="28">
                      <option value="1" selected="selected">Ativo</option>
                      <option value="0">Inativo</option>
                    </select></td>
                    <td>&nbsp;</td>
                  </tr>
            <% else %>
            	<input type="hidden" name="stu" id="stu" />
            <% end if %>
                    <tr>
                    <td>&nbsp;</td>
                    <td>
                      <input type="submit" name="button" id="button" value="CONCLUIR" tabindex="12"/>
                      <input type="hidden" name="id_usu" id="id_usu" />
                      <input type="hidden" name="str" id="str" value="<%=request("str")%>" />
                      <input type="hidden" name="emp_porte" id="emp_porte" style="width:140px;" tabindex="5"/>
                    </td>
                    <td>
                  </td>
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
<script type="text/javascript">
<!--
var sprytextfield1 = new Spry.Widget.ValidationTextField("sprytextfield1", "none", {validateOn:["blur"]});
var sprytextfield2 = new Spry.Widget.ValidationTextField("sprytextfield2", "none", {validateOn:["blur"]});
var sprytextfield3 = new Spry.Widget.ValidationTextField("sprytextfield3", "none", {validateOn:["blur"]});
var sprytextfield4 = new Spry.Widget.ValidationTextField("sprytextfield4", "email");
var sprytextfield5 = new Spry.Widget.ValidationTextField("sprytextfield5", "none", {validateOn:["blur"]});
var sprytextfield8 = new Spry.Widget.ValidationTextField("sprytextfield8", "none", {validateOn:["blur"]});
var spryselect1 = new Spry.Widget.ValidationSelect("spryselect1", {validateOn:["blur"], isRequired:false});
var spryselect2 = new Spry.Widget.ValidationSelect("spryselect2", {validateOn:["blur"], isRequired:false});
//-->
</script>
<!-- #Include File="rodape.asp" --> 
</body>
</html>

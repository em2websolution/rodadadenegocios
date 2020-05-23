<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<%Session.LCID = 1046%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title>Conceito Brazil - Rodada de Negócios</title>
<style type="text/css">
body {
	margin-left: 0px;
	margin-top: 0px;
}
</style>
<link rel="stylesheet" type="text/css" href="css/style_new.css"/>
<link rel="stylesheet" type="text/css" href="css/style_form.css"/>
<link type="text/css" rel="stylesheet" href="dhtmlgoodies_calendar/dhtmlgoodies_calendar.css?random=20051112" media="screen">
<SCRIPT type="text/javascript" src="dhtmlgoodies_calendar/dhtmlgoodies_calendar.js?random=20060118"></script>
<script src="https://kit.fontawesome.com/1c93fd985b.js" crossorigin="anonymous"></script>

<script type="text/javascript" language="javascript">
function maskIt(w,e,m,r,a){
    
    // Cancela se o evento for Backspace
    if (!e) var e = window.event
    if (e.keyCode) code = e.keyCode;
    else if (e.which) code = e.which;
    
    // Variáveis da função
    var txt  = (!r) ? w.value.replace(/[^\d]+/gi,'') : w.value.replace(/[^\d]+/gi,'').reverse();
    var mask = (!r) ? m : m.reverse();
    var pre  = (a ) ? a.pre : "";
    var pos  = (a ) ? a.pos : "";
    var ret  = "";

    if(code == 9 || code == 8 || txt.length == mask.replace(/[^#]+/g,'').length) return false;

    // Loop na máscara para aplicar os caracteres
    for(var x=0,y=0, z=mask.length;x<z && y<txt.length;){
        if(mask.charAt(x)!='#'){
            ret += mask.charAt(x); x++;
        } else{
            ret += txt.charAt(y); y++; x++;
        }
    }
    
    // Retorno da função
    ret = (!r) ? ret : ret.reverse()    
    w.value = pre+ret;
}

// Novo método para o objeto 'String'
String.prototype.reverse = function(){
    return this.split('').reverse().join('');
};

function Mudacor(x,y){
	if (x == 0) {
		document.getElementById(y).className = 'fundo2';
	}
	else {
		document.getElementById(y).className = 'fundo1';
	}
}

function exclui(x,y) {
	if (window.confirm("Deseja excluir o registo: "+x)) {
		window.location = y
	}
}

function fmtMoney(n, c, d, t){ 
	var m = (c = Math.abs(c) + 1 ? c : 2, d = d || ",", t = t || ".", 
	/(\d+)(?:(\.\d+)|)/.exec(n + "")), x = m[1].length > 3 ? m[1].length % 3 : 0; 
	return (x ? m[1].substr(0, x) + t : "") + m[1].substr(x).replace(/(\d{3})(?=\d)/g, 
	"$1" + t) + (c ? d + (+m[2] || 0).toFixed(c).substr(2) : ""); 
}; 

//-- função para validar o e-mail
function ValidaMail(valor){
	prim = valor.indexOf('@')
	if(prim < 2) return false;
	if(valor.indexOf('@',prim + 1) != -1) return false
	if(valor.indexOf('.') < 1) return false;
	if(valor.indexOf('zipmail.com') >= 0 && valor.indexOf('zipmail.com.br') == -1 && valor.indexOf('zipmeil.com') >= 0) return false;
	if(valor.indexOf('hotmail.com.br') >= 0 && valor.indexOf('hotmeil.com') >= 0) return false;
	if(valor.indexOf('.@') >= 0 && valor.indexOf('@.') >= 0) return false;
	if(valor.indexOf('.com.br.') >= 0 && valor.indexOf('/') >= 0) return false;
	if(valor.indexOf('[') >= 0 && valor.indexOf(']') > 0) return false;
	if(valor.indexOf('(') >= 0 && valor.indexOf(')') > 0) return false;
	if(valor.indexOf('..') >= 0) return false;
	if(valor.indexOf(';') >= 0) return false;
	return true;
}

function validaForm(){
	var form = document.form_top_content;

	var b_Avancar = document.getElementById('enviar');

	if(form.email.value==''){
		alert('O campo EMAIL não foi preenchido!');
		form.email.focus();
		return;
	}

	if(form.senha.value==''){
		alert('O campo SENHA não foi preenchido!');
		form.senha.focus();
		return;
	}
	form.submit();

}
</script> 
</head>

<body>



<!-- #Include File="includes/abre_conexao.asp" --> 


<%
mostra_contato = 1
'response.write("Session0: " & session("logo_emp") & "<br>" )

if request("param") = 1 then
	set adoRS = Server.CreateObject("ADODB.Recordset")
	adoRS.ActiveConnection = adoConn
	adoRS.Open("select * from tb_login where email='"& request("email") &"' and senha='"& request("senha") &"' and stu=1")
		
	if not adoRS.eof then
		session("user_rd") = int(adoRS("id_usu"))
		session("lv_user") = int(adoRS("nivel"))
		session("user_name") = adoRS("email")
		session("acesso") = adoRS("acesso")
		session("empresa") = adoRS("empresa")
		
		'Log de acesso
		set atualiza = adoConn.execute("update tb_login set acesso=now() where email='"& request("email") &"' and senha='"& request("senha") &"'")
		set atualiza = nothing

		'Grava um cookie na maquina do usuário    
		response.cookies("user_rd")("l") = request("email")
    response.cookies("user_rd")("s") = request("senha")
    
    if session("lv_user") > 1 then
      'Logo Empresa
      set pega_logo_emp = adoConn.execute("SELECT distinct R.logo_emp FROM conceitobrazil.tb_selecao as S inner join conceitobrazil.tb_login as L on S.id_usu = L.id_usu inner join conceitobrazil.tb_rodada as R on R.id_rod = S.id_rod where S.id_usu = '"& session("user_rd") &"' and R.logo_emp <> '' ORDER BY dta_ini desc")

      'response.write("SELECT distinct R.logo_emp FROM conceitobrazil.tb_selecao as S inner join conceitobrazil.tb_login as L on S.id_usu = L.id_usu inner join conceitobrazil.tb_rodada as R on R.id_rod = S.id_rod where S.id_usu = '"& session("user_rd") &"' and R.logo_emp <> '' ORDER BY dta_ini desc")

      if not pega_logo_emp.eof then
        session("logo_emp") = "arquivos/" & pega_logo_emp("logo_emp")
        mostra_contato = 0
      else
        session("logo_emp") = "images/logo_conceitobrazil-novo2.png"
      end if 
      set pega_logo_emp = nothing

      'response.write("Session login : " & session("logo_emp") & "<br>" )

      link_logo = "projetos.asp"
      response.Redirect("cadastros_novo.asp?id_cad="&session("user_rd")&"")
    else
      link_logo = "projetos.asp"

      'response.write("Session adm : " & session("logo_emp") & "<br>" )

      response.Redirect("projetos.asp")
    end if

    response.end
  else
    link_logo = "index.asp"
		session("user_rd") = empty
		session("lv_user") = 0
    session("user_name") = empty
    session("logo_emp") = empty
		session.Abandon()
  end if 
else 
if request("id") <> "" then
  if session("lv_user") > 1 or session("lv_user") = empty then
    link_logo = "projetos.asp"
    set adoRS = Server.CreateObject("ADODB.Recordset")
    adoRS.ActiveConnection = adoConn

    'Logo Empresa
    set pega_logo_emp = adoConn.execute("SELECT distinct R.logo_emp FROM conceitobrazil.tb_rodada as R where R.id_rod = '"& request("id") &"' and R.logo_emp <> '' ORDER BY dta_ini desc")
    if not pega_logo_emp.eof then
      session("logo_emp") = "arquivos/" & pega_logo_emp("logo_emp")
      mostra_contato = 0
    else
      session("logo_emp") = "images/logo_conceitobrazil-novo2.png"
    end if 
    set pega_logo_emp = nothing

    'response.write("Session  empresa : " & session("logo_emp") & "<br>" )
  end if
else
  if session("lv_user") <> empty then
    select case session("lv_user")
      case 1
        logo_emp = "images/logo_conceitobrazil-novo2.png"
        link_logo = "projetos.asp"
      case 2, 3, 4
        link_logo = "projetos.asp"
        mostra_contato = 0
      case else
        link_logo = "index.asp"
        logo_emp = "images/logo_conceitobrazil-novo2.png"
    end select
    'response.write("Session inicio 2 : " & session("logo_emp") & " - Nivel:" & session("lv_user") & "<br>" )
  else
      link_logo = "index.asp"
      session("logo_emp") = "images/logo_conceitobrazil-novo2.png"
      'response.write("Session1 : " & session("logo_emp") & "<br>" )
    end if
  end if
end if
%>



<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td scope="row" class="topo" id="topo">
      <table width="1052" border="0" align="center" cellpadding="0" cellspacing="0">
      <tr>
        <td valign="top" class="content_top" id="content_top" scope="row">
          <table width="1052" border="0" cellspacing="0" cellpadding="0">
          <tr>
            <th>
              <span class="content_top">
                <a href="<%=link_logo%>">
                  <br><img src="<%=session("logo_emp")%>" width="90px" height="90px" border="0" />
                </a>
              </span>
            </th>
            <td width="794" rowspan="2" valign="top"><table width="610" border="0" align="right" cellpadding="0" cellspacing="0">
              <tr>
                <td width="559" id="login_top" scope="row">
                  <% if trim(session("user_name")) = "" then %>
                <table width="488" border="0" cellpadding="0" cellspacing="2">
                   <form action="index.asp" method="post" accept-charset="utf-8" name="form_top_content" id="form_top_content" onsubmit="validaForm();return false;">
                  <tr>
                    <td width="69" scope="row">e-mail:</td>
                    <td width="203"><input name="email" type="text" id="email" accesskey="1" tabindex="1" size="28" />
                      <input name="param" type="hidden" id="param" value="1" /></td>
                    <td width="61">senha:</td>
                    <td width="82"><input name="senha" type="password" id="senha" accesskey="2" tabindex="2" size="12" /></td>
                    <td width="61"><input type="image" name="enviar" id="enviar" src="images/b_enviar.png" accesskey="3" tabindex="3" /></td>
                  </tr>
                  </form>
                </table>
				<% else 
					response.Write("Ola: " & session("user_name") & " - Último acesso: " & session("acesso") & " | <a href='sair.asp'>Sair</a>")
				              
                end if %>
                                </td>
              </tr>
              <tr>
                <td valign="top" id="login_top2" scope="row"><table width="100%" border="0" align="right" cellpadding="0" cellspacing="0">
                  <tr>
                    <td scope="row" ><div class="contato_top" id="contato_top">
                     <% if mostra_contato = 1 then %>
                      <ul class="phone">
                        <li><span class="fas fa-phone-volume fa-1x"><tel>11 3527-5000</tel></span></li>
                        <li><span class="fab fa-whatsapp fa-1x"></span> <a href="https://api.whatsapp.com/send?phone=5511992755644" target="_blank">(+55) 11 992755644</a></li>
                        <li><span class="fas fa-envelope fa-1x"></span> <a href="mailto:contato@conceitobrazil.com.br" target="_blank">contato@conceitobrazil.com.br</a></li>
                      </ul>
                    <% else %>
                      <h1>BEM VINDO A RODADA DE NEGÓCIOS</h1>                    
                    <% end if %>
                    </div></td>
                    <td width="58" scope="row" >&nbsp;</td>
                  </tr>
                </table></td>
              </tr>
            </table></td>
          </tr>
        </table></td>
      </tr>
      <tr>
        <td valign="top" scope="row">
        <% if session("lv_user") = 1 or session("lv_user") = 2 then %>
        <table width="980" border="0" align="center" cellpadding="0" cellspacing="2">
            <tr>
              <th width="61" scope="row"><a href="cadastros.asp"><img src="images/icon_cad.png" width="36" height="41" border="0" /></a></th>
              <td width="166" class="calendar_titulo"><a href="cadastros.asp">Cadastros</a></td>
              <th width="61"><img src="images/icon_projetos.png" width="30" height="41" border="0" /></th>
              <td width="754" class="calendar_titulo"><a href="projetos.asp">Projetos</a></td>
            </tr>
        </table>
        <% end if  
		if session("lv_user") = 3 or session("lv_user") = 4 then
		%>
        <table width="980" border="0" align="center" cellpadding="0" cellspacing="2">
            <tr>
              <th width="58" scope="row" ><a href="cadastros.asp"><img src="images/icon_cad.png" width="36" height="41" border="0" /></a></th>
              <td width="281" class="calendar_titulo" align="left"><a href="cadastros_novo.asp">Atualização de cadastro</a></td>
              <th width="80"><img src="images/icon_projetos.png" width="30" height="41" border="0" /></th>
              <td width="551" class="calendar_titulo"><a href="projetos.asp">Rodada de Negócios</a></td>
            </tr>
        </table>        
		<% end if %>
        </td>
      </tr>
    </table></td>
  </tr>
</table>

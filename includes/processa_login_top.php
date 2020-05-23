<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title>Conceito Brazil - Rodada de Negócios - Login Clientes</title>
<link href="../css/style_form.css" rel="stylesheet" type="text/css" />
<script type="text/javascript" language="javascript">
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
	form.action = 'processa_login_top.php';
	form.submit();

}
</script>  
</head>
<body>
<?
if ($_POST["email"]) {
	if ($_POST["email"]=='demonstracao@conceitobrazil.com.br' && $_POST["senha"]=='U483mv9#2e'){ ?>
		<script>window.parent.location='/abimo2012/welcome.php';</script>
	<? } else if ($_POST["email"]=='kenfoulstonicon.co.za' && $_POST["senha"]=='RbH4#92c74'){ ?>
		<script>window.parent.location='/compradores/welcome.php';</script>
	<? } else if ($_POST["email"]=='roberto@angelus.ind.br' && $_POST["senha"]=='bsD234e8#9'){ ?>
		<script>window.parent.location='/vendedores/welcome.php';</script>
<? }else{ ?>
		<script>alert('Login ou senha inválidos! Verifique os dados e tente novamente.')</script>
	<? }
}
?>
<form action="" method="post" accept-charset="utf-8" name="form_top_content" id="form_top_content" onsubmit="validaForm();return false;">e-mail:<input name="email" type="text" id="email" accesskey="1" tabindex="1" size="30" />&nbsp;&nbsp;&nbsp;senha:<input name="senha" type="password" id="senha" accesskey="2" tabindex="2" size="12" />&nbsp;&nbsp;&nbsp;<input type="image" name="enviar" id="enviar" src="../images/b_enviar.png" accesskey="3" tabindex="3" /></form>
</body>
</html>
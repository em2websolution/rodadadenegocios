<?php

//função para redirecionamento de URL..
function redireciona($link){
if ($link==-1){
	echo" <script>history.go(-1);</script>";
}else{
	echo" <script>document.location.href='$link'</script>";
}
}
$link = 'http://rodadadenegocios.conceitobrazil.com.br/index.asp'; // especifica o endereço
redireciona($link); // chama a função

?>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title>Conceito Brazil - Rodada de Negócios - Login Vendedores</title>
<link href="css/style.css" rel="stylesheet" type="text/css" />
</head>

<body>
<div class="main" id="main">
	<?php include('includes/top.php'); ?>
      <div class="content" id="content">
        <figure><img src="images/img_home.jpg" /></figure>
    <h1><br />
      PRÓXIMAS RODADAS</h1>
      <div class="pontilhado" id="pontilhado">
        <div class="calendar_data" id="calendar_data"><strong>Data</strong><br />
          28, 29 e 30/01/12<br />
        28, 29 e 30/01/13</div>
        <div class="calendar_local" id="calendar_local"><strong>Local</strong><br />
          Expo Center Norte SP<br />
          A definir</div>
        <div class="calendar_assunto" id="calendar_assunto"><strong>Assunto</strong><br />
          Rodada de Negócios ABIMO 2012<br />
          Rodada Demonstração ABIMO 2013
        </div>
      </div>
    </div>      
</div>
</body>
</html>

<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
</head>
<!-- #Include File="abre_conexao.asp" --> 
<%
	id_rod = int(request("id_rod"))

	set executa = Server.CreateObject("ADODB.Recordset")
	executa.ActiveConnection = adoConn

	if session("lv_user") <> 1 then
		str = " and L.nivel <> 1 "	
	else
		str = ""		
	end if
	

	select case int(request.QueryString("tipo"))
		case 1 'Indicados.asp - Completo
			sql = ("select  R.dta_ini, R.dta_fim, case L.nivel when 1 then 'Administrador' when 2 then 'Gerente' when 3 then 'Vendedor' when 4 then 'Comprador' END as Nivel, L.empresa, L.cnpj, L.emp_fantasia, L.nome, L.cpf, L.email, L.telefone, L.celular, L.perfil, I.dta_indicado, (select Di.emp_fantasia from conceitobrazil.tb_login as Di where I.id_indicado = Di.id_usu) as emp_indicada, (select Di.nome from conceitobrazil.tb_login as Di where I.id_indicado = Di.id_usu) as nome_indicada, (select Di.email from conceitobrazil.tb_login as Di where I.id_indicado = Di.id_usu) as email_indicada, (select Di.telefone from conceitobrazil.tb_login as Di where I.id_indicado = Di.id_usu) as telefone_indicada, (select Di.celular from conceitobrazil.tb_login as Di where I.id_indicado = Di.id_usu) as celular_indicada, (select Di.perfil from conceitobrazil.tb_login as Di where I.id_indicado = Di.id_usu) as perfil_indicada, I.hr FROM conceitobrazil.tb_indicacao as I inner join conceitobrazil.tb_login as L on I.id_usu = L.id_usu inner join conceitobrazil.tb_rodada as R on R.id_rod = I.id_rod where I.id_rod="&id_rod & str &" group by I.id_usu, I.id_ind")
			tipo = "Completo"
		case 2 'Indicados.asp - Resuido
			sql = ("select L.empresa, L.nome, L.email, L.polo, case L.nivel when 1 then 'Administrador' when 2 then 'Gerente' when 3 then 'Vendedor' when 4 then 'Comprador' END as Nivel, L.idioma, L.telefone, L.celular, L.perfil, (select Di.emp_fantasia from conceitobrazil.tb_login as Di where I.id_indicado = Di.id_usu) as emp_indicada, I.hr FROM conceitobrazil.tb_indicacao as I inner join conceitobrazil.tb_login as L on I.id_usu = L.id_usu inner join conceitobrazil.tb_rodada as R on R.id_rod = I.id_rod where I.id_rod="&id_rod & str &" group by I.id_usu, I.id_ind")
			tipo = "Resumido"
		case 3 'projetos_selecao.asp - Listagem
			sql = "select L.emp_fantasia, L.empresa, L.cnpj, L.nome, L.cargo, L.site, L.email, L.senha, L.endereco, L.complemento, L.bairro, L.cidade, L.uf, L.cep, L.polo, case L.nivel when 1 then 'Administrador' when 2 then 'Gerente' when 3 then 'Vendedor' when 4 then 'Comprador' END as Nivel, L.idioma, L.telefone, L.perfil from conceitobrazil.tb_selecao as S inner join conceitobrazil.tb_login as L on L.id_usu = S.id_usu where S.id_rod="&id_rod & str &""
			tipo = "Akna"
	end select		
	
	'response.Write(sql)
	
	executa.Open(sql)

If not executa.eof then

	Response.ContentType = "application/vnd.ms-excel"
	Response.AddHeader "Content-Disposition", "attachment; filename="&formatdatetime(now(),2)&"-"&tipo&".xls" 

%>
<TABLE BORDER=0 cellpadding="2" cellspacing="0" width="100%">
<TR>
<%
'Percorre cada campo e imprime o nome dos campos da tabela
For i = 0 to executa.fields.count - 1
%>
<th><% = executa(i).name %></th>
<% next %>
</TR>
<%
'Percorre cada linha e exibe cada campo da tabela

while not executa.eof
%>
<TR>
<% For i = 0 to executa.fields.count - 1
%>
<TD VALIGN=TOP><% = executa(i) %></TD>
<% Next %>
</TR>
<%
executa.MoveNext
wend
%>
</TABLE>
<% end if %>


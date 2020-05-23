<!-- #Include File="top.asp" -->
<script language="javascript" type="text/javascript">
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
        <td scope="row" style="text-align:left"><h1>Importação de cadastros</h1></td>
      </tr>
      <tr>
        <td scope="row">
          <table width="950" border="0" cellspacing="2" cellpadding="0">
            <tr>
              <td align="right">Importar</td>
              <td width="49" align="right"><img src="images/icon_excel.png" alt="Exportar XLS" width="49" height="42" border="0" onclick="javascript:upload(3);" style="cursor:pointer" /></td>
              </tr>
          </table>
        </td>
      </tr>
      <form id="form1" name="form1" method="post" action="">
      <tr>
        <td scope="row">
        
        
        <%  
 Dim sSourceXLS
	Dim sDestXLS
				

	sDestXLS = server.mappath(".") & "\arquivos\"& request("arq")   ' caminho e nome do arquivo xls
	Dim oConn
	Set oConn = Server.CreateObject("ADODB.Connection") 'conexao com o xls
	oConn.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & sDestXLS & ";Extended Properties=""Excel 8.0;HDR=NO;"""
	Dim oRS 
	Set oRS = Server.CreateObject("ADODB.Recordset") 'objeto recordset que armazena os dados do xls 'F1 = coluna 1 do xls, F2 = coluna 2, e assim sucessivamente...se for * seleciona todas as colunas
	
	sSQL = "SELECT * FROM [Plan1$]"
	set oRS = oConn.Execute(sSQL)

%>
<TABLE BORDER=0 cellpadding="2" cellspacing="0" width="950">
<tr>
  <td width="336">Vermelho = <span style="color:#FF0000">Já cadastrado</span> | Azul = <span style="color:#000066">Importado</span></td></tr>
<TR>
  <th VALIGN=TOP>NOME</th>
  <th width="263" VALIGN=TOP>E-MAIL</th>
  <th width="339" VALIGN=TOP>CARGOS</th>
</TR>
<TR>

<%
cont = 0
while not oRS.eof
cont = cont + 1
%>
<% 
lista = ""
grv = 0

For i = 0 to oRS.fields.count - 1

if oRS(0) <> "nome" and trim(oRS(0)) <> "" then 
	if grv = 0 then
		set verifica = adoConn.execute("select email from tb_login where email = '"& oRS(1) &"'")
		if verifica.eof then
			grv = 1
			bcolor = "#000066"
		else
			bcolor = "#FF0000"
		end if
		set verifica = nothing
	end if 

	select case  i
		case 0
	%>
			<TD VALIGN=TOP style="color:<%=bcolor%>">
				<% response.Write(cont-1 & " - "& oRS(i))%>
			</TD>
	<%		
		case 1
	%>
			<TD VALIGN=TOP style="color:<%=bcolor%>">
				<% response.Write(oRS(i))%>
			</TD>
	<%		
		case 3
	%>
			<TD VALIGN=TOP style="color:<%=bcolor%>">
			<% 		response.Write(oRS(i)) %>
			</TD>
	<%		
	end select
	
		if oRS(i) <> "" then
			str_dados = replace(oRS(i),"'","´")
		else
			str_dados = ""
		end if
		
		if i = 2 then
			lista = lista & str_dados &", "
		else
			lista = lista & "'"& str_dados &"', "
		end if
end if
Next %>
</TR>
<%
if oRS(0) <> "nome" and trim(oRS(0)) <> "" then 

	if grv = 1  then 
		lista = lista & "'123456',1"
		sql = ("insert into tb_login (nome, email, nivel, cargo, idioma, telefone, celular, site, cpf, sexo, empresa, emp_fantasia, cnpj, emp_porte, cep, endereco, complemento, bairro, cidade, uf, polo, perfil, obs, produtos, certificacao, senha, stu) values ("& lista &")")
		'response.Write(sql & "<br")
		'response.End()
		
		set novo = adoConn.execute(sql)
	end if
end if
oRS.MoveNext
wend
%>
</TABLE>
<%
 oRS.close
 set oRS=nothing
	oConn.Close
	Set oConn = Nothing
%>
        
        
        </td>
      </tr>
      </form>
    </table></td>
  </tr>
</table>
</body>
</html>

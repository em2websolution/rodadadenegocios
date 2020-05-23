<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!-- #Include File="includes/abre_conexao.asp" --> 
<%
id_rod = request("id_rod")
page = request("page")
str1 = request("str")

'----------------------------------------------------
'string ou DSN de conexão
'strDSN = "DSN=conceitobrazil"
strDSN = "Driver={MariaDB ODBC 3.1 Driver};SERVER=192.168.40.15;USER=conceitobrazil;PASSWORD=@9vRp7i5;DATABASE=conceitobrazil;PORT=3306"

'instrução SQL de pesquisa
		if trim(request("str")) <> "" then
			str = "empresa like '%"&request("str")&"%' or email like '%"&request("str")&"%' "
			if request("polo") <> "" then
				str = str & "and polo = '"&request("polo")&"'"
			end if
			strSQL = "select id_usu, empresa, nivel from conceitobrazil.tb_login where nivel in (3,4) and stu=1 or "&str&" order by nivel desc, empresa"		
		else
			if trim(request("polo")) <> "" then
				str = str & "and polo = '"&request("polo")&"'"
			else
				str = ""
			end if
			strSQL = "select id_usu, empresa, nivel from conceitobrazil.tb_login where nivel in (3,4) and stu=1 "&str&" order by nivel desc, empresa"		
		end if


'número de registros por página
intRecordsForPage = 500

'mensagem 
strMessageOne = "registro encontrado" 
strMessageMore = "registros encontrados" 
'---------------------------------------------------- 

set objConnection=server.createobject("adodb.connection") 
objConnection.open strDSN
set objRecordset = Server.CreateObject("adodb.recordset")
with objRecordset 
.Open strSQL,objConnection
MyList = .GetRows()
.Close
end with
set objRecordset = nothing
objConnection.close
set objConnection = nothing

intCol = cint(ubound(MyList,1))+1: intLin = cint(ubound(MyList,2))+1
boolMountTable = true
strMessage = "<b>" & intLin & "</b> " & strMessageOne
if intLin > 1 then strMessage = "<b>" & intLin & "</b> " & strMessageMore

intNumberThisPage=Request("page")
if intNumberThisPage = "" then intNumberThisPage = 1
if isnumeric(intNumberThisPage) = false then intNumberThisPage = 1
varTotalRecords = intLin 
intTotalPages = int(varTotalRecords/intRecordsForPage)
if intTotalPages < varTotalRecords/intRecordsForPage then intTotalPages = intTotalPages + 1

intLastRecord = intNumberThisPage * intRecordsForPage
intFirstRecord = intLastRecord - intRecordsForPage
if int(intTotalPages) = int(intNumberThisPage) then 
intLastRecord = varTotalRecords
end if
strTarget=Mid(Trim(Request.ServerVariables("PATH_INFO")),InstrRev(varLocal,"/")+1 )

set executa = Server.CreateObject("ADODB.Recordset")
executa.ActiveConnection = adoConn

for iList = intFirstRecord to intLastRecord-1
	set lista = Server.CreateObject("ADODB.Recordset")
	lista.ActiveConnection = adoConn
	lista.Open("select * from tb_selecao where id_rod="&id_rod&" and id_usu="&MyList(0,iList)&"")
	
	if lista.eof then
		executa.Open("insert into tb_selecao (id_rod,id_usu) values ("&id_rod&","&MyList(0,iList)&")")
	end if
	
	lista.close
	set lista = nothing	
next

response.Redirect("projetos_selecao.asp?id="&id_rod&"&page="&page&"&str="&str1&"&polo="&request("polo")&"")
%>

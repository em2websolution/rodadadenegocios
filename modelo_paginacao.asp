<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title>Untitled Document</title>
</head>

<body>
<%
'----------------------------------------------------
'string ou DSN de conexão
strDSN = "DSN=conceitobrazil"

'instrução SQL de pesquisa
strSQL = "select id_usu, empresa, nivel from tb_login where nivel in (3,4) and stu=1 order by nivel desc"

'número de registros por página
intRecordsForPage = 5

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

intNumberThisPage=1000

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
<table align=center border=1>
<tr>
<td bgcolor="black"><font color=white>Nro</font></td>
<td bgcolor="black"><font color=white>Nome da Coluna</font></td>
</tr> 
<%for iList = intFirstRecord to intLastRecord-1%><tr><td><%=iList+1%></td><td><%=MyList(0,iList)%></td></tr><%next%> 
</table>

<hr>
<table align=center border=1 width=100%>
<tr> 
<td align=left width="25%"><%=strMessage%></td>
<td align=center width="50%">
<%if intNumberThisPage <> 1 then%> 
<a href=<%=strTarget%>?page=1>primeira</a> 
| <a href=<%=strTarget%>?page=<%=intNumberThisPage-1%>>anterior</a> 
<%else%> 
primeira | anterior 
<%end if%><%if cLng(intNumberThisPage) = intTotalPages or boolMountTable = false then%> 
próxima | ultima 
<%else%> 
| <a href=<%=strTarget%>?page=<%=intNumberThisPage+1%>>próxima</a> 
| <a href=<%=strTarget%>?page=<%=intTotalPages%>>última</a> 
<%end if%> </td>

<td align=right width="25%"> 
<%if boolMountTable = true then%> 
página <b><%=intNumberThisPage%></b> de <b><%=intTotalPages%></b></td>
<%else%> 
&nbsp; 
<%end if%>
</td>
</tr>
</table>
</body>
</html>

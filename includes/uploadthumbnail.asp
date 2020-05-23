<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<!-- #include file="incUpload.asp" -->
<%

dim filename
up=request("up")
tipo=request("tipo")

if up<>"" then
		Set su = New FileUploader
		su.Upload()
		if su.Files.Count>0 then
			thisfile=su.Files("file").FileName
			thisfile=replace(thisfile,"-","_")
			filetitle=su.form("filetitle")
			if filetitle="" then filetitle=thisfile
			
			if thisfile<>"" then
				su.files("file").SaveToDisk(server.mappath("..") & "\arquivos\"& thisfile)
				button=su.form("button")
				select case tipo
					case 1
						campo = "inpImgURL2"
					case 2
						campo = "inpImgURL"
					case 3
						campo = ""
%>
					<script language="JavaScript">
                        window.opener.location = "../importacao.asp?arq=<%= thisfile%>";
                        window.close();
                    </script>
                 
<%		
          case 4
          campo = "inpImgURL3"
        end select
				
				if campo <> "" then
			
%>
					<script language="JavaScript">
                        window.opener.document.getElementById('<%=campo%>').value='<%= thisfile%>';
                        window.close();
                    </script>
                 
<%		
				end if
			end if
			
			set rs=nothing
		end if
		set su=nothing
end if
%>
</head>

<body bgcolor="#FFFFFF" text="#000000">
<form name="form1" method="post" ENCTYPE="multipart/form-data" action="uploadthumbnail.asp?up=1&tipo=<%=tipo%>">
  <table border="0" cellspacing="2" cellpadding="2" align="center">
    <tr> 
      <td><font face="Arial, Helvetica, sans-serif" size="2"><b> Upload Arquivo</b></font></td>
    </tr>
    <tr> 
      <td bgcolor="#666666" align="center"></td>
    </tr>
    <tr> 
      <td bgcolor="#F6F6F6"><font face="Arial, Helvetica, sans-serif" size="1">Selecione o arquivo que você deseja enviar e clique no botão de upload.
</font></td>
    </tr>
    <tr> 
      <td bgcolor="#666666" align="center"></td>
    </tr>
    <tr> 
      <td> <font face="Arial, Helvetica, sans-serif" size="2"> File : 
        <input type="file" name="file" size="30" style="">
        </font></td>
    </tr>
    <tr> 
      <td align="center"><font face="Arial, Helvetica, sans-serif" size="2">
        <input type="submit" name="Submit" value="Upload File">
        </font></td>
    </tr>
  </table>
</form>
</body>
</html>


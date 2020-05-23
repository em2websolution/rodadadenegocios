<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
</head>
<body> 
<%
		session("user_rd") = empty
		session("lv_user") = 0
		session("user_name") = empty
		session("logo_emp") = empty
		session.Abandon()
		response.Redirect("index.asp")
		response.End()
%>
</body>
</html>

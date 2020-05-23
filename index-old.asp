<!-- #Include File="top.asp" --> 
<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <th scope="row">&nbsp;</th>
  </tr>
	<% if session("lv_user") < 1 then %>
	<tr>
    <th scope="row"><img src="images/img_home.jpg" width="950" height="297" /></th>
  </tr>
<% end if %>
	<tr>
    <td scope="row"><table width="950" border="0" align="center" cellpadding="0" cellspacing="0">
      <tr>
        <td height="30" scope="row">&nbsp;</td>
      </tr>
      <tr>
        <td scope="row" style="text-align:left"><h1>PRÓXIMAS RODADAS</h1></td>
      </tr>
      <tr>
        <td scope="row"><table width="950" border="0" align="left" cellpadding="0" cellspacing="0">
          <tr class="calendar_titulo">
            <td width="344" align="left" class="borda" scope="row"><strong>Data</strong></td>
            <td width="345" align="left" class="borda" ><strong>Local</strong></td>
            <td width="261" align="left" class="borda" ><strong>Assunto</strong></td>
          </tr>
    <%	
	
sql = ("SELECT * from conceitobrazil.tb_rodada")

set dados = Server.CreateObject("ADODB.Recordset")
dados.ActiveConnection = adoConn
dados.Open(sql)
	
if not dados.eof then		
		y = 0
		while not dados.eof
		
	dta_i = day(dados("dta_ini")) &"/"& month(dados("dta_ini")) &"/"& year(dados("dta_ini"))
	dta_f = day(dados("dta_fim")) &"/"& month(dados("dta_fim")) &"/"& year(dados("dta_fim"))
	if hour(dados("dta_ini")) < 10 then 
		h1 = "0" & hour(dados("dta_ini"))
	else
		h1 = hour(dados("dta_ini"))
	end if

	if minute(dados("dta_ini")) < 10 then 
		m1 = "0" & minute(dados("dta_ini"))
	else
		m1 = minute(dados("dta_ini"))
	end if

	if hour(dados("dta_fim")) < 10 then 
		h2 = "0" & hour(dados("dta_fim"))
	else
		h2 = hour(dados("dta_fim"))
	end if

	if minute(dados("dta_fim")) < 10 then 
		m2 = "0" & minute(dados("dta_fim"))
	else
		m2 = minute(dados("dta_fim"))
	end if

	
	hs_i = h1 &":"& m1
	hs_f = h2 &":"& m2

	%>           
         
          <tr class="calendar_texto">
            <td scope="row" align="left" class="borda" ><%= (dta_i &" "& hs_i & "  à  " & dta_f &" "& hs_f)%></td>
            <td align="left" class="borda" ><%= dados("local")%></td>
            <td align="left" class="borda" ><%= dados("assunto")%></td>
          </tr>
    <%
		dados.movenext
		wend
		
		set dados = nothing
	%>
        </table>
    <% else 
		response.Write("Registro não encontrado")
		end if
	%>
</td>
      </tr>
    </table></td>
  </tr>
</table>
</body>
</html>

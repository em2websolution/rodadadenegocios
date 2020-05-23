<!-- #Include File="top.asp" --> 
<table width="100%" border="0" cellspacing="0" cellpadding="0">
	<% if session("lv_user") < 1 then %>
	<tr>
    <th scope="row"><img src="images/img_home2.png" width="950" height="203" /></th>
  </tr>
	<tr>
    <td scope="row">
			<div class="index-container">
					<ul style="list-style-type: none;">
						<li class="index-container-organizador fas fa-sitemap fa-4x">				
							<div>
								<br>Sou <br>Organizador
							</div>		
						</li>						
						<li class="index-container-comprador fas fa-shopping-cart fa-4x">				
							<div>
									<br>Sou <br>Comprador
							</div>
						</li>
						<li class="index-container-vendedor fas fa-dumpster fa-4x">				
							<div>
								<br>Sou <br>Vendedor
							</div>
						</li>
					</ul>
			</div>
		</td>
	</tr>
	<% else %>
	<tr>
    <th scope="row" style="height: 200px;">&nbsp;</th>
  </tr>
	<% end if %>
</table></td>
	</tr>
</table>
<!-- #Include File="rodape.asp" --> 
</body>
</html>

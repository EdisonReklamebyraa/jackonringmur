<html>

<head>
<title>HTTP Server Variables Collection</title>
</head>

<body>

<table border=1>
<tr>
	<td><b>Variabel Name</b></td><td><b>Value</b></td>
</tr>
<% For Each key In Request.ServerVariables %>
<tr>
	<td><%=key%>&nbsp;</td>
	<td>
	<% If Request.ServerVariables(key) = "" Then %>
		&nbsp;
	<% Else %>
		<%=Request.ServerVariables(key)%>
	<% End If %> 
	</td>
</tr>
<% Next %>
</table>

</body>

</html>
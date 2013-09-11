<!-- #Include file="adovbs.inc" -->

<%
Dim db
Set db = Server.CreateObject("ADODB.Connection")
db.Open "Driver={Microsoft Access Driver (*.mdb)}; DBQ=" & Request.Servervariables("APPL_PHYSICAL_PATH") & "jackon\nettcalc\DBNettCalc.mdb"
%>
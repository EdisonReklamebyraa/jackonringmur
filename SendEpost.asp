<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//NO" 	"http://www.w3.org/TR/html4/loose.dtd">

<HTML>

<HEAD>
<LINK REL="Stylesheet" TYPE="text/css" HREF="stil.css">
<TITLE>Jackon NettCalc - E-post bekreftelse</TITLE>
</HEAD>

<BODY>

<%
Dim Mailer, strHTML
Set Mailer = Server.CreateObject("SMTPsvg.Mailer")

With Mailer
	.RemoteHost  = "213.160.225.122"
	.FromName    = Request("N")
	.FromAddress = Request("Ep")
	.AddRecipient "Jackon", "jackon@jackon.no"
	.Subject     = "Materialliste / " & Request("N")
	.ContentType = "text/html"
	
	' Begynner å bygge opp HTML e-post
	strHTML = "<HTML><HEAD><STYLE TYPE=text/css> body	{ font-family: Arial; font-size: 9pt } table { font-family: Arial; font-size: 9pt } .storskrift { font-size: 14pt; font-weight: bold } .bunntekst { font-size: 8pt }</STYLE></HEAD><BODY>"
	strHTML = strHTML & "<P CLASS=storskrift>Materialliste for Jackon Ringmur inkl. gulv- og markisolasjon</P>"
	strHTML = strHTML & "<P><B>Byggeherre/Entreprenør:</B><BR>" & Request("N") & "<BR>" & Request("Ep") & "<BR>" & Request("Adr") & "<BR>" & Request("Ps") & "<BR>" & "Tlf./mob.: " & Request("Tel") & "<BR>" & "Faks: " & Request("Fax") & "</P>"
	strHTML = strHTML & "<P><TABLE><TR><TD COLSPAN=3><B>Byggeplass:</B></TD></TR><TR><TD WIDTH=100>Kommune:</TD><TD WIDTH=100>" & Request("Tk") & "</TD><TD>&nbsp;</TD></TR>"
	strHTML = strHTML & "<TR><TD>Adresse/tomt:</TD><TD>" & Request("Ta") & "</TD><TD>&nbsp;</TD></TR>"
	strHTML = strHTML & "<TR><TD>Areal:&nbsp;&nbsp;" & Request("A") & " m2</TD><TD>Omkrets:&nbsp;&nbsp;" & Request("P") & " m</TD><TD>Byggegrunn:&nbsp;&nbsp;" & Request("Byggegrunn") & "</TD></TR></TABLE></P>"	
	strHTML = strHTML & "<TABLE WIDTH=600><TR><TD WIDTH=300><U><I>Artikkelnavn</I></U></TD><TD WIDTH=100 ALIGN=Right><U><I>Antall</I></U></TD><TD WIDTH=200><U><I>Enhet</I></U></TD></TR>"
	strHTML = strHTML & Request("Artikler") & "</TABLE>"
	
	strHTML = strHTML & "<P>&nbsp;</P><DIV ALIGN=Center><P CLASS=bunntekst><HR SIZE=0,5>Jackon AS, Sørkilen 3, Postboks 1410, 1602 FREDRIKSTAD<BR>Tlf: 69 36 33 00, Fax: 800 32 874</P></DIV><P>&nbsp;</P></BODY></HTML>"
	
	.BodyText    = strHTML
End With

If Not Mailer.SendMail Then
  Response.Write "<p>Det oppstod en feil ved sending av e-post, bruk gjerne fax: 800 32 874. Feilen er: </p>"
  Response.Write "<p>" & Mailer.Response & "</p>"
  Response.Write "<p>Vennligst prøv igjen senere. Trykk Tilbake for å komme tilbake til materiallista.</p>"
Else
  Response.Write "<p>E-post ble sendt med suksess til Jackon. Trykk Tilbake for å komme tilbake til materiallista.</p>"
End if
%>

<FORM>
	<INPUT TYPE="button" VALUE="&nbsp;&nbsp;  Tilbake&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;" NAME="Tilbake" onClick="javascript:history.back()">
</FORM>

</BODY>

</HTML>
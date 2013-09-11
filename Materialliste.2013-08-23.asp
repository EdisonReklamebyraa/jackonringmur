<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//NO" 	"http://www.w3.org/TR/html4/loose.dtd">

<!-- #Include file="db.inc" -->

<%

' Hente brukerinformasjon
Dim N, Ep, Adr, Ps, Tf, Tk, Ta, Tiso
N = Request("N")
Ep = Request("Ep")
Adr = Request("Adr")
Ps = Request("Ps")
Tf = Request("Tf")
Ta = Request("Ta")
Tel = Request("Tel")
Fax = Request("Fax")
Tiso = Request("Tiso")

' Slår på error handling
On Error Resume Next

' Koble til kommuneliste
Dim rsTk
Set rsTk = Server.CreateObject("ADODB.Recordset")
sql = "Select * From Kommune"
rsTk.Open sql, db
rsTk.Filter = "KommuneNr = '" & Request("Tk") & "'"
Tk = rsTk("KommuneNavn")

' Sjekk av Error
If Err.Number <> 0 Then utputt = ByggErrorHTML("Kunne ikke filtrere på kommune: " & Request("Tk"),Err.number,Err.Description,Err.source)

' Hente byggeplass informasjon og beregne formler
Dim A, P, Hu, Hi, lambdag, Ms, A1, P1
'A1 = Replace(Request("A"),",",".")
'P1 = Replace(Request("P"),",",".")
A1 = Replace(Request("A"),".",",")
P1 = Replace(Request("P"),".",",")
A = CDbl(A1)
P = CDbl(P1)
Hu = Request("Hu")
Hi = Request("Hi")
'lambdag = Replace(Request("lambdag"),",",".")
lambdag = Request("lambdag")
Ms = Replace(Request("Ms"),",",".")

Dim nLambdag
Dim Byggegrunn, Yg
Select Case lambdag
Case "1,5"
	Byggegrunn = "Leire"
	Yg = "1,5"
Case "2,0"
	Byggegrunn = "Sand & Grus"
	Yg = "2,0"
Case "3,5"
	Byggegrunn = "Fjell"
	Yg = "3,5"
End Select

'lambdag = Replace(Request("lambdag"),",",".")
'nLambdag = CDbl(lambdag)

If IsNumeric(lambdag) = vbTrue Then
	nLambdag = CDbl(lambdag)
Else
	lambdag = TekstTilTall(lambdag)
	nLambdag = CDbl(lambdag)
End If

Dim Hjo, Rett, innHjo
Hjo = Hu * 2.65
Rett = AvrundOpp((((P*1.03)-Hjo)/2.4))*2.4
innHjo = Hi

Dim Plastkile
Dim Elasr
Elasr = 0
if (Request("Tr") = "RS" or Request("Tr") = "RU") then
	Plastkile = ((Rett/2.4) + (Hjo/2.65) + Hi) * 2
	'Elasr = (Cint(Hu) + Cint(Hi)) * 2
	Elasr = (Cint(Hu) + Cint(Hi))
else
	Elasr = (Rett/2.4) + (Hjo/2.65) + Hi
end if

Dim Fmas
Fmas = AvrundOpp(P/20)

Dim BI
If rsTk("Frostmengde") >= 50000 Or Request("Tr") = "R" Then
	BI = AvrundOpp(P/1.1)*1.1/1.1
End If

Dim stkSal
If IsNumeric(Ms) Then stkSal = AvrundOpp(Ms/1.2)

Dim Ti
If Request("Tr") = "R" Then
	Ti = (0.9385*Yg-0.1371*A/P+0.1073+0.0213*P/A+0.0582*Yg*P/A)/(0.0042*Yg-0.0039*Yg*P/A) + 25
Else
	Ti = (Yg-0.1371*A/P-0.0225-0.0765*Yg+0.0274+0.0045*P/A+0.153*Yg*P/A)/(0.0042*Yg-0.0008*Yg*P/A) - 9
End If
Ti = AvrundNed(Ti/10)*10

Dim Mi
Mi = A - (0.2*P)

Dim Plast
Plast = AvrundOpp(A/135)*135

Dim Tm, bmSidekant, BmHjorne, Mm
Tm = rsTk("Tm")
bmSidekant = rsTk("bmSidekant")
BmHjorne = rsTk("BmHjorne")
Mm = (P*bmSidekant) + (Hu*(3*(BmHjorne^2) - (2*(BmHjorne)*bmSidekant)))
Mm = AvrundOpp(Mm*10)/10


'' Start ===>> NY beregning av U-verdi
' Rg-verdi
Dim Rg
If (Tiso = 200) then
	Rg = 5.47
elseif (Tiso = 250) then
	Rg = 6.79
elseif (Tiso = 300) then
	Rg = 8.10
end if


' Areal grunn
Dim Ag
'=A-(0,25*P)+(Hu*0,0625)-Hi*0,0625
Ag = A - (0.25 * P) + (Hu * 0.0625) - Hi * 0.0625


' Omkrets innv
Dim Pg
'=P-(0,5*Hu)+(0,5*Hi)
Pg = P - (0.5 * Hu) + (0.5 * Hi)


' Jackopor 80, 
Dim Liso
Liso = 0.038

Dim B
' =Ag/(0,5*Pg)
B = Ag / (0.5 * Pg)

Dim Dt 
' =0,2+(lambdag*Rg)
Dt = 0.2 + (nlambdag * Rg)

Dim Uverdi
' =lambdag/(0,457*B+Dt)
Uverdi = nlambdag / (0.457 * B + Dt)
Uverdi = FormaterTallTilDesimal(Uverdi, 2)


'' Slutt ===>> NY beregning av U-verdi



' Koble til vareliste
Dim rsVare
Set rsVare = Server.CreateObject("ADODB.Recordset")
sql = "Select * From Varer"
rsVare.Open sql, db


' ***** GAMMEL KODE FOR OPPSETT AV GULVISOLASJONSLAG *****	
' ***** ENDRET 21.12.2009 RJ ********
'' Koble til gulvisolasjon
'Dim rsGI
'Set rsGI = Server.CreateObject("ADODB.Recordset")
'sql = "Select * From Gulvisolasjon"
'rsGI.Open sql, db

Dim strErrorHTML

Function ByggErrorHTML(feiltekst,feilnr,feilbeskr,feilkilde)
    
    'strErrorHTML = strErrorHTML & "<font color=red>"
    strErrorHTML = strErrorHTML & "Feil tekst: " & feiltekst & vbCrLf
    strErrorHTML = strErrorHTML & "<br/>Feil nummer: " & feilnr & vbCrLf
    strErrorHTML = strErrorHTML & "Feil beskrivelse: " & feilbeskr & vbCrLf
    strErrorHTML = strErrorHTML & "Feil kilde: " & feilkilde & vbCrLf         
    'strErrorHTML = strErrorHTML & "</font><br>"
    strErrorHTML = strErrorHTML & "<br>"
    err.clear 
End Function

%>

<HTML xmlns:v="urn:schemas-microsoft-com:vml" xmlns:o="urn:schemas-microsoft-com:office:office" xmlns="http://www.w3.org/TR/REC-html40">

<HEAD>
<LINK REL="Stylesheet" TYPE="text/css" href="stil.css">
<title>Jackon NettCalc - Materialliste</title>

<script LANGUAGE="JavaScript"> 
	<!--
	function Verifiser()
	{
		var strFeil;
		var intFeil=0;

		//if (document.Materialliste.Ep.value == "")
		//{
		//	strFeil = "Du kan ikke sende e-post hvis du ikke har registrert din e-postadresse.\n\n"
		//	strFeil = strFeil + "Trykk Tilbake for å fylle inn e-post adresse."
		//	intFeil = 1
		//}		

		if (intFeil == 1)
		{
			alert (strFeil);
			return false;
		}
		else
			return true;
	}
	-->
</SCRIPT>

</HEAD>

<BODY>

<% If strErrorHTML <> "" Then %>

    <p><b>Det oppsto en eller flere feil og materiallisten kunne ikke beregnes.</b></p><br />

    <p>Gå tilbake og forsøk igjen, og pass på at:</p><br />
    <ul>
        <li>- alle felter merket med * er fylt ut</li>
        <li>- både fylke og kommune er valgt</li>
        <li>- internett leseren ikke blokkerer javascript</li>
        <li>- det er skrevet inn tall og ikke tekst i tallfeltene</li>
    </ul> 

    <br />
    <p>Teknisk feilinformasjon: </p><br />

    <%=strErrorHTML%>
    

<% Else %>    

<FORM NAME="Materialliste" METHOD="Post" ACTION="SendEpost.asp" onSubmit="return Verifiser()">
  <TABLE WIDTH="606">
    <TR>
      <TD align="left" colspan="2">
      <TABLE width="600">
        <TR>
          <TD CLASS="storskrift" ALIGN="Left" width="614" colspan="2">Materialliste fra 
          Jackon AS</TD>
        </TR>
        <TR>
          <TD class="mellomskrift" align="left" valign="top" width="374">Jackon Ringmur inkl. gulv- og markisolasjon</TD>
          <TD width="216">
          <TABLE>
            <TR>
          <TD width="101" align="right">
          <p align="center">
          <INPUT TYPE="button" VALUE="&nbsp;&nbsp;  Tilbake&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;" NAME="Tilbake" onClick="javascript:history.back()"></TD>
          <TD width="113" align="right">
          <p align="left">til registrerings-skjema.</TD>
            </TR>
            <TR>
              <TD>
              <INPUT TYPE="button" VALUE="  &nbsp;&nbsp;Skriv ut&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;" NAME="SkrivUt" onClick="javascript:window.print()">
              </TD>
              <TD CLASS="bunntekst">
              <p align="left">og ta med til din byggevareforhandler. </TD>
            </TR>
              			<!--<TR><TD CLASS="bunntekst"><INPUT TYPE="Submit" VALUE="Send E-post" NAME="SendEpost"> 
                         </TD><TD CLASS="bunntekst"> til Jackon.</TD></TR>-->
              		</TABLE></TD>
        </TR>
      </TABLE></TD>
    </TR>
    <TR>
      <TD ALIGN="Left" colspan="2">
      <TABLE WIDTH="600" BORDER="0" style="border-left-width: 0px; border-right-width: 0px; border-top-width: 0px">
        <TR>
          <TD style="border-left-style:none; border-left-width:medium; border-right-style:none; border-right-width:medium; border-top-style:none; border-top-width:medium">&nbsp;</TD>
          <TD COLSPAN="2" style="border-left-style:none; border-left-width:medium; border-right-style:none; border-right-width:medium; border-top-style:none; border-top-width:medium">
          &nbsp;</TD>
        </TR>
        <TR>
          <TD style="border-top-style: solid; border-top-width: 1px"><B>Byggherre/Entreprenør:</TD>
          <TD COLSPAN="2" style="border-top-style: solid; border-top-width: 1px">
          <B>Byggeplass:</TD>
        </TR>
        <TR>
          <TD WIDTH="225" VALIGN="Top">
          <TABLE>
            <TR>
              <TD><%=N%><INPUT TYPE="Hidden" NAME="N" VALUE="<%=N%>">&nbsp;</TD>
            </TR>
            <TR>
              <TD><%=Ep%><INPUT TYPE="Hidden" NAME="Ep" VALUE="<%=Ep%>">&nbsp;</TD>
            </TR>
            <TR>
              <TD><%=Adr%><INPUT TYPE="Hidden" NAME="Adr" VALUE="<%=Adr%>">&nbsp;</TD>
            </TR>
            <TR>
              <TD><%=Ps%><INPUT TYPE="Hidden" NAME="Ps" VALUE="<%=Ps%>">&nbsp;</TD>
            </TR>
            <TR>
              <TD>Tel./mob.: <%=Tel%><INPUT TYPE="Hidden" NAME="Tel" VALUE="<%=Tel%>"></TD>
            </TR>
            <TR>
              <TD>Faks: <%=Fax%><INPUT TYPE="Hidden" NAME="Fax" VALUE="<%=Fax%>"></TD>
            </TR>
          </TABLE></TD>
          <TD WIDTH="225" VALIGN="Top">
          <TABLE>
            <TR>
              <TD WIDTH="100">Kommune:</TD>
              <TD WIDTH="100"><%=Tk%><INPUT TYPE="Hidden" NAME="Tk" VALUE="<%=Tk%>">
              </TR>
            <TR>
              <TD>Adresse/tomt:</TD>
              <TD><%=Ta%><INPUT TYPE="Hidden" NAME="Ta" VALUE="<%=Ta%>"></TD>
            </TR>
            <TR>
              <TD>Areal:</TD>
              <TD><%=A%> m2<INPUT TYPE="Hidden" NAME="Ta" VALUE="<%=A%>"></TD>
            </TR>
            <TR>
              <TD>Omkrets:</TD>
              <TD><%=P%> m<INPUT TYPE="Hidden" NAME="Ta" VALUE="<%=P%>"></TD>
            </TR>
            <TR>
              <TD>Byggegrunn:</TD>
              <TD><%=Byggegrunn%><INPUT TYPE="Hidden" NAME="Ta" VALUE="<%=Byggegrunn%>"></TD>
            </TR>
            <TR>
              <TD>Gulvisolasjon:</TD>
              <TD><%=Tiso%> mm<INPUT TYPE="Hidden" NAME="Ta" VALUE="<%=Tiso%>" ID="Hidden1"></TD>
            </TR>

          </TABLE></TD>
          <TD WIDTH="150" VALIGN="Top" align="right">
          <TABLE>
            <TR>
              <TD colspan="2">Dato: <%=FormaterDato(Date)%></TD>
            </TR>
          </TABLE>
          <p>&nbsp;</TD>
        </TR>
        <TR>
          <TD COLSPAN="3">&nbsp;</TD>
        </TR>
      </TABLE></TD>
    </TR>
    <TR>
      <TD ALIGN="Left" colspan="2" style="border-top-style: solid; border-top-width: 1px">
      <B>Jackon Ringmur:</B><br></TD>
    </TR>
    <TR>
      <TD ALIGN="Left" colspan="2">
      <TABLE WIDTH="600" border="0">
        <TR>
          <TD WIDTH="13%"><u><i>NOBB</i></u></TD>
          <TD style="width: 49%"><u><i>Varenavn</i></u></TD>
          <TD WIDTH="29%"><u><i>Beskrivelse</i></u></TD>
          <TD WIDTH="8%" ALIGN="Right"><u><i>Antall</i></u></TD>
          <TD WIDTH="8%"><u><i>Enhet</i></u></TD>
        </TR>
			<% SettInnRingmurselement %><INPUT TYPE="Hidden" NAME="Ringmurselement" VALUE="<%SettInnRingmurselement%>">
        <TR>
          <TD COLSPAN=5><br></TD>
        </TR>
      </TABLE></TD>
    </TR>
    <TR>
      <TD ALIGN="Left" colspan="2" style="border-top-style: solid; border-top-width: 1px">
      <B>Jackopor 80, Gulvisolasjon:</B><br></TD>
    </TR>
    <TR>
      <TD ALIGN="Left" colspan="2">
      <TABLE WIDTH="600" border="0">
        <TR>
          <TD WIDTH="13%"><u><i>NOBB</i></u></TD>
          <TD style="width: 49%"><u><i>Varenavn</i></u></TD>
          <TD WIDTH="29%"><u><i>Beskrivelse</i></u></TD>
          <TD WIDTH="8%" ALIGN="Right"><u><i>Antall</i></u></TD>
          <TD WIDTH="8%"><u><i>Enhet</i></u></TD>
        </TR>
			<% SettInnGulvisolasjon %><INPUT TYPE="Hidden" NAME="Gulvisolasjon" VALUE="<%SettInnGulvisolasjon%>">
        <TR>
          <TD COLSPAN=5><br></TD>
        </TR>
      </TABLE></TD>
    </TR>
    <TR>
      <TD ALIGN="Left" colspan="2" style="border-top-style: solid; border-top-width: 1px">
      <B>Jackofoam 200, Markisolasjon:</B><br></TD>
    </TR>
    <TR>
      <TD ALIGN="Left" colspan="2">
      <TABLE WIDTH="600" border="0">
        <TR>
          <TD WIDTH="13%"><u><i>NOBB</i></u></TD>
          <TD style="width: 49%"><u><i>Varenavn</i></u></TD>
          <TD WIDTH="29%"><u><i>Beskrivelse</i></u></TD>
          <TD WIDTH="8%" ALIGN="Right"><u><i>Antall</i></u></TD>
          <TD WIDTH="8%"><u><i>Enhet</i></u></TD>
        </TR>
			<% SettInnMarkisolasjon %><INPUT TYPE="Hidden" NAME="Markisolasjon" VALUE="<%SettInnMarkisolasjon%>">
        <TR>
          <TD COLSPAN=5><br></TD>
        </TR>
      </TABLE></TD>
    </TR>
    <TR>
      <TD align="left" width="316" style="border-top-style: solid; border-top-width: 1px">
      <b>Betongforbruk:</b></TD>
      <TD align="left" width="280" style="border-top-style: solid; border-top-width: 1px">
      <b>Armering:</b> </TD>
    </TR>
    <TR>
      <TD>
      <table border="1" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" width="94%" id="AutoNumber1" align="left">
        <tr>
          <td width="20%" align="center">Type</td>
          <td width="30%" align="center">Høyde (mm)</td>
          <td width="50%" align="center">Betong volum (liter/meter)</td>
        </tr>
        <tr>
          <td width="20%" rowspan="4" align="center">RS, RU, R</td>
          <td width="30%" align="center">300</td>
          <td width="50%" align="center">ca. 40</td>
        </tr>
        <tr>
          <td width="30%" align="center">450</td>
          <td width="50%" align="center">ca. 60</td>
        </tr>
        <tr>
          <td width="30%" align="center">600</td>
          <td width="50%" align="center">ca. 80</td>
        </tr>
        <tr>
          <td width="30%" align="center">750</td>
          <td width="50%" align="center">ca. 100</td>
        </tr>
        <tr>
          <td width="20%" align="center">Såleblokk</td>
          <td width="30%" align="center">250</td>
          <td width="50%" align="center">ca. 80</td>
        </tr>
      </table>
      <p><br><br><br><br><br><br>&nbsp;</TD>
      <TD>2 x 10mm kamjern oppe og nede. Ved bygging på 
        fjellgrunn er det tilstrekkelig med 2 jern oppe.<br><br><br><br><br>&nbsp;</TD>
    </TR>
    <TR>
      <TD colspan="2">
      <p align="center"><b>NB! </b>Det er ikke tatt hensyn til belastninger fra eventuelle 
        innvendige bærevegger og peisfundament. </TD>
    </TR>
    <TR>
      <TD ALIGN="Center" CLASS="bunntekst" colspan="2" style="border-top-style: solid; border-top-width: 1px">Jackon AS, Sørkilen 3, Postboks 1410, 1602 FREDRIKSTAD<BR>Tlf: 69 36 33 00, Fax: 800 32 874</TD>
    </TR>
  </TABLE>
  <% 'SkrivUtDebug %>
</FORM>

<% End If %>

<P>&nbsp;</P>

</BODY>

</HTML>

<%

Sub SettInnRingmurselement()

	Dim strHTML
	' Jackon Ringmur Rett
	If Rett > 0 Then
		rsVare.Filter = "Varekode = '" & Request("Tr") & "-" & Request("Hr") & "'"
		strHTML = strHTML & "<TR VALIGN=Top><TD>" & rsVare("NOBB") & "</TD><TD>"
		strHTML = strHTML & rsVare("Varenavn") & "</TD><TD>" & (Rett / 2.4) & " " & rsVare("Beskrivelse") & "</TD><TD ALIGN=Right>"
		strHTML = strHTML & Rett & "</TD><TD>"
		strHTML = strHTML & rsVare("Enhet") & "</TD></TR>"
		rsVare.Filter = ""
	End If

	' Jackon Ringmur Hjørne
	If Hjo > 0 Then
		rsVare.Filter = "Varekode = '" & Request("Tr") & "-" & Request("Hr") & "H'"
		strHTML = strHTML & "<TR VALIGN=Top><TD>" & rsVare("NOBB") & "</TD><TD>"
		strHTML = strHTML & rsVare("Varenavn") & "</TD><TD>" & (Hjo / 2.65) & " " & rsVare("Beskrivelse") & "</TD><TD ALIGN=Right>"
		strHTML = strHTML & Hjo & "</TD><TD>"
		strHTML = strHTML & rsVare("Enhet") & "</TD></TR>"
		rsVare.Filter = ""
	End If

	' Jackon Ringmur Innvendig Hjørne
	If innHjo > 0 Then
		rsVare.Filter = "Varekode = '" & Request("Tr") & "-" & Request("Hr") & "IH'"
		strHTML = strHTML & "<TR VALIGN=Top><TD>" & rsVare("NOBB") & "</TD><TD>"
		strHTML = strHTML & rsVare("Varenavn") & "</TD><TD>" & " " & rsVare("Beskrivelse") & "</TD><TD ALIGN=Right>"
		strHTML = strHTML & innHjo & "</TD><TD>"
		strHTML = strHTML & rsVare("Enhet") & "</TD></TR>"
		rsVare.Filter = ""
	End If

    If (Request("Tr") = "RS" or Request("Tr") = "RU") then
		rsVare.Filter = "Varekode = 'PLAST'"
		strHTML = strHTML & "<TR VALIGN=Top><TD>" & rsVare("NOBB") & "</TD><TD>"
		strHTML = strHTML & rsVare("Varenavn") & "</TD><TD>" & " " & rsVare("Beskrivelse") & "</TD><TD ALIGN=Right>"
		strHTML = strHTML & Plastkile & "</TD><TD>"
		strHTML = strHTML & rsVare("Enhet") & "</TD></TR>"
		rsVare.Filter = ""		
    end if


	' Elementlås Rett
	If Elasr > 0 Then
		rsVare.Filter = "Varekode = 'ELAS'"
		strHTML = strHTML & "<TR VALIGN=Top><TD>" & rsVare("NOBB") & "</TD><TD>"
		strHTML = strHTML & rsVare("Varenavn") & "</TD><TD>" & " " & rsVare("Beskrivelse") & "</TD><TD ALIGN=Right>"
		strHTML = strHTML & Elasr & "</TD><TD>"
		strHTML = strHTML & rsVare("Enhet") & "</TD></TR>"
		rsVare.Filter = ""
	End If

	' Fugemasse
	If Fmas > 0 Then
		rsVare.Filter = "Varekode = 'FMAS'"
		strHTML = strHTML & "<TR VALIGN=Top><TD>" & rsVare("NOBB") & "</TD><TD>"
		strHTML = strHTML & rsVare("Varenavn") & "</TD><TD>" & " " & rsVare("Beskrivelse") & "</TD><TD ALIGN=Right>"
		strHTML = strHTML & Fmas & "</TD><TD>"
		strHTML = strHTML & rsVare("Enhet") & "</TD></TR>"
		rsVare.Filter = ""
	End If

	' Jackofoam Bunnisolasjon
	If BI > 0 Then
		rsVare.Filter = "Varekode = 'BI-150'"
		strHTML = strHTML & "<TR VALIGN=Top><TD>" & rsVare("NOBB") & "</TD><TD>"
		strHTML = strHTML & rsVare("Varenavn") & "</TD><TD>" & " " & rsVare("Beskrivelse") & "</TD><TD ALIGN=Right>"
		strHTML = strHTML & BI & "</TD><TD>"
		strHTML = strHTML & rsVare("Enhet") & "</TD></TR>"
		rsVare.Filter = ""
	End If

	' Såleblokk
	If stkSal > 0 Then
		rsVare.Filter = "Varekode = 'SAL'"
		strHTML = strHTML & "<TR VALIGN=Top><TD>" & rsVare("NOBB") & "</TD><TD>"
		strHTML = strHTML & rsVare("Varenavn") & "</TD><TD>" & (stkSal) & " " & rsVare("Beskrivelse") & " = " & (stkSal * 1.2) & " m</TD><TD ALIGN=Right>"
		strHTML = strHTML & stkSal & "</TD><TD>"
		strHTML = strHTML & rsVare("Enhet") & "</TD></TR>"
		rsVare.Filter = ""
	End If
	
	' Skriver HTML koden
	Response.Write strHTML
		
End Sub

Sub SkrivUtDebug()

   strHTML = strHTML & "Lambdag = " & nLambdag & "<br>"

   strHTML = strHTML & "Tiso = " & Tiso & "<br>"
   strHTML = strHTML & "Rg = " & Rg  & "<br>"
   strHTML = strHTML & "Ag = " & Ag  & "<br>"
   strHTML = strHTML & "Pg = " & Pg  & "<br>"
   strHTML = strHTML & "Liso = " & Liso  & "<br>"
   strHTML = strHTML & "B = " & B  & "<br>"
   strHTML = strHTML & "Dt = " & Dt  & "<br>"
   strHTML = strHTML & "Uverdi = " & Uverdi  & "<br>"

   Response.Write strHTML

End Sub

Sub SettInnGulvisolasjon()

	Dim strHTML


' ***** GAMMEL KODE FOR OPPSETT AV GULVISOLASJONSLAG *****	
' ***** ENDRET 21.12.2009 RJ ********

'	' Henter lagvis inndeling av gulvisolasjon
'	rsGI.Filter = "Beregnet = '" & Ti & "'"
'	' KUN FOR TESTINGrsGI.Filter = "Beregnet = '" & 420 & "'"
	
	
	Dim bSkriv2linjer 
	bSkriv2linjer = false
	Dim nAntall
	nAntall = 0
	' Lag 1
	if (Tiso = 200) then
		rsVare.Filter = "Varekode = 'GI-100'"
		nAntall = Mi * 2
	elseif (Tiso = 250) then
		rsVare.Filter = "Varekode = 'GI-100'"
		nAntall = Mi
		bSkriv2linjer = true
	elseif (Tiso = 300) then
		rsVare.Filter = "Varekode = 'GI-150'"
		nAntall = Mi + Mi
	end if
	
	strHTML = strHTML & "<TR VALIGN=Top><TD>" & rsVare("NOBB") & "</TD><TD>"
	strHTML = strHTML & rsVare("Varenavn") & "</TD><TD>" & " " & rsVare("Beskrivelse") & "</TD><TD ALIGN=Right>"
	strHTML = strHTML & nAntall & "</TD><TD>"
	strHTML = strHTML & rsVare("Enhet") & "</TD></TR>"
	
	rsVare.Filter = ""
	if (bSkriv2linjer = true) then
		rsVare.Filter = "Varekode = 'GI-150'"
		strHTML = strHTML & "<TR VALIGN=Top><TD>" & rsVare("NOBB") & "</TD><TD>"
		strHTML = strHTML & rsVare("Varenavn") & "</TD><TD>" & " " & rsVare("Beskrivelse") & "</TD><TD ALIGN=Right>"
		strHTML = strHTML & nAntall & "</TD><TD>"
		strHTML = strHTML & rsVare("Enhet") & "</TD></TR>"
	end if
	
	rsVare.Filter = ""




' ***** GAMMEL KODE FOR OPPSETT AV GULVISOLASJONSLAG *****	
' ***** ENDRET 21.12.2009 RJ ********

'	' Setter inn første lag gulvisolasjon
'	rsVare.Filter = "Varekode = 'GI-" & rsGI("Lag1") & "'"
'	strHTML = strHTML & "<TR VALIGN=Top><TD>" & rsVare("NOBB") & "</TD><TD>"
'	strHTML = strHTML & rsVare("Varenavn") & "</TD><TD>" & " " & rsVare("Beskrivelse") & "</TD><TD ALIGN=Right>"
'	If rsGI("Lag1") = rsGI("Lag2") And rsGI("Lag2") = rsGI("Lag3") Then
'		strHTML = strHTML & (Mi * 3) & "</TD><TD>"
'	ElseIf rsGI("Lag1") = rsGI("Lag2") Then
'		strHTML = strHTML & (Mi * 2) & "</TD><TD>"
'	Else
'		strHTML = strHTML & Mi & "</TD><TD>"
'	End If
'	strHTML = strHTML & rsVare("Enhet") & "</TD></TR>"
'	rsVare.Filter = ""
'
'	' Setter inn andre lag gulvisolasjon hvis nødvendig
'	If rsGI("Lag2") <> 0 And rsGI("Lag1") <> rsGI("Lag2") Then
'		rsVare.Filter = "Varekode = 'GI-" & rsGI("Lag2") & "'"
'		strHTML = strHTML & "<TR VALIGN=Top><TD>" & rsVare("NOBB") & "</TD><TD>"
'		strHTML = strHTML & rsVare("Varenavn") & "</TD><TD>" & " " & rsVare("Beskrivelse") & "</TD><TD ALIGN=Right>"
'		If rsGI("Lag2") = rsGI("Lag3") Then
'			strHTML = strHTML & (Mi * 2) & "</TD><TD>"
'		Else
'			strHTML = strHTML & Mi & "</TD><TD>"
'		End If
'		strHTML = strHTML & rsVare("Enhet") & "</TD></TR>"
'		rsVare.Filter = ""
'	End If
'
'	' Setter inn tredje lag gulvisolasjon hvis nødvendig
'	If rsGI("Lag3") <> 0 And (rsGI("Lag2") <> rsGI("Lag3")) Then
'		rsVare.Filter = "Varekode = 'GI-" & rsGI("Lag3") & "'"
'		strHTML = strHTML & "<TR VALIGN=Top><TD>" & rsVare("NOBB") & "</TD><TD>"
'		strHTML = strHTML & rsVare("Varenavn") & "</TD><TD>" & " " & rsVare("Beskrivelse") & "</TD><TD ALIGN=Right>"
'		strHTML = strHTML & Mi & "</TD><TD>"
'		strHTML = strHTML & rsVare("Enhet") & "</TD></TR>"
'		rsVare.Filter = ""
'	End If


	' Setter inn samlet tykkelse
	strHTML = strHTML & "<TR VALIGN=Top><TD>&nbsp;</TD><TD>U-verdi = "
	strHTML = strHTML & Uverdi & " W/m2K</TD><TD>&nbsp;</TD><TD ALIGN=Right>"
	strHTML = strHTML & " " & "</TD><TD>"
	strHTML = strHTML & " " & "</TD></TR>"
	
	rsVare.Filter = "Varekode = 'PLAST-02'"
	strHTML = strHTML & "<TR VALIGN=Top><TD>" & rsVare("NOBB") & "</TD><TD>"
	strHTML = strHTML & rsVare("Varenavn") & "</TD><TD>" & (Plast/135) & " " & rsVare("Beskrivelse") & "</TD><TD ALIGN=Right>"
	strHTML = strHTML & Plast & "</TD><TD>"
	strHTML = strHTML & rsVare("Enhet") & "</TD></TR>"
	rsVare.Filter = ""
		
	' Gulvisolasjon
	'strHTML = strHTML & "<TR><TD COLSPAN=5><U><I>Gulvisolasjon, U-verdi = 0.15 W/m2K</I></U></TD></TR>"
	'strHTML = strHTML & "<TR><TD>"
	'strHTML = strHTML & "Jackopor 80, tykkelse = " & Ti & " mm, mengde = </TD><TD ALIGN=Right>"
	'strHTML = strHTML & (Mi) & "</TD><TD>"
	'strHTML = strHTML & "m2" & "</TD></TR>"

	' Skriver HTML koden

	Response.Write strHTML
	
End Sub

Sub SettInnMarkisolasjon()

	Dim strHTML

	' Markisolasjon, hvis byggegrunn ikke fjell eller fylke lik Finnmark
	If Not lambdag = "3,5" Then
		If Not Tf = "20" Then
			rsVare.Filter = "Varekode = 'MI-" & Tm & "'"
			strHTML = strHTML & "<TR VALIGN=Top><TD>" & rsVare("NOBB") & "</TD><TD>"
			strHTML = strHTML & rsVare("Varenavn") & "<BR>bredde = " & bmSidekant & " m / " & bmHjorne & " m</TD><TD>" & " " & rsVare("Beskrivelse") & "</TD><TD ALIGN=Right>"
			strHTML = strHTML & Mm & "</TD><TD>"
			strHTML = strHTML & rsVare("Enhet") & "</TD></TR>"
			rsVare.Filter = ""
			'strHTML = strHTML & "<TR><TD COLSPAN=5><U><I>Markisolasjon, benyttes ved telefarlig byggegrunn</I></U></TD></TR>"
			'strHTML = strHTML & "<TR><TD>"
			'strHTML = strHTML & "Jackofoam 200, tykkelse = " & Tm & " mm, mengde = </TD><TD ALIGN=Right>"
			'strHTML = strHTML & (Mm) & "</TD><TD>"
			'strHTML = strHTML & "m2" & "</TD></TR>"
			'strHTML = strHTML & "<TR><TD COLSPAN=5>bredde = " & bmSidekant & " m / " & bmHjorne & " m</TD></TR>"
		End If
	End If
	
	' Skriver HTML koden
	Response.Write strHTML
	
End Sub

Function AvrundOpp(Tall)

	If Int(Tall) <> Tall Then
		AvrundOpp = Int(Tall) + 1
	Else
		AvrundOpp = Int(Tall)
	End If
	
End Function

Function AvrundNed(Tall)

	AvrundNed = Int(Tall)
	
End Function

Function FormaterDato(Dato)

	Dag = Day(Dato)
	If Len(Dag) = 1 Then Dag = "0" & Dag
	Mnd = Month(Dato)
	If Len(Mnd) = 1 Then Mnd = "0" & Mnd
	Ar = Year(Dato)

	FormaterDato = Dag & "." & Mnd & "." & Ar
	
End Function

Function FormaterKlokke(Dato)

	Tim = Hour(Dato)
	If Len(Tim) = 1 Then Tim = "0" & Tim
	Min = Minute(Dato)
	If Len(Min) = 1 Then Min = "0" & Min
		
	FormaterKlokke = Tim & ":" & Min
	
End Function


Function FormaterTallTilDesimal(Tall, AntDes)
	
	FormaterTallTilDesimal = FormatNumber(Tall, AntDes)
End Function

Function TekstTilTall(innTekst)

	If InStr(innTekst,",") > 0 Then
		TekstTilTall = Replace(innTekst,",",".")
	ElseIf InStr(innTekst,".") > 0 Then
		TekstTilTall = Replace(innTekst,".",",")
	End If

End Function


%>
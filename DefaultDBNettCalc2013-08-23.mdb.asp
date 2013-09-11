<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//NO" 	"http://www.w3.org/TR/html4/loose.dtd">

<!-- #Include file="db.inc" -->

<SCRIPT LANGUAGE="JavaScript">

	var arrayKommuneNr = new Array();
	var arrayKommuneNavn = new Array();
	
	for(var i = 0; i < 3; i++)
	{
		
	}
</SCRIPT>

<%

Dim rsTf
Set rsTf = Server.CreateObject("ADODB.Recordset")
sql = "Select * From Fylke Order By FylkeNavn"
rsTf.Open sql, db

Dim rsTk
Set rsTk = Server.CreateObject("ADODB.Recordset")
sql = "Select KommuneNr, KommuneNavn From Kommune Order By KommuneNavn"
rsTk.Open sql, db

Dim intNr
intNr = 0

Do Until rsTk.EOF = True
	%>
	<SCRIPT LANGUAGE="JavaScript">
		arrayKommuneNr["<%=intNr%>"] = "<%=rsTk("KommuneNr")%>"
		arrayKommuneNavn["<%=intNr%>"] = "<%=rsTk("KommuneNavn")%>"
	</SCRIPT>
	<%
	intNr = intNr + 1
	rsTk.MoveNext
Loop

rsTk.MoveFirst

%>

<HTML>

<HEAD>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1" />
<LINK REL="Stylesheet" TYPE="text/css" HREF="stil.css">
<TITLE>Jackon NettCalc - Prosjektopplysninger</TITLE>

<SCRIPT LANGUAGE="JavaScript"> 
	<!--
	function Verifiser()
	{
		var strFeil;
		var intFeil=0;
		strFeil = "Du må fylle inn alle nødvendige data for å beregne materialliste:\n\n"
		if (document.Prosjekt.Tf.value == "")
		{
			strFeil = strFeil + "- Velg byggeplass fylke\n"
			intFeil = 1
		}		
		if (document.Prosjekt.Tk.value == "")
		{
			strFeil = strFeil + "- Velg byggeplass kommune\n"
			intFeil = 1
		}		
		if (document.Prosjekt.A.value == "")
		{
			strFeil = strFeil + "- Fyll inn bygningens areal\n"
			intFeil = 1
		}		
		if (document.Prosjekt.P.value == "")
		{
			strFeil = strFeil + "- Fyll inn bygningens omkrets\n"
			intFeil = 1
		}		
		if (document.Prosjekt.Hu.value == "")
		{
			strFeil = strFeil + "- Fyll inn antall utvendige hjørner\n"
			intFeil = 1
		}		
		if (document.Prosjekt.Hi.value == "")
		{
			strFeil = strFeil + "- Fyll inn antall innvendige hjørner\n"
			intFeil = 1
		}		
		if (document.Prosjekt.lambdag.value == "")
		{
			strFeil = strFeil + "- Velg byggegrunn\n"
			intFeil = 1
		}		
		if (document.Prosjekt.Hr.value == "")
		{
			strFeil = strFeil + "- Velg høyde ringmur\n"
			intFeil = 1
		}		
		if (document.Prosjekt.Tr.value == "")
		{
			strFeil = strFeil + "- Velg type ringmur\n"
			intFeil = 1
		}
        if (document.Prosjekt.Tiso.value == "") 
        {
            strFeil = strFeil + "- Velg gulvisolasjon\n"
        intFeil = 1
        }		
		
		if (intFeil == 1)
		{
			alert (strFeil);
			return false;
		}
		else
			return true;
	}

	function OppdaterKommuneListe(FylkeNr)
	{
	
		
        var oElem = document.getElementById("Tk");
  
		// Fjerne eksisterende verdier fra listen
		for (var q=oElem.options.length; q >=0; q--)
		{
			//alert('test');
			oElem.options[q] = null;
		}
		//	Prosjekt.Tk.options[q]=null;
		
		
  	

  		// Legge til et standardvalg		
  		addOption('Tk', "-Velg kommune-", "");
  	
  		// Loope gjennom array for å legge til kommuner for fylket
  		for ( x = 0 ; x < arrayKommuneNr.length  ; x++ )
    	{
      		if ( arrayKommuneNr[x].substr(0,2) == FylkeNr )
        		{
        		addOption('Tk', arrayKommuneNavn[x], arrayKommuneNr[x]);
        		}
    	}
	}
	

function addOption(elementRef, optionText, optionValue) 
{
 if ( typeof elementRef == 'string' )
  elementRef = document.getElementById(elementRef);

 if ( elementRef.type.substr(0, 6) != 'select' )
  return;

 var newOption = document.createElement('option');

 newOption.appendChild(document.createTextNode(optionText));

 if ( arguments.length >= 3 )
  newOption.setAttribute('value', optionValue);

 elementRef.appendChild(newOption);
}


	-->
</SCRIPT>

</HEAD>

<BODY>
<img src="http://www.jackon.no/eway/custom/design/jackon/Jackon_NYlogo.jpg" />
        <h2 class="headerText">Jackon NettCalc</h2>
<FORM NAME="Prosjekt" METHOD="post" ACTION="materialliste.asp" onSubmit="return Verifiser()">
	<TABLE>
		<TR>
			<TD><label for="N">Navn:</label></TD>
			<TD><INPUT TYPE="Text" NAME="N" ID="N" SIZE="40"></TD>
		</TR>
		<TR>
			<TD>E-post:</TD>
			<TD><INPUT TYPE="Text" NAME="Ep" SIZE="40"></TD>
		</TR>
		<TR>
			<TD>Adresse:</TD>
			<TD><INPUT TYPE="Text" NAME="Adr" SIZE="40"></TD>
		</TR>
		<TR>
			<TD>Postnr/Sted:</TD>
			<TD><INPUT TYPE="Text" NAME="Ps" SIZE="40"></TD>
		</TR>
		<TR>
			<TD>Telefon/mobil:</TD>
			<TD><INPUT TYPE="Text" NAME="Tel" SIZE="40"></TD>
		</TR>
		<TR>
			<TD>Telefaks:</TD>
			<TD><INPUT TYPE="Text" NAME="FAX" SIZE="40"></TD>
		</TR>
		<TR>
			<TD>Byggeplass, fylke: (*)</TD>
			<TD>
				<SELECT NAME="Tf" ID="Tf" STYLE="font-family: Arial; font-size: 9pt" onChange="OppdaterKommuneListe(this.value)">
				<OPTION SELECTED>-Velg fylke-</OPTION>
				<% Do Until rsTf.EOF %>
					<OPTION VALUE="<%=rsTf("FylkeNr")%>"><%=rsTf("FylkeNavn")%></OPTION>
					<% rsTf.MoveNext %>
				<% Loop %>
				</SELECT>
			</TD>
		</TR>
		<TR>
			<TD>Byggeplass, kommune: (*)</TD>
			<TD>
				<SELECT NAME="Tk" ID="Tk" STYLE="font-family: Arial; font-size: 9pt">
				<OPTION SELECTED>-Velg kommune-</OPTION>
				<!-- <% Do Until rsTk.EOF %>
					<OPTION VALUE="<%=rsTk("KommuneNr")%>"><%=rsTk("KommuneNavn")%></OPTION>
					<% rsTk.MoveNext %>
				<% Loop %> -->
			    </SELECT>
			</TD>
		</TR>
		<TR>
			<TD>Adresse, byggeplass/tomt:</TD>
			<TD><INPUT TYPE="Text" NAME="Ta" SIZE="40"></TD>
		</TR>
		<TR>
			<TD>Bygningens areal: (*)</TD>
			<TD><INPUT TYPE="Text" NAME="A" SIZE="10">&nbsp;m2</TD>
		</TR>
		<TR>
			<TD>Bygningens omkrets: (*)</TD>
			<TD><INPUT TYPE="Text" NAME="P" SIZE="10">&nbsp;m</TD>
		</TR>
		<TR>
			<TD>Antall utvendige hjørner: (*)</TD>
			<TD><INPUT TYPE="Text" NAME="Hu" SIZE="10">&nbsp;stk</TD>
		</TR>
		<TR>
			<TD>Antall innvendige hjørner: (*)</TD>
			<TD><INPUT TYPE="Text" NAME="Hi" SIZE="10">&nbsp;stk</TD>
		</TR>
		<TR>
			<TD>Byggegrunn: (*)</TD>
			<TD>
				<SELECT NAME="lambdag" ID="lambdag" STYLE="font-family: Arial; font-size: 9pt">
				<OPTION SELECTED>-Velg-</OPTION>
				<OPTION VALUE="1,5">Leire</OPTION>
				<OPTION VALUE="2,0">Sand & Grus</OPTION>
				<OPTION VALUE="3,5">Fjell</OPTION>
			    </SELECT>
				&nbsp;W/mK</TD>
		</TR>	
		<TR>
			<TD>Høyde ringmur: (*)</TD>
			<TD>
				<SELECT NAME="Hr" ID="Hr" STYLE="font-family: Arial; font-size: 9pt">
				<OPTION SELECTED>-Velg-</OPTION>
				<OPTION VALUE="300">300</OPTION>
				<OPTION VALUE="450">450</OPTION>
				<OPTION VALUE="600">600</OPTION>
				<OPTION VALUE="750">750</OPTION>
			    </SELECT>
			    &nbsp;mm
			</TD>
		</TR>	
		<TR>
			<TD>Type ringmur: (*)</TD>
			<TD>
				<SELECT NAME="Tr" ID="Tr" STYLE="font-family: Arial; font-size: 9pt">
				<OPTION SELECTED>-Velg-</OPTION>
				<OPTION VALUE="RS">RS</OPTION>
				<OPTION VALUE="RU">RU</OPTION>
				<OPTION VALUE="R">R</OPTION>
			    </SELECT>
			</TD>
		</TR>
		
		<TR>
			<TD>Jackopor 80 gulvisolasjon: (*)</TD>
			<TD>
				<SELECT NAME="Tiso" ID="Tiso" STYLE="font-family: Arial; font-size: 9pt">
				<OPTION SELECTED>-Velg-</OPTION>
				<OPTION VALUE="200">200</OPTION>
				<OPTION VALUE="250">250</OPTION>
				<OPTION VALUE="300">300</OPTION>
			    </SELECT>
			    &nbsp;mm
			</TD>
		</TR>
				
			
		<TR>
			<TD>Antall meter såleblokk:</TD>
			<TD><INPUT TYPE="Text" NAME="Ms" SIZE="10">&nbsp;m</TD>
		</TR>
	</TABLE>
	<P><INPUT TYPE="Submit" NAME="BeregnMaterialliste" VALUE="Beregn Materialliste"></P>
    <p>Felter merket med (*) må fylles ut</p>
</FORM>

<p>&nbsp;</p>

<SCRIPT LANGUAGE="JavaScript">
	OppdaterKommuneListe(Prosjekt.Tf.value)
</SCRIPT>

</BODY>

</HTML>
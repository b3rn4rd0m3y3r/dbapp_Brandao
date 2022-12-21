<style>
	B.big {font-size: 10px;}
	DIV#conn {font-family: Helvetica;font-size: 7px;color:lightgrey;}
	TD, TH {font-family: Helvetica;font-size: 9px;}
	TH {background: gray;color: white;}
	TR.zebra {background: antiquewhite;}
</style>
<%
'Provedor = "Provider=Microsoft.ACE.OLEDB.12.0;"
Provedor = "Provider=Microsoft.Jet.OLEDB.4.0;"
DBQ = "Dbq=C:\Ap\dbase\MassaDeDados.xlsx;"
Driver = "Driver={Microsoft Excel Driver (*.xls, *.xlsx, *.xlsm, *.xlsb)};"
DataSource = "Data Source=C:\Ap\dbase\MassaDeDados.xls;"
CONN_STRING = Provedor & DataSource & "Extended Properties=""Excel 8.0;HDR=YES;IMEX=1"""
CONN_STRING = Driver & DBQ
Response.write "<div id=""Conn""><br><span>" & CONN_STRING & "</span><br>"
' Objetos
ADOCON = "ADODB.Connection"
ADOREC = "ADODB.Recordset"
' Conex�o
Set conn = Server.CreateObject(ADOCON)
Set rs = Server.CreateObject(ADOREC)
conn.Open CONN_STRING
strSQL = "SELECT [UN], [GR], [CodLoc], [NomeLoc], [Matricula], [Usuario], [NomeCli], "
strSQL = strSQL & "[CodLograd], [TipoLograd], [NomeLograd], [NoImovel], [TipoComp], [NomeBairro]"
strSQL = strSQL & " FROM [tab1$]"
strSQL = strSQL & " ORDER BY [UN],[GR],[NomeLoc],[NomeCli]"
Response.write "<br><span>" & strSQL & "</span><br>"
Response.write "</div>"
rs.Open strSQL, conn, 3, 3
// Dicionarios
Set Unidades = Server.CreateObject("Scripting.Dictionary")
Set Gerencias = Server.CreateObject("Scripting.Dictionary")
Set Municipios = Server.CreateObject("Scripting.Dictionary")
'* COMP�E A TABELA DE REGISTROS
CID_ant = "@#$%"
UNID_ant = "@#$%"
UNID_soma = 0
UNID = ""
GER_ant = "@#$%"
GER_soma = 0
GER = ""
Chave_CID = 1
Response.write "<table>"
Response.write "<tr><th>UN</th><th>GR</th><th>CodLoc</th><th>Nome da Localidade</th><th>Matricula</th><th>Usu�rio</th>"
Response.write "<th>Cliente</th><th>Cod.Lograd.</th><th>Tipo Lograd.</th><th>Nome Lograd.</th><th>N�mero</th>"
Response.write "<th>Tipo<br>Complem.</th><th>Bairro</th></tr>"
'Response.end
Do While Not rs.EOF
	CID = rs.Fields("NomeLoc").Value
	CLOC = rs.Fields("CodLoc").Value
	GER = rs.Fields("GR").Value
	UNID = rs.Fields("UN").Value
	'* Detecta quebra de munic�pio
	if CID <> CID_ant then
		CID_imp = CID
		CLOC_imp = CLOC
		if Chave_CID = 1 then
			Chave_CID = 0
		else
			Chave_CID = 1
		end if
		CID_ant = CID
	else
		CID_imp = ""
		CLOC_imp = ""
	end if

	'* Detecta quebra de Ger�ncia
	if GER <> GER_ant then
		'* Detecta quebra de Unidade
		if UNID <> UNID_ant and UNID_ant <> "@#$%" then
			'* Quebrou Ger�ncia e Unidade
			Unidades.Add UNID_ant, UNID_soma
			Gerencias.Add UNID_ant & GER_ant, GER_soma
			UNID_ant = UNID
			UNID_soma = 0
			GER_ant = GER
			GER_soma = 0
		else
			'* Quebrou Ger�ncia apenas
			UNID_soma = UNID_soma + 1
			if  GER_ant <> "@#$%" then
				Gerencias.Add UNID_ant & GER_ant, GER_soma
			end if
			GER_ant = GER
			GER_soma = 0
		end if
	else
		GER_soma = GER_soma + 1
	end if	
		'* Detecta quebra de Unidade (Recheck)
		if UNID <> UNID_ant then
			Unidades.Add UNID_ant, UNID_soma
			UNID_ant = UNID
			UNID_soma = 0
		else
			UNID_soma = UNID_soma + 1
		end if	



	'* Zebra por Munic�pio
	if Chave_CID = 1 then
		Response.write "<tr>"
	else
		Response.write "<tr class=zebra>"
	end if
	Response.write "<td>" & UNID & "(" & UNID_Soma & ")</td>"
	Response.write "<td>" & GER & "(" & GER_Soma & ")</td>"
	Response.write "<td><b class=""big"">" & CLOC_imp & "</b></td>"
	Response.write "<td><b class=""big"">" & CID_imp & "</b></td>"
	Response.write "<td>" & rs.Fields("Matricula").Value & "</td>"
	Response.write "<td>" & rs.Fields("Usuario").Value & "</td>"
	Response.write "<td>" & rs.Fields("NomeCli").Value & "</td>"
	Response.write "<td>" & rs.Fields("CodLograd").Value & "</td>"
	Response.write "<td>" & rs.Fields("TipoLograd").Value & "</td>"
	Response.write "<td>" & rs.Fields("NomeLograd").Value & "</td>"
	Response.write "<td>" & rs.Fields("NoImovel").Value & "</td>"
	Response.write "<td>" & rs.Fields("TipoComp").Value & "</td>"
	Response.write "<td>" & rs.Fields("NomeBairro").Value & "</td>"
	Response.write "</tr>"
	rs.MoveNext
Loop

Response.write "</table>"
rs.Close
conn.Close
Unidades.Add UNID, UNID_soma
Response.write UNID & GER
Response.write GER_soma
Gerencias.Add UNID & GER, GER_soma
'Response.end
'*
For each chave in Unidades.Keys
    Response.Write( chave & " => " & Unidades.Item(chave) & "<br>")
Next
'*
For each chave1 in Gerencias.Keys
    Response.Write( chave1 & " => " & Gerencias.Item(chave1) & "<br>")
Next
%>
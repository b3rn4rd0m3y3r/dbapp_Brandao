<style>
	B.big {font-size: 10px;}
	DIV#conn {font-family: Helvetica;font-size: 7px;color:gray;}
	TD, TH {font-family: Helvetica;font-size: 9px;}
	TH {background: gray;color: white;}
	TR.zebra {background: antiquewhite;}
</style>
<%
'* Parâmetros da URL
if Request.QueryString("Unidade") <> "" then
	UNID_get = Request.QueryString("Unidade")
else
	UNID_get = ""
end if
if Request.QueryString("Gerencia") <> "" then
	GER_get = Request.QueryString("Gerencia")
else
	GER_get = ""
end if
if Request.QueryString("Cidade") <> "" then
	CID_get = Request.QueryString("Cidade")
else
	CID_get = ""
end if
Response.write UNID_get
'Response.end
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
' Conexão
Set conn = Server.CreateObject(ADOCON)
Set rs = Server.CreateObject(ADOREC)
conn.Open CONN_STRING
strSQL = "SELECT [UN], [GR], [CodLoc], [NomeLoc], [Matricula], [Usuario], [NomeCli], "
strSQL = strSQL & "[CodLograd], [TipoLograd], [NomeLograd], [NoImovel], [TipoComp], [NomeBairro]"
strSQL = strSQL & " FROM [tab1$] WHERE [UN]"
if UNID_get <> "" then
	strSQL = strSQL & " AND [UN] = '" & UNID_get & "'"
end if
if GER_get <> "" then
	strSQL = strSQL & " AND [GR] = '" & GER_get & "'"
end if
if CID_get <> "" then
	strSQL = strSQL & " AND [NomeLoc] = '" & CID_get & "'"
end if
strSQL = strSQL & " ORDER BY [UN],[GR],[NomeLoc],[NomeCli]"
Response.write "<br><span>" & strSQL & "</span><br>"
Response.write "</div>"
rs.Open strSQL, conn, 3, 3
// Dicionarios
SCRIPT_DICT = "Scripting.Dictionary"
Set Unidades = Server.CreateObject(SCRIPT_DICT)
Set Gerencias = Server.CreateObject(SCRIPT_DICT)
Set Municipios = Server.CreateObject(SCRIPT_DICT)
'* COMPÕE A TABELA DE REGISTROS
CID_ant = "@#$%"
CID_soma = 0
UNID_ant = "@#$%"
UNID_soma = 0
UNID = ""
GER_ant = "@#$%"
GER_soma = 0
GER = ""
Chave_CID = 1
Response.write "<table>"
Response.write "<tr><th>UN</th><th>GR</th><th>CodLoc</th><th>Nome da Localidade</th><th>Matricula</th><th>Usuário</th>"
Response.write "<th>Cliente</th><th>Cod.Lograd.</th><th>Tipo Lograd.</th><th>Nome Lograd.</th><th>Número</th>"
Response.write "<th>Tipo<br>Complem.</th><th>Bairro</th></tr>"
'Response.end
Do While Not rs.EOF
	CID = rs.Fields("NomeLoc").Value
	CLOC = rs.Fields("CodLoc").Value
	GER = rs.Fields("GR").Value
	UNID = rs.Fields("UN").Value
	'* Detecta quebra de município
	if CID <> CID_ant then
		'* QCid
		CID_imp = CID
		CLOC_imp = CLOC
		'* Sinais de quebra
		CID_ant = CID
		CID_soma = 0
		if Chave_CID = 1 then
			Chave_CID = 0
		else
			Chave_CID = 1
		end if
		'* QCid
		'* Detecta quebra de Gerência
		if GER <> GER_ant then
			'* Detecta quebra de Unidade
			if UNID <> UNID_ant and UNID_ant <> "@#$%" then
				'* Quebrou Cidade, Gerência e Unidade
				Unidades.Add UNID_ant, UNID_soma
				Gerencias.Add UNID_ant & GER_ant, GER_soma
				Municipios.Add UNID_ant & GER_ant & CID_ant & "_1" , CID_soma
				UNID_ant = UNID
				UNID_soma = 0
				GER_ant = GER
				GER_soma = 0
			else
				'* Quebrou Cidade e Gerência
				UNID_soma = UNID_soma + 1
				if  GER_ant <> "@#$%" then
					Gerencias.Add UNID_ant & GER_ant, GER_soma
					'* Municipios.Add UNID_ant & GER_ant & CID_ant & "_2", CID_soma
				end if
				GER_ant = GER
				GER_soma = 0
			end if
		else
		'* Quebrou Cidade apenas
			if CID_ant <> "@#$%" then
				Municipios.Add UNID_ant & GER_ant & CID_ant , CID_soma
			end if
			'CID_ant = CID
			'CID_soma = 0
			'GER_soma = GER_soma + 1
		end if
	else
		CID_imp = ""
		CLOC_imp = ""
	end if
	'* Detecta quebra de Gerência (Recheck)
	if GER <> GER_ant then
		'* Detecta quebra de Unidade (Recheck)
		if UNID <> UNID_ant then
			Unidades.Add UNID_ant, UNID_soma
			UNID_ant = UNID
			UNID_soma = 0
		end if
		if  GER_ant <> "@#$%" then
			Gerencias.Add UNID_ant & GER_ant, GER_soma
		end if
		GER_ant = GER
		GER_soma = 0
	end if
	'* Detecta quebra de Unidade (Recheck 2)
	if UNID <> UNID_ant then
		Unidades.Add UNID_ant, UNID_soma
		UNID_ant = UNID
		UNID_soma = 0
	end if
	'* Somas
	UNID_soma = UNID_soma + 1
	GER_soma = GER_soma + 1
	CID_soma = CID_soma + 1
	
	'* Zebra por Município
	if Chave_CID = 1 then
		Response.write "<tr>"
	else
		Response.write "<tr class=zebra>"
	end if
	Response.write "<td>" & UNID & "(" & UNID_Soma & ")</td>"
	Response.write "<td>" & GER & "(" & GER_Soma & ")</td>"
	Response.write "<td><b class=""big"">" & CLOC_imp & "</b></td>"
	Response.write "<td><b class=""big"">" & CID_imp & "(" & CID_Soma & ")</b></td>"
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
'*
For each chave2 in Municipios.Keys
    Response.Write( chave2 & " => " & Municipios.Item(chave2) & "<br>")
Next
%>
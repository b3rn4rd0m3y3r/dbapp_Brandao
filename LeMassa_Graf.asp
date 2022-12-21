<style>
	A.botao, A.botao2 {font-family: Arial; padding: 4px; background: orange; border-radius: 6px; text-decoration: none;}
	A.botao2 { position: absolute; top: 460px; left: 350px;}
	B.big {font-size: 10px;}
	DIV.barra, DIV.barra2, DIV.barra3 {font-family: Arial; font-size: 12px; padding: 6px; background-color:red;cursor:pointer;}
	DIV.barra2 {background-color:green;}
	DIV.barra3 {background-color:teal;font-size: 7px;}
	DIV.barra SPAN, DIV.barra2 SPAN, DIV.barra3 SPAN {position: absolute;display: inline-block; bottom:-30px;}
	DIV#conn {font-family: Helvetica;font-size: 7px;color:gray;}
	TABLE {background: palegoldenrod;}
	TD, TH {font-family: Helvetica;font-size: 9px;}
	TH {background: gray;color: white;}
	TR.zebra {background: antiquewhite;}
</style>
<%
'*
'* Funções que tratam a hierarquia da tag
'*
function unidade(texto)
	unidade = Mid(texto,1,4)
end function

function gerencia(texto)
	gerencia = Mid(texto,5,4)
end function

function municipio(texto)
	municipio = Mid(texto,9)
end function

'Response.write unidade("UUUUXXXX")
'*
'* Parâmetros da URL
'*
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
'Response.end
'*
'* Parâmetros para conexão ao banco
'*
'Provedor = "Provider=Microsoft.ACE.OLEDB.12.0;"
Provedor = "Provider=Microsoft.Jet.OLEDB.4.0;"
DBQ = "Dbq=C:\Ap\dbase\MassaDeDados.xlsx;"
Driver = "Driver={Microsoft Excel Driver (*.xls, *.xlsx, *.xlsm, *.xlsb)};"
DataSource = "Data Source=C:\Ap\dbase\MassaDeDados.xls;"
CONN_STRING = Provedor & DataSource & "Extended Properties=""Excel 8.0;HDR=YES;IMEX=1"""
CONN_STRING = Driver & DBQ
'Response.write "<div id=""Conn""><br><span>" & CONN_STRING & "</span><br>"
'
' Classes de Objetos ADO
'
ADOCON = "ADODB.Connection"
ADOREC = "ADODB.Recordset"
'* Conexão
Set conn = Server.CreateObject(ADOCON)
Set rs = Server.CreateObject(ADOREC)
conn.Open CONN_STRING
'* Sentença SQL de extração de dados
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
'Response.write "<br><span>" & strSQL & "</span><br>"
'Response.write "</div>"
rs.Open strSQL, conn, 3, 3
'* Teste de existência de registros para este filtro
if rs.RecordCount = 0 then
	Response.write "[No registers in the recordset.]"
	Response.end
end if
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
'*
'* Loop de contabilização
'*
BRK = "@#$%"
Do While Not rs.EOF
	CID = rs.Fields("NomeLoc").Value
	CLOC = rs.Fields("CodLoc").Value
	GER = rs.Fields("GR").Value
	UNID = rs.Fields("UN").Value
	'* Verificação de quebra da menor para a maior hierarquia
	'* Detecta quebra de município
	if CID <> CID_ant then
		'* Sinais de quebra - 19/12/2022
		if CID_ant <> BRK then
			Municipios.Add UNID_ant & GER_ant & CID_ant , CID_soma
		end if
		CID_ant = CID
		CID_soma = 0
	end if
	'* Detecta quebra de Gerência
	if GER <> GER_ant then
		if  GER_ant <> BRK then
			Gerencias.Add UNID_ant & GER_ant, GER_soma
		end if
		GER_ant = GER
		GER_soma = 0
	end if
	'* Detecta quebra de Unidade
	if UNID <> UNID_ant then 
		if UNID_ant <> BRK then
			Unidades.Add UNID_ant, UNID_soma
		end if
		UNID_ant = UNID
		UNID_soma = 0
	end if
	'* Somas
	UNID_soma = UNID_soma + 1
	GER_soma = GER_soma + 1
	CID_soma = CID_soma + 1
	rs.MoveNext
Loop
'Response.write "</table>"
rs.Close
conn.Close
'*
'* Ajuste da quebra final
'*
'Response.write UNID & GER
'Response.write GER_soma
Unidades.Add UNID, UNID_soma
Gerencias.Add UNID & GER, GER_soma
Municipios.Add UNID & GER & CID , CID_soma
'*
'* Loop nas Unidades
'*
UNID_max = 0
Response.write "<table>"
Response.write "<tr><th>Unidade<th>Quantidade</tr>"
For each chave in Unidades.Keys
    if chave <> "@#$%" then
		Response.Write( "<tr><td>" & chave & "</td><td align=right>" & Unidades.Item(chave) & "</td></tr>")
		if Unidades.Item(chave) > UNID_max then
			UNID_max = Unidades.Item(chave)
		end if
	end if
Next
Response.write "</table>"
'Response.write UNID_max
'*
'* Loop nas Gerências
'*
Response.write "<table>"
Response.write "<tr><th>Unidade<th>Gerência<th>Quantidade</tr>"
GER_max = 0
For each chave1 in Gerencias.Keys
    Response.Write( "<tr><td>" & unidade(chave1) & "</td><td>" & gerencia(chave1) & "</td><td align=right>" & Gerencias.Item(chave1) & "</td></tr>")
	if CInt(Gerencias.Item(chave1)) > GER_max then
		GER_max = CInt(Gerencias.Item(chave1))
	end if
Next
Response.write "</table>"
'Response.write "GER:" & GER_max
'*
'* Loop nos Municípios
'*
Response.write "<table>"
Response.write "<tr><th>Unidade<th>Gerência<th>Município<th>Quantidade</tr>"
CID_max = 0
For each chave2 in Municipios.Keys
    Response.Write( "<tr><td>" & unidade(chave2) & "</td><td>" & gerencia(chave2) & "</td><td>" & municipio(chave2) & "</td><td align=right>" & Municipios.Item(chave2) & "</td></tr>")
	if Municipios.Item(chave2) > CID_max then
		CID_max = Municipios.Item(chave2)
	end if
Next
Response.write "</table>"
'*
'* GRAFICOS
'*
Esquerda = 20
LARG = 38
Distancia = 180
if UNID_get <> "" OR GER_get <> "" then
	Distancia = 350
end if
FUNDO = 480
TOPO_BASE = 150
ALTURA_MAX = UNID_MAX
Razao_altura = ALTURA_MAX/(FUNDO - TOPO_BASE)
'* Gráfico das Unidades
if ( UNID_get & GER_get & CID_get ) = "" then
	For each chave in Unidades.Keys
			if chave <> "@#$%" then
				Altura = CInt(CInt(Unidades.Item(chave))/Razao_altura)
				Margem_topo = ALTURA_MAX/Razao_altura - Altura
				Estilo = "position: absolute;top: 80px;margin-top:" & Margem_topo & "px;width: " & CStr(LARG) & "px" & ";height:" & Altura & ";left:" & CStr(Distancia) & "px;"
				Response.Write( "<div class=barra style=""" & Estilo & """ onclick=""window.location.href = 'LeMassa_Graf.asp?Unidade=" & unidade(chave) & "';""><span>" & chave & "</span></div>")
				Distancia = Distancia + Esquerda + LARG
			end if
	Next
end if
'* Gráfico de gerências
if ( GER_get & CID_get ) = "" AND UNID_get <> "" then
	For each chave in Gerencias.Keys
			if chave <> "@#$%" then
				Altura = CInt(CInt(Gerencias.Item(chave))/Razao_altura)
				Margem_topo = ALTURA_MAX/Razao_altura - Altura
				Estilo = "position: absolute;top: 80px;margin-top:" & Margem_topo & "px;width: " & CStr(LARG) & "px" & ";height:" & Altura & ";left:" & CStr(Distancia) & "px;"
				Response.Write( "<div class=barra2 style=""" & Estilo & """ onclick=""window.location.href = 'LeMassa_Graf.asp?Gerencia=" & gerencia(chave) & "';""><span>" & gerencia(chave) & "</span></div>")
				Distancia = Distancia + Esquerda + LARG
			end if
	Next
	Response.write "<br><a class=botao2 href=""?"">Voltar à Tela Inicial</a>"
end if
'* Gráfico de cidades
if ( GER_get <> "" ) then
	chave_ant = ""
	For each chave in Municipios.Keys
			'Response.Write chave & "<br>"
			if chave <> "@#$%" then
				Altura = CInt(CInt(Municipios.Item(chave))/Razao_altura)
				Margem_topo = ALTURA_MAX/Razao_altura - Altura
				Estilo = "position: absolute;top: 80px;margin-top:" & Margem_topo & "px;width: " & CStr(LARG) & "px" & ";height:" & Altura & ";left:" & CStr(Distancia) & "px;"
				Response.Write( "<div class=barra3 style=""" & Estilo & """ onclick=""window.location.href = 'LeMassa.asp?Cidade=" & municipio(chave) & "';""><span style=""line-height: 110%;"">" & municipio(chave) & "</span></div>")
				Distancia = Distancia + Esquerda + LARG
			end if
			chave_ant = chave
	Next
	Response.write "<br><a class=botao href=""?Unidade=" & unidade(chave_ant) & """>Voltar à " & unidade(chave_ant) & "</a>"
end if
%>
<style>
	DIV#conn {font-family: Helvetica;font-size: 7px;color:lightgrey;}
	TD, TH {font-family: Helvetica;font-size: 9px;}
	TH {background: gray;color: white;}
</style>
<%
'Provedor = "Provider=Microsoft.ACE.OLEDB.12.0;"
Provedor = "Provider=Microsoft.Jet.OLEDB.4.0;"
DBQ = "Dbq=C:\Ap\dbase\MassaDeDados.xlsx;"
Driver = "Driver={Microsoft Excel Driver (*.xls, *.xlsx, *.xlsm, *.xlsb)};"
'Driver = "Driver={Microsoft Excel Driver (*.xls)};"
DataSource = "Data Source=C:\Ap\dbase\MassaDeDados.xls;"
'DataSource = "Data Source=" & Server.MapPath("tabela1.xls") & ";"
CONN_STRING = Provedor & DataSource & "Extended Properties=""Excel 8.0;HDR=YES;IMEX=1"""
'CONN_STRING = Provedor & DataSource
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
strSQL = strSQL & " FROM [tab1$]"
strSQL = strSQL & " ORDER BY [UN],[GR],[NomeLoc],[NomeCli]"
Response.write "<br><span>" & strSQL & "</span><br>"
Response.write "</div>"
rs.Open strSQL, conn, 3, 3
Response.write "<table>"
Response.write "<tr><th>UN</th><th>GR</th><th>CodLoc</th><th>Nome da Localidade</th><th>Matricula</th><th>Usuário</th>"
Response.write "<th>Cliente</th><th>Cod.Lograd.</th><th>Tipo Lograd.</th><th>Nome Lograd.</th><th>Número</th>"
Response.write "<th>Tipo<br>Complem.</th><th>Bairro</th></tr>"
Do While Not rs.EOF
	Response.write "<tr>"
	Response.write "<td>" & rs.Fields("UN").Value & "</td>"
	Response.write "<td>" & rs.Fields("GR").Value & "</td>"
	Response.write "<td>" & rs.Fields("CodLoc").Value & "</td>"
	Response.write "<td>" & rs.Fields("NomeLoc").Value & "</td>"
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


%>
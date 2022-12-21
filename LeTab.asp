<%
'Provedor = "Provider=Microsoft.ACE.OLEDB.12.0;"
Provedor = "Provider=Microsoft.Jet.OLEDB.4.0;"
DBQ = "Dbq=C:\Ap\dbase\tabela1.xlsx;"
Driver = "Driver={Microsoft Excel Driver (*.xls, *.xlsx, *.xlsm, *.xlsb)};"
'Driver = "Driver={Microsoft Excel Driver (*.xls)};"
DataSource = "Data Source=C:\Ap\dbase\tabela1.xls;"
'DataSource = "Data Source=" & Server.MapPath("tabela1.xls") & ";"
CONN_STRING = Provedor & DataSource & "Extended Properties=""Excel 8.0;HDR=YES;IMEX=1"""
'CONN_STRING = Provedor & DataSource
CONN_STRING = Driver & DBQ
Response.write "<br>" & CONN_STRING & "<br>"
' Objetos
ADOCON = "ADODB.Connection"
ADOREC = "ADODB.Recordset"
' Conexão
Set conn = Server.CreateObject(ADOCON)
Set rs = Server.CreateObject(ADOREC)
conn.Open CONN_STRING
rs.Open "SELECT [ID], [DESCRI] FROM [tab1$A1:B10]", conn, 3, 3
Response.write "<br>"
Do While Not rs.EOF
	Response.write rs.Fields("ID").Value & " - " & rs.Fields("DESCRI").Value & "<br>"
	rs.MoveNext
Loop
rs.Close
conn.Close


%>
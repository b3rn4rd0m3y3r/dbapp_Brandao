<%
Const adOpenStatic = 3
Const adLockOptimistic = 3
Const adCmdText = &H1

Response.write Server.MapPath("tabela1.dbf")
Provedor = "Provider=Microsoft.ACE.OLEDB.12.0;"
Provedor = "OLEDB;Provider=Microsoft.Jet.OLEDB.4.0;"
Provedor = "Provider=Microsoft.Jet.OLEDB.4.0;"
DBQ = "Dbq=C:\Inetpub\wwwroot\dbase\tabela1.xls;DriverId=790;FIL=excel 8.0;"
DBQ = "Dbq=" & Server.MapPath("tabela1.dbf") & ";"
'Driver = "Driver={Microsoft Excel Driver (*.xls, *.xlsx, *.xlsm, *.xlsb)};"
Driver = "Driver={Microsoft dBase Driver (*.dbf)};DriverID=277;"
DataSource = "Data Source=" & Server.MapPath("tabela1.dbf") & ";"
CONN_STRING = Provedor & DataSource & "Extended Properties=""Excel 8.0;HDR=YES;IMEX=1"""
CONN_STRING = Driver & DBQ & "Extended Properties=""DBASE III;"";"
CONN_STRING = Driver & DBQ 
Response.write "<br>" & CONN_STRING & "<br>"
'CONN_STRING = Provedor & DataSource & "Extended Properties=""Excel 8.0;HDR=YES;IMEX=1"""
' Objetos
ADOCON = "ADODB.Connection"
ADOREC = "ADODB.Recordset"
' Conex�o
Set conn = Server.CreateObject(ADOCON)
'Set rs = Server.CreateObject(ADOREC)
'conn.Open CONN_STRING
'rs.Open "SELECT [ID], [DESCRI] FROM [tab1$A1:B10]", conn, 3, 3
Response.write "<br>"
'Do While Not rs.EOF
'	Response.write rs.Fields("ID").Value & " - " & rs.Fields("DESCRI").Value & "<br>"
'	rs.MoveNext
'Loop
'rs.Close
conn.Close
%>
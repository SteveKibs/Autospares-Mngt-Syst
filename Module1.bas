Attribute VB_Name = "Module1"


Global con As adodb.Connection
Global rslogin As New adodb.Recordset
Global rsNewItem As New adodb.Recordset
Global rsOrder As New adodb.Recordset
Global rsPurchase As New adodb.Recordset
Global rsSales As New adodb.Recordset
Global rsStock As New adodb.Recordset



Public Sub connect()
Set con = New adodb.Connection
Set rsNewItem = New adodb.Recordset
Set rslogin = New adodb.Recordset
Set rsPurchase = New adodb.Recordset
Set rsSales = New adodb.Recordset
Set rsStock = New adodb.Recordset
con.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=Mwitidb.mdb;Persist Security Info=False"
con.Open
End Sub






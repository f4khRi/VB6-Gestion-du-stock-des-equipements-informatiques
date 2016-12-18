Attribute VB_Name = "Module1"
Public rs As ADODB.Recordset
Public cn As ADODB.Connection
Public rs2 As ADODB.Recordset
Public rsfourni As ADODB.Recordset
Public rsEQUIPEMENT As ADODB.Recordset


Public Sub connect()
Set cn = New ADODB.Connection
cn.Provider = "Microsoft.jet.oledb.4.0"
cn.ConnectionString = App.Path & "\GStock.mdb"
cn.Open
End Sub


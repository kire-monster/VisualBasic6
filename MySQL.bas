Attribute VB_Name = "MySQL"


Public Sub openMySQL(db As ADODB.Connection, connString As String)
On Error Resume Next
    db.CommandTimeout = 180
    db.CursorLocation = 1 'adUseClient '1
    db.Open "DRIVER={MySQL ODBC 3.51 Driver};DATABASE=__DB__;SERVER=__SERVER__;UID=__USER__;password=__PWD__;PORT=3306;"
End Sub


Public Sub execMySQL(db As ADODB.Connection, rs As ADODB.Recordset, query As String)
On Error Resume Next
    Dim cmd As ADODB.Command
    Set cmd = New ADODB.Command
    Set cmd.ActiveConnection = db
    cmd.CommandText = query
    rs.CursorLocation = adUseClient
    
    'rs.Open cmd, , 1, 1
    rs.Open cmd, db, adOpenStatic, adLockOptimistic, adAsyncConnect
End Sub

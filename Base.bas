Attribute VB_Name = "Base"

Sub Main()
    frmMain.Show
End Sub


'Sub Main()
'    Dim con As Connection
'    Dim rs As Recordset
'    Set con = New Connection
'    Set rs = New Recordset
'
'    'abrimos la conexion
'    openDB con, "PROVIDER=SQLOLEDB; DATA SOURCE=192.168.1.70; UID=sa; PWD=123;" ' DATABASE=__DB__
'    If Err.Description <> "" Then
'        MsgBox "Error al conectarse, Detalle: " + Err.Description
'        Exit Sub
'    End If
'
'
'    'configuramos el objeto de recorrido
'    rs.CursorLocation = adUseClient
'    rs.CursorType = adOpenStatic
'    rs.LockType = adLockBatchOptimistic
'
'    'ejecutamos la consulta
'    execSQL con, rs, "select @@version as version"
'    If Err.Description <> "" Then
'        MsgBox "Error al ejecutar SQL, Detalle: " + Err.Description
'        Exit Sub
'    End If
'
'
'    'recorremos registros
'    Do Until rs.EOF
'        MsgBox rs(0)
'        rs.MoveNext
'    Loop
'
'
'    ' cerramos sesion
'    closeDB con
'
'End Sub

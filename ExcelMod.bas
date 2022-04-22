Attribute VB_Name = "ExcelMod"
Public Sub LeerExcel(filename As String)

    Dim xlApp As Excel.Application
    Dim xlLibro As Excel.Workbook
    Dim xlHoja As Excel.Worksheet
    Dim varMatriz As Variant

    Set xlApp = New Excel.Application
    xlApp.Visible = True
'
'
    Set xlLibro = xlApp.Workbooks.Open(App.Path & "\" & filename, True, True, , "")
    Set xlHoja = xlApp.Worksheets("Hoja1") 'Nombre de la hoja del excel que desea leer

    'El rango a leer
    'varMatriz = xlHoja.Range("A1:C10").Value
    
    
    
    
    
    'leer datos
    For x = 1 To 110 Step 1
        'FileWrite "datos.txt", xlHoja.Cells(x, 1)
    Next x


    'cerramos el archivo Excel
    xlLibro.Close SaveChanges:=False
    xlApp.Quit

    'reset variables de los objetos
    Set xlHoja = Nothing
    Set xlLibro = Nothing
    Set xlApp = Nothing




End Sub



Private Sub CrearExcel(rs As Recordset)
On Error Resume Next


    Dim filename As String
    Dim name As String
    filename = "reporte.xlsx"
    name = Dir$(App.Path & "\" & filename)
    
    If Len(name) > 0 Then
        Kill App.Path & "\" & filename
    End If
    
    
    
    Dim xlApp As Excel.Application
    Dim xlLibro As Excel.Workbook
    Dim xlHoja As Excel.Worksheet

    Set xlApp = New Excel.Application
    xlApp.Visible = True

    Set xlLibro = xlApp.Workbooks.Add 'xlApp.Workbooks.Open(App.Path & "\resultado.xlsx", True, True, , "")
    Set xlHoja = xlApp.Worksheets(1)


    Dim x As Integer
    x = 1
    Do Until rs.EOF
        xlHoja.Cells(x, 1).Value = rs(0)
        xlHoja.Cells(x, 2).Value = rs(1)
        xlHoja.Cells(x, 3).Value = rs(2)
        xlHoja.Cells(x, 4).Value = rs(3)
        xlHoja.Cells(x, 5).Value = rs(4)
        xlHoja.Cells(x, 6).Value = rs(5)
        xlHoja.Cells(x, 7).Value = rs(6)
        xlHoja.Cells(x, 8).Value = rs(7)
        xlHoja.Cells(x, 9).Value = rs(8)
        xlHoja.Cells(x, 10).Value = rs(9)
        xlHoja.Cells(x, 11).Value = rs(10)
        xlHoja.Cells(x, 12).Value = rs(11)
        xlHoja.Cells(x, 13).Value = rs(12)
        x = x + 1
        'DoEvents
        rs.MoveNext
        
    Loop


    xlLibro.SaveAs App.Path & "\" & filename
    xlApp.Quit

End Sub



Attribute VB_Name = "File"
Public Sub FileWrite(filename As String, cadena As String)
    Open App.path & "\" & filename For Append As #1
        Write #1, cadena
    Close #1
End Sub

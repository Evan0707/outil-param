' ===================================================
' MODULE: mod_Config
' Lecture centralisée des paramètres depuis la feuille Config
' ===================================================

Function GetConfig(parametre As String) As String
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets("Config")
    On Error GoTo 0
    
    If ws Is Nothing Then
        MsgBox "Feuille 'Config' introuvable !", vbCritical
        GetConfig = ""
        Exit Function
    End If
    
    Dim i As Long, lastRow As Long
    lastRow = ws.Cells(ws.Rows.count, "A").End(xlUp).row
    
    For i = 2 To lastRow
        If UCase(Trim(ws.Cells(i, 1).Value)) = UCase(Trim(parametre)) Then
            GetConfig = CStr(ws.Cells(i, 2).Value)
            Exit Function
        End If
    Next i
    
    GetConfig = ""
End Function

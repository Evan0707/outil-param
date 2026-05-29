Attribute VB_Name = "mod_Logger"
' ===================================================
' MODULE: mod_Logger
' Système centralisé de gestion des erreurs et logs
' ===================================================

Public Const LOG_INFO As String = "INFO"
Public Const LOG_WARNING As String = "WARNING"
Public Const LOG_ERROR As String = "ERROR"

Sub LogInfo(source As String, message As String)
    Call WriteLog(LOG_INFO, source, message, 0)
End Sub

Sub LogWarning(source As String, message As String)
    Call WriteLog(LOG_WARNING, source, message, 0)
End Sub

Sub LogError(source As String, message As String, Optional errNumber As Long = 0)
    Call WriteLog(LOG_ERROR, source, message, errNumber)
End Sub

Private Sub WriteLog(typeLog As String, source As String, message As String, errNumber As Long)
    ' Vérifier si log actif
    If UCase(GetConfig("LogActif")) = "NON" Then Exit Sub
    
    ' Filtrer par niveau
    Dim niveauMin As String
    niveauMin = UCase(GetConfig("LogNiveauMin"))
    If niveauMin = "" Then niveauMin = LOG_INFO
    
    If niveauMin = LOG_WARNING And typeLog = LOG_INFO Then Exit Sub
    If niveauMin = LOG_ERROR And (typeLog = LOG_INFO Or typeLog = LOG_WARNING) Then Exit Sub
    
    ' Récupérer/créer feuille Logs
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets("Logs")
    On Error GoTo 0
    If ws Is Nothing Then Call CreerFeuilleLogs
    Set ws = ThisWorkbook.Sheets("Logs")
    
    ' Trouver ligne suivante
    Dim row As Long
    row = ws.Cells(ws.Rows.count, "A").End(xlUp).row + 1
    If row < 2 Then row = 2
    
    ' Écrire le log
    ws.Cells(row, 1).Value = Format(Now, "dd/mm/yyyy hh:mm:ss")
    ws.Cells(row, 2).Value = typeLog
    ws.Cells(row, 3).Value = source
    ws.Cells(row, 4).Value = message
    If errNumber <> 0 Then ws.Cells(row, 5).Value = errNumber
    
    ' Coloration
    Select Case typeLog
        Case LOG_INFO
            ws.Cells(row, 2).Font.Color = RGB(0, 0, 0)
        Case LOG_WARNING
            ws.Cells(row, 2).Font.Color = RGB(255, 140, 0)
            ws.Cells(row, 2).Font.Bold = True
        Case LOG_ERROR
            ws.Cells(row, 2).Font.Color = RGB(192, 0, 0)
            ws.Cells(row, 2).Font.Bold = True
            ws.Rows(row).Interior.Color = RGB(255, 235, 235)
    End Select
    
    ' Limite de lignes
    Call NettoyerLogsAnciens(ws)
End Sub

Private Sub CreerFeuilleLogs()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets.Add
    ws.Name = "Logs"
    
    ws.Cells(1, 1).Value = "Date/Heure"
    ws.Cells(1, 2).Value = "Type"
    ws.Cells(1, 3).Value = "Source"
    ws.Cells(1, 4).Value = "Message"
    ws.Cells(1, 5).Value = "N° Erreur"
    
    ws.Rows(1).Font.Bold = True
    ws.Rows(1).Interior.Color = RGB(220, 220, 220)
    
    ws.Columns("A").ColumnWidth = 18
    ws.Columns("B").ColumnWidth = 10
    ws.Columns("C").ColumnWidth = 25
    ws.Columns("D").ColumnWidth = 60
    ws.Columns("E").ColumnWidth = 10
End Sub

Private Sub NettoyerLogsAnciens(ws As Worksheet)
    Dim maxLignes As Long
    maxLignes = CLng(Val(GetConfig("LogMaxLignes")))
    If maxLignes <= 0 Then maxLignes = 1000
    
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.count, "A").End(xlUp).row
    
    If lastRow - 1 > maxLignes Then
        Dim aSupprimer As Long
        aSupprimer = (lastRow - 1) - maxLignes
        ws.Rows("2:" & (1 + aSupprimer)).Delete
    End If
End Sub

Sub AfficherLogs()
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets("Logs")
    On Error GoTo 0
    
    If ws Is Nothing Then
        MsgBox "Aucun log enregistré pour le moment.", vbInformation
        Exit Sub
    End If
    
    ws.Activate
End Sub

Sub ViderLogs()
    Dim reponse As VbMsgBoxResult
    reponse = MsgBox("Voulez-vous vraiment vider tous les logs ?", vbYesNo + vbQuestion, "Confirmation")
    If reponse <> vbYes Then Exit Sub
    
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets("Logs")
    On Error GoTo 0
    
    If ws Is Nothing Then Exit Sub
    
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.count, "A").End(xlUp).row
    If lastRow > 1 Then ws.Rows("2:" & lastRow).Delete
    
    MsgBox "Logs vidés.", vbInformation
End Sub

' Cleanup global appelé en cas d'erreur fatale
Sub RestaurerEtatExcel()
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    
    ' Fermer fichiers externes restés ouverts
    On Error Resume Next
    Dim wbTemp As Workbook
    For Each wbTemp In Workbooks
        If wbTemp.Name <> ThisWorkbook.Name Then wbTemp.Close False
    Next wbTemp
    On Error GoTo 0
End Sub

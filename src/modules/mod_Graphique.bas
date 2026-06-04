
' ===================================================
' MODULE: mod_Graphique
' Génération du graphique PDM
' ===================================================

Sub VoirGraphiqueDepuisBouton()
    Dim btn As Button
    Set btn = ActiveSheet.Buttons(Application.Caller)
    Dim ligneSelectionnee As Long
    ligneSelectionnee = btn.TopLeftCell.row
    Call VoirGraphiqueAvecLigne(ligneSelectionnee)
End Sub

Sub VoirGraphiqueAvecLigne(ligneSelectionnee As Long)
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Résultats")
    On Error GoTo ErrorHandler
    
    Dim nomParam As String
    Dim codeEqt As String
    nomParam = ws.Cells(ligneSelectionnee, 1).Value
    codeEqt = Replace(ws.Cells(1, 1).Value, "Équipement : ", "")
    
    If nomParam = "" Then
        MsgBox "Aucun paramètre sur cette ligne"
        Exit Sub
    End If
    
    Dim cheminReleve As String
    cheminReleve = GetCheminReleve()
    
    If Dir(cheminReleve) = "" Then
        MsgBox "Fichier relevés introuvable : " & cheminReleve
        Exit Sub
    End If
    
    Application.ScreenUpdating = False
    Dim wbReleve As Workbook, wsReleve As Worksheet
    Set wbReleve = Workbooks.Open(cheminReleve, ReadOnly:=True)
    Set wsReleve = wbReleve.Worksheets(1)
    Dim lastRow As Long
    lastRow = wsReleve.Cells(wsReleve.Rows.count, "T").End(xlUp).row
    
    Dim dates() As String
    Dim valeurs() As Double
    Dim seuilB As Double, seuilH As Double, seuilTB As Double, seuilTH As Double
    Dim count As Integer: count = 0
    ReDim dates(1 To lastRow)
    ReDim valeurs(1 To lastRow)
    
    Dim i As Long
    Dim codeSansTranche As String
    codeSansTranche = SansTrancheCode(codeEqt)
    
    For i = 2 To lastRow
        Dim codeReleve As String
        codeReleve = Trim(wsReleve.Cells(i, "T").Value)
        Dim paramReleve As String
        paramReleve = Trim(wsReleve.Cells(i, "O").Value)
        
        If UCase(SansTrancheCode(codeReleve)) = UCase(codeSansTranche) And _
           UCase(paramReleve) = UCase(nomParam) Then
            count = count + 1
            dates(count) = wsReleve.Cells(i, "P").Value
            If IsNumeric(wsReleve.Cells(i, "A").Value) Then
                valeurs(count) = CDbl(wsReleve.Cells(i, "A").Value)
            End If
            If count = 1 Then
                If IsNumeric(wsReleve.Cells(i, "J").Value) Then seuilB = CDbl(wsReleve.Cells(i, "J").Value)
                If IsNumeric(wsReleve.Cells(i, "K").Value) Then seuilH = CDbl(wsReleve.Cells(i, "K").Value)
                If IsNumeric(wsReleve.Cells(i, "L").Value) Then seuilTH = CDbl(wsReleve.Cells(i, "L").Value)
                If IsNumeric(wsReleve.Cells(i, "M").Value) Then seuilTB = CDbl(wsReleve.Cells(i, "M").Value)
            End If
        End If
    Next i
    
    wbReleve.Close False
    Application.ScreenUpdating = True
    
    If count = 0 Then
        MsgBox "Aucun relevé trouvé pour : " & nomParam
        Exit Sub
    End If
    
    Dim wsGraph As Worksheet
    On Error Resume Next
    Set wsGraph = ThisWorkbook.Sheets("Graphique PDM")
    On Error GoTo 0
    If wsGraph Is Nothing Then
        Set wsGraph = ThisWorkbook.Sheets.Add
        wsGraph.Name = "Graphique PDM"
    End If
    wsGraph.Cells.Clear
    wsGraph.ChartObjects.Delete
    
    wsGraph.Cells(1, 1).Value = "Date"
    wsGraph.Cells(1, 2).Value = "Valeur"
    wsGraph.Cells(1, 3).Value = "Seuil B"
    wsGraph.Cells(1, 4).Value = "Seuil H"
    wsGraph.Cells(1, 5).Value = "Seuil TB"
    wsGraph.Cells(1, 6).Value = "Seuil TH"
    wsGraph.Rows(1).Font.Bold = True
    
    For i = 1 To count
        wsGraph.Cells(i + 1, 1).Value = dates(i)
        wsGraph.Cells(i + 1, 2).Value = valeurs(i)
        wsGraph.Cells(i + 1, 3).Value = seuilB
        wsGraph.Cells(i + 1, 4).Value = seuilH
        wsGraph.Cells(i + 1, 5).Value = seuilTB
        wsGraph.Cells(i + 1, 6).Value = seuilTH
    Next i
    
    wsGraph.Columns("A").NumberFormat = "dd/mm/yyyy"
    
    Dim chartObj As ChartObject
    Set chartObj = wsGraph.ChartObjects.Add(Left:=10, top:=count * 18 + 60, Width:=700, Height:=400)
    
    Dim ch As Chart
    Set ch = chartObj.Chart
    ch.ChartType = xlLine
    ch.SetSourceData source:=wsGraph.Range("A1:F" & count + 1)
    
    ch.SeriesCollection(1).Name = "Valeur mesurée"
    ch.SeriesCollection(1).Format.Line.ForeColor.RGB = RGB(0, 112, 192)
    ch.SeriesCollection(1).Format.Line.Weight = 2
    
    Dim seuilNoms(1 To 4) As String
    Dim seuilColors(1 To 4) As Long
    seuilNoms(1) = "Seuil B": seuilColors(1) = RGB(0, 176, 80)
    seuilNoms(2) = "Seuil H": seuilColors(2) = RGB(255, 165, 0)
    seuilNoms(3) = "Seuil TB": seuilColors(3) = RGB(192, 0, 0)
    seuilNoms(4) = "Seuil TH": seuilColors(4) = RGB(255, 0, 0)
    
    Dim s As Integer
    For s = 1 To 4
        ch.SeriesCollection(s + 1).Name = seuilNoms(s)
        ch.SeriesCollection(s + 1).Format.Line.ForeColor.RGB = seuilColors(s)
        ch.SeriesCollection(s + 1).Format.Line.DashStyle = msoLineDash
        ch.SeriesCollection(s + 1).Format.Line.Weight = 1.5
    Next s
    
    ch.HasTitle = True
    ch.ChartTitle.Text = codeEqt & " - " & nomParam
    
    ch.SeriesCollection(1).Trendlines.Add(Type:=xlLinear).Name = "Tendance"
    ch.SeriesCollection(1).Trendlines(1).Format.Line.ForeColor.RGB = RGB(128, 128, 128)
    ch.SeriesCollection(1).Trendlines(1).Format.Line.DashStyle = msoLineDashDot
    
    wsGraph.Columns.AutoFit
    ThisWorkbook.Sheets("Graphique PDM").Activate
    MsgBox "Graphique généré pour : " & nomParam & " (" & count & " relevés)"
    
    Exit Sub
ErrorHandler:
    Call LogError("VoirGraphiqueAvecLigne", Err.Description, Err.Number)
    Call RestaurerEtatExcel
    MsgBox "Erreur graphique : " & Err.Description, vbExclamation
End Sub


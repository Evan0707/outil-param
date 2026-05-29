Attribute VB_Name = "mod_Affichage"
Sub AfficherResultats(codeEqt As String, resultats As Variant)
    On Error GoTo ErrorHandler
    
    Application.ScreenUpdating = False
    
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets("Résultats")
    On Error GoTo ErrorHandler
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Sheets.Add
        ws.Name = "Résultats"
    End If
    ws.Cells.Clear
    
    Dim btn As Button
    For Each btn In ws.Buttons
        btn.Delete
    Next btn
    
    ws.Cells(1, 1).Value = "Équipement : " & codeEqt
    ws.Cells(1, 1).Font.Bold = True
    ws.Cells(1, 1).Font.Size = 12
    
    ws.Cells(2, 1).Value = "SAL"
    ws.Cells(2, 1).Font.Bold = True
    ws.Cells(2, 1).Font.Color = RGB(0, 112, 192)
    
    Call AfficherEnTetes(ws, 3, RGB(220, 235, 250), False)
    
    Dim chemin13E As String, cheminCAT As String
    chemin13E = GetCheminFichier13E(codeEqt)
    cheminCAT = GetCheminFichierCAT(codeEqt)
    
    Dim row As Integer: row = 4
    Dim i As Integer
    Dim nomsSAL() As String
    
    ' Vérifie si SAL a des params
    Dim hasSAL As Boolean
    hasSAL = Not IsEmpty(resultats)
    If hasSAL Then
        hasSAL = Not (UBound(resultats, 2) = 1 And resultats(1, 1) = "")
    End If
    
    If hasSAL Then
        ReDim nomsSAL(1 To UBound(resultats, 2))
        For i = 1 To UBound(resultats, 2)
            ws.Cells(row, 1).Value = resultats(1, i)
            ws.Cells(row, 2).Value = resultats(2, i)
            ws.Cells(row, 3).Value = resultats(3, i)
            ws.Cells(row, 4).Value = resultats(4, i)
            ws.Cells(row, 5).Value = resultats(5, i)
            ws.Cells(row, 6).Value = resultats(6, i)
            ws.Cells(row, 7).Value = resultats(7, i)
            nomsSAL(i) = UCase(Trim(resultats(1, i)))
            row = row + 1
        Next i
    Else
        ws.Cells(row, 1).Value = "Aucun paramètre SAL trouvé"
        ws.Cells(row, 1).Font.Italic = True
        ws.Cells(row, 1).Font.Color = RGB(150, 150, 150)
        ReDim nomsSAL(1 To 1)
        nomsSAL(1) = ""
        row = row + 1
    End If
    
    Call FermerFichiersExternes
    
    ' TABLEAU 13E
    Dim resultats13E As Variant
    resultats13E = ChercherTousParametres(codeEqt, chemin13E)
    Call AfficherTableauComplementaire(ws, resultats13E, nomsSAL, resultats, row, "13E", RGB(0, 176, 80), RGB(220, 255, 220))
    
    Call FermerFichiersExternes
    
    ' TABLEAU CAT
    Dim resultatsCAT As Variant
    resultatsCAT = ChercherTousParametres(codeEqt, cheminCAT)
    Call AfficherTableauComplementaire(ws, resultatsCAT, nomsSAL, resultats, row, "CAT", RGB(255, 102, 0), RGB(255, 235, 220))
    
    ' TABLEAU SYGMA
    Call FermerFichiersExternes
    Dim resultatsSYGMA As Variant
    resultatsSYGMA = ChercherParametresSYGMA(codeEqt, GetCheminFichierSYGMA())
    Call AfficherTableauSYGMA(ws, resultatsSYGMA, row)
    
    ws.Columns.AutoFit
    ThisWorkbook.Sheets("Résultats").Activate
    
    Application.ScreenUpdating = True
    Exit Sub

ErrorHandler:
    Call LogError("AfficherResultats", Err.Description, Err.Number)
    Call RestaurerEtatExcel
    MsgBox "Erreur affichage. Consultez les Logs.", vbExclamation
End Sub

Sub AfficherEnTetes(ws As Worksheet, row As Integer, couleurFond As Long, avecGraphique As Boolean)
    ws.Cells(row, 1).Value = "Paramètre"
    ws.Cells(row, 2).Value = "Seuil B"
    ws.Cells(row, 3).Value = "Seuil H"
    ws.Cells(row, 4).Value = "Seuil TB"
    ws.Cells(row, 5).Value = "Seuil TH"
    ws.Cells(row, 6).Value = "Val Normale"
    ws.Cells(row, 7).Value = "Unité"
    If avecGraphique Then ws.Cells(row, 7).Value = "Graphique"
    ws.Rows(row).Font.Bold = True
    ws.Rows(row).Interior.Color = couleurFond
End Sub

Sub AfficherTableauComplementaire(ws As Worksheet, resultats As Variant, nomsSAL() As String, resultatsOriginaux As Variant, ByRef row As Integer, titre As String, couleurTitre As Long, couleurFond As Long)
    On Error GoTo ErrorHandler
    If IsEmpty(resultats) Then Exit Sub
    
    Dim extra() As String
    Dim couleurs() As Long
    Dim countExtra As Integer: countExtra = 0
    ReDim extra(1 To 7, 1 To UBound(resultats, 2))
    ReDim couleurs(1 To UBound(resultats, 2))
    
    Dim i As Integer, j As Integer
    For i = 1 To UBound(resultats, 2)
        Dim nomParam As String
        nomParam = UCase(Trim(resultats(1, i)))
        
        If Not EstDansListe(nomParam, nomsSAL) Then
            countExtra = countExtra + 1
            For j = 1 To 7
                extra(j, countExtra) = resultats(j, i)
            Next j
            couleurs(countExtra) = 0
        Else
            Dim idxSAL As Integer
            idxSAL = TrouverIndexParam(nomParam, resultatsOriginaux)
            
            If idxSAL > 0 Then
                Dim different As Boolean: different = False
                If CStr(resultats(2, i)) <> CStr(resultatsOriginaux(2, idxSAL)) Then different = True
                If CStr(resultats(3, i)) <> CStr(resultatsOriginaux(3, idxSAL)) Then different = True
                If CStr(resultats(4, i)) <> CStr(resultatsOriginaux(4, idxSAL)) Then different = True
                If CStr(resultats(5, i)) <> CStr(resultatsOriginaux(5, idxSAL)) Then different = True
                
                If different Then
                    countExtra = countExtra + 1
                    For j = 1 To 7
                        extra(j, countExtra) = resultats(j, i)
                    Next j
                    couleurs(countExtra) = RGB(240, 77, 0)
                End If
            End If
        End If
    Next i
    
    If countExtra = 0 Then Exit Sub
    
    ReDim Preserve extra(1 To 7, 1 To countExtra)
    ReDim Preserve couleurs(1 To countExtra)
    
    row = row + 2
    ws.Cells(row, 1).Value = "Paramètres supplémentaires / différents " & titre
    ws.Cells(row, 1).Font.Bold = True
    ws.Cells(row, 1).Font.Color = couleurTitre
    row = row + 1
    
    Call AfficherEnTetes(ws, row, couleurFond, False)
    row = row + 1
    
    For i = 1 To countExtra
        ws.Cells(row, 1).Value = extra(1, i)
        ws.Cells(row, 2).Value = extra(2, i)
        ws.Cells(row, 3).Value = extra(3, i)
        ws.Cells(row, 4).Value = extra(4, i)
        ws.Cells(row, 5).Value = extra(5, i)
        ws.Cells(row, 6).Value = extra(6, i)
        ws.Cells(row, 7).Value = extra(7, i)
        
        If couleurs(i) <> 0 Then
            ' Trouve index SAL pour comparer cellule par cellule
            Dim idx As Integer
            idx = TrouverIndexParam(UCase(Trim(extra(1, i))), resultatsOriginaux)
            If idx > 0 Then
                If CStr(extra(2, i)) <> CStr(resultatsOriginaux(2, idx)) Then ws.Cells(row, 2).Font.Color = RGB(240, 77, 0)
                If CStr(extra(3, i)) <> CStr(resultatsOriginaux(3, idx)) Then ws.Cells(row, 3).Font.Color = RGB(240, 77, 0)
                If CStr(extra(4, i)) <> CStr(resultatsOriginaux(4, idx)) Then ws.Cells(row, 4).Font.Color = RGB(240, 77, 0)
                If CStr(extra(5, i)) <> CStr(resultatsOriginaux(5, idx)) Then ws.Cells(row, 5).Font.Color = RGB(240, 77, 0)
            End If
        End If
        
        row = row + 1
    Next i
    Exit Sub
ErrorHandler:
    Call LogError("AfficherTableauComplementaire", Err.Description, Err.Number)
End Sub

' Fonction helper pour trouver l'index d'un param dans resultatsOriginaux
Function TrouverIndexParam(nomParam As String, resultats As Variant) As Integer
    Dim i As Integer
    For i = 1 To UBound(resultats, 2)
        If UCase(Trim(CStr(resultats(1, i)))) = UCase(Trim(nomParam)) Then
            TrouverIndexParam = i
            Exit Function
        End If
    Next i
    TrouverIndexParam = 0
End Function


Sub AfficherTableauSYGMA(ws As Worksheet, resultats As Variant, ByRef row As Integer)
    On Error GoTo ErrorHandler
    If IsEmpty(resultats) Then Exit Sub
    If UBound(resultats, 2) = 0 Then Exit Sub
    
    row = row + 2
    ws.Cells(row, 1).Value = "Paramètres SYGMA"
    ws.Cells(row, 1).Font.Bold = True
    ws.Cells(row, 1).Font.Color = RGB(102, 0, 204)  ' Violet
    row = row + 1
    
    ' En-têtes
    ws.Cells(row, 1).Value = "Paramètre"
    ws.Cells(row, 2).Value = "Val Basse"
    ws.Cells(row, 3).Value = "Val Normale"
    ws.Cells(row, 4).Value = "Val Haute"
    ws.Cells(row, 5).Value = "Unité"
    ws.Rows(row).Font.Bold = True
    ws.Rows(row).Interior.Color = RGB(235, 220, 255)  ' Violet clair
    row = row + 1
    
    Dim i As Integer
    For i = 1 To UBound(resultats, 2)
        ws.Cells(row, 1).Value = resultats(1, i)
        ws.Cells(row, 2).Value = resultats(2, i)
        ws.Cells(row, 3).Value = resultats(3, i)
        ws.Cells(row, 4).Value = resultats(4, i)
        ws.Cells(row, 5).Value = resultats(5, i)
        row = row + 1
    Next i
    Exit Sub

ErrorHandler:
    Call LogError("AfficherTableauSYGMA", Err.Description, Err.Number)
End Sub


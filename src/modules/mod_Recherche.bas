' ===================================================
' MODULE: mod_Recherche
' ===================================================

Sub OuvrirFormulaire()
    On Error GoTo ErrorHandler
    RechercherEquipement.Show
    Exit Sub
ErrorHandler:
    Call LogError("OuvrirFormulaire", Err.Description, Err.Number)
    Call RestaurerEtatExcel
    MsgBox "Erreur ouverture formulaire. Consultez les Logs.", vbExclamation
End Sub

Sub RechercherEquipement()
    On Error GoTo ErrorHandler

    Dim popup As New FrmRecherche
    popup.Show
    If Not popup.Confirmed Then Exit Sub

    Dim codeEqt As String
    codeEqt = Trim(popup.codeEqt)
    If codeEqt = "" Then Exit Sub

    Call LogInfo("RechercherEquipement", "Recherche : " & codeEqt)

    If Not EquipementExiste(codeEqt, GetCheminFichier(codeEqt)) Then
        Call LogWarning("RechercherEquipement", "Non trouv  : " & codeEqt)
        MsgBox Chr(201) & "quipement non trouv" & Chr(233) & " : " & codeEqt, vbInformation
        Exit Sub
    End If

    Dim frm As New resultats
    frm.SetSources popup.loadSAL, popup.load13E, popup.loadCAT, popup.loadSYGMA
    frm.InitAvecCode codeEqt
    frm.Show

    Exit Sub
ErrorHandler:
    Call LogError("RechercherEquipement", Err.Description, Err.Number)
    Call RestaurerEtatExcel
    MsgBox "Erreur. Consultez les Logs.", vbExclamation
End Sub

Function EquipementExiste(codeEqt As String, fichier As String) As Boolean
    On Error GoTo ErrorHandler
    
    If fichier = "" Then
        Call LogWarning("EquipementExiste", "Chemin fichier vide pour : " & codeEqt)
        Exit Function
    End If
    
    If Dir(fichier) = "" Then
        Call LogWarning("EquipementExiste", "Fichier introuvable : " & fichier)
        MsgBox "Fichier introuvable : " & fichier, vbExclamation
        Exit Function
    End If
    
    Dim wb As Workbook
    Application.ScreenUpdating = False
    Set wb = Workbooks.Open(fichier, ReadOnly:=True, UpdateLinks:=0)
    wb.Windows(1).Visible = False
    
    Dim ws As Worksheet
    Set ws = wb.Worksheets(1)
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.count, "F").End(xlUp).row
    
    Dim i As Long
    Dim codeUC As String: codeUC = UCase(Trim(codeEqt))
    For i = 2 To lastRow
        If UCase(Trim(ws.Cells(i, "F").Value)) = codeUC Then
            EquipementExiste = True
            wb.Close False
            Exit Function
        End If
    Next i
    
    wb.Close False
    Exit Function

ErrorHandler:
    Call LogError("EquipementExiste", Err.Description, Err.Number)
    On Error Resume Next
    If Not wb Is Nothing Then wb.Close False
    EquipementExiste = False
End Function

Function ChercherParametres(codeEqt As String, fichier As String) As Variant
    On Error GoTo ErrorHandler
    
    If fichier = "" Or Dir(fichier) = "" Then
        Call LogWarning("ChercherParametres", "Fichier param tres introuvable : " & fichier)
        Exit Function
    End If
    
    Dim wb As Workbook
    Application.ScreenUpdating = False
    Set wb = Workbooks.Open(fichier, ReadOnly:=True, UpdateLinks:=0)
    wb.Windows(1).Visible = False
    Dim ws As Worksheet
    Set ws = wb.Worksheets(1)
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.count, "B").End(xlUp).row
    
    Dim resultats() As String
    Dim count As Integer: count = 0
    ReDim resultats(1 To 13, 1 To lastRow)

    Dim codeUC As String: codeUC = UCase(Trim(codeEqt))
    Dim i As Long
    For i = 2 To lastRow
        If UCase(Trim(ws.Cells(i, "B").Value)) = codeUC Then
            count = count + 1
            resultats(1, count) = ws.Cells(i, "G").Value
            resultats(2, count) = ws.Cells(i, "I").Value
            resultats(3, count) = ws.Cells(i, "L").Value
            resultats(4, count) = ws.Cells(i, "M").Value
            resultats(5, count) = ws.Cells(i, "P").Value
            resultats(6, count) = ws.Cells(i, "N").Value
            resultats(7, count) = ws.Cells(i, "O").Value
            resultats(8, count) = ws.Cells(i, "Q").Value
            resultats(9, count) = ws.Cells(i, "J").Value
            resultats(10, count) = ws.Cells(i, "K").Value
            resultats(11, count) = ws.Cells(i, "R").Value
            resultats(12, count) = ws.Cells(i, "S").Value
            resultats(13, count) = ws.Cells(i, "T").Value
        End If
    Next i

    wb.Close False
    
    If count = 0 Then
        Call LogInfo("ChercherParametres", "0 param tre trouv  pour : " & codeEqt)
        Exit Function
    End If
    
    Call LogInfo("ChercherParametres", count & " param tres trouv s pour : " & codeEqt)
    ReDim Preserve resultats(1 To 13, 1 To count)
    ChercherParametres = resultats
    Exit Function

ErrorHandler:
    Call LogError("ChercherParametres", Err.Description, Err.Number)
    On Error Resume Next
    If Not wb Is Nothing Then wb.Close False
End Function

Function ChercherTousParametres(codeEqt As String, fichier As String) As Variant
    On Error GoTo ErrorHandler
    
    If fichier = "" Or Dir(fichier) = "" Then
        Call LogWarning("ChercherTousParametres", "Fichier introuvable (optionnel) : " & fichier)
        Exit Function
    End If
    
    Dim wb As Workbook
    Application.ScreenUpdating = False
    Set wb = Workbooks.Open(fichier, ReadOnly:=True, UpdateLinks:=0)
    wb.Windows(1).Visible = False
    
    If wb.Name <> Mid(fichier, InStrRev(fichier, "\") + 1) Then
        Call LogWarning("ChercherTousParametres", "Mauvais fichier ouvert : " & wb.Name)
        wb.Close False
        Exit Function
    End If
    
    Dim ws As Worksheet
    Set ws = wb.Worksheets(1)
    Call LogInfo("ChercherTousParametres", "Chemin demand  : [" & fichier & "]")
    Call LogInfo("ChercherTousParametres", "Fichier ouvert : [" & wb.Name & "]")
    Call LogInfo("ChercherTousParametres", "Ligne 3 col B : [" & ws.Cells(3, "B").Value & "]")
    
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.count, "B").End(xlUp).row
    
    Dim codeSansTrancheUC As String
    codeSansTrancheUC = UCase(Trim(SansTrancheCode(codeEqt)))
    
    
    Call LogInfo("ChercherTousParametres", "Ouverture fichier : " & fichier)
    Call LogInfo("ChercherTousParametres", "Fichier ouvert : " & wb.Name)
    Call LogInfo("ChercherTousParametres", "Nb lignes : " & lastRow)
    Call LogInfo("ChercherTousParametres", "Cherche : [" & codeSansTrancheUC & "]")
    Call LogInfo("ChercherTousParametres", "Exemple ligne 3 col B : [" & ws.Cells(3, "B").Value & "]")
    
    Dim resultats() As String
    Dim count As Integer: count = 0
    ReDim resultats(1 To 13, 1 To lastRow)


    Dim i As Long
    Dim raw As String, codeCompare As String
    For i = 3 To lastRow
        raw = ws.Cells(i, "B").Value
        If InStr(raw, "_") > 0 Then
            codeCompare = Mid(raw, InStr(raw, "_") + 1)
        Else
            codeCompare = SansTrancheCode(raw)
        End If

        If UCase(Trim(codeCompare)) = codeSansTrancheUC Then
            count = count + 1
            resultats(1, count) = ws.Cells(i, "G").Value
            resultats(2, count) = ws.Cells(i, "I").Value
            resultats(3, count) = ws.Cells(i, "L").Value
            resultats(4, count) = ws.Cells(i, "M").Value
            resultats(5, count) = ws.Cells(i, "P").Value
            resultats(6, count) = ws.Cells(i, "N").Value
            resultats(7, count) = ws.Cells(i, "O").Value
            resultats(8, count) = ws.Cells(i, "Q").Value
            resultats(9, count) = ws.Cells(i, "J").Value
            resultats(10, count) = ws.Cells(i, "K").Value
            resultats(11, count) = ws.Cells(i, "R").Value
            resultats(12, count) = ws.Cells(i, "S").Value
            resultats(13, count) = ws.Cells(i, "T").Value
        End If
    Next i

    wb.Close False

    If count = 0 Then Exit Function
    
    Call LogInfo("ChercherTousParametres", count & " param tres trouv s dans " & Mid(fichier, InStrRev(fichier, "\") + 1))
    ReDim Preserve resultats(1 To 13, 1 To count)
    ChercherTousParametres = resultats
    Exit Function

ErrorHandler:
    Call LogError("ChercherTousParametres", Err.Description, Err.Number)
    On Error Resume Next
    If Not wb Is Nothing Then wb.Close False
End Function

Function ChercherParametresAlternatifs(nomParam As String, codeEqt As String, chemin13E As String, cheminCAT As String) As Variant
    On Error GoTo ErrorHandler
    
    Dim chemins(1) As String
    chemins(0) = chemin13E
    chemins(1) = cheminCAT
    
    Dim codeSansTrancheUC As String
    codeSansTrancheUC = UCase(Trim(SansTrancheCode(codeEqt)))
    Dim nomParamUC As String
    nomParamUC = UCase(Trim(nomParam))
    
    Dim f As Integer
    For f = 0 To 1
        If chemins(f) = "" Or Dir(chemins(f)) = "" Then GoTo SuivantF
        
        Dim wb As Workbook
        On Error Resume Next
        Application.ScreenUpdating = False
        Set wb = Workbooks.Open(chemins(f), ReadOnly:=True, UpdateLinks:=0)
        wb.Windows(1).Visible = False
        On Error GoTo ErrorHandler
        If wb Is Nothing Then GoTo SuivantF
        
        Dim ws As Worksheet
        Set ws = wb.Worksheets(1)
        Dim lastRow As Long
        lastRow = ws.Cells(ws.Rows.count, "B").End(xlUp).row
        
        Dim i As Long
        Dim raw As String, codeCompare As String
        For i = 3 To lastRow
            raw = ws.Cells(i, "B").Value
            If InStr(raw, "_") > 0 Then
                codeCompare = Mid(raw, InStr(raw, "_") + 1)
            Else
                codeCompare = SansTrancheCode(raw)
            End If
            
            If UCase(Trim(codeCompare)) = codeSansTrancheUC Then
                If UCase(Trim(ws.Cells(i, "I").Value)) = nomParamUC Then
                    Dim res(1 To 13) As String
                    res(1) = ws.Cells(i, "G").Value
                    res(2) = ws.Cells(i, "I").Value
                    res(3) = ws.Cells(i, "L").Value
                    res(4) = ws.Cells(i, "M").Value
                    res(5) = ws.Cells(i, "P").Value
                    res(6) = ws.Cells(i, "N").Value
                    res(7) = ws.Cells(i, "O").Value
                    res(8) = ws.Cells(i, "Q").Value
                    res(9) = ws.Cells(i, "J").Value
                    res(10) = ws.Cells(i, "K").Value
                    res(11) = ws.Cells(i, "R").Value
                    res(12) = ws.Cells(i, "S").Value
                    res(13) = ws.Cells(i, "T").Value
                    wb.Close False
                    ChercherParametresAlternatifs = res
                    Exit Function
                End If
            End If
        Next i
        
        wb.Close False
SuivantF:
    Next f
    Exit Function

ErrorHandler:
    Call LogError("ChercherParametresAlternatifs", Err.Description, Err.Number)
    On Error Resume Next
    If Not wb Is Nothing Then wb.Close False
End Function


Function ChercherParametresSYGMA(codeEqt As String, fichier As String) As Variant
    On Error GoTo ErrorHandler
    If fichier = "" Or Dir(fichier) = "" Then
        Call LogWarning("ChercherParametresSYGMA", "Fichier SYGMA introuvable : " & fichier)
        Exit Function
    End If
    
    Dim wb As Workbook
    Application.ScreenUpdating = False
    Set wb = Workbooks.Open(fichier, ReadOnly:=True, UpdateLinks:=0)
    wb.Windows(1).Visible = False
    Dim ws As Worksheet
    Set ws = wb.Worksheets(1)
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.count, "A").End(xlUp).row
    
    Dim resultats() As String
    Dim count As Integer: count = 0
    ReDim resultats(1 To 5, 1 To lastRow)
    
    Dim codeSansTranche As String
    codeSansTranche = UCase(Trim(SansTrancheCode(codeEqt)))
    Dim codeUC As String
    codeUC = UCase(Trim(codeEqt))
    
    Dim i As Long
    For i = 2 To lastRow
        Dim codeCell As String
        codeCell = UCase(Trim(ws.Cells(i, "A").Value))
        
        If codeCell = codeUC Or UCase(Trim(SansTrancheCode(codeCell))) = codeSansTranche Then
            count = count + 1
            resultats(1, count) = ws.Cells(i, "Q").Value   ' Nom param tre
            resultats(2, count) = ws.Cells(i, "V").Value   ' Val basse
            resultats(3, count) = ws.Cells(i, "X").Value   ' Val normale
            resultats(4, count) = ws.Cells(i, "Z").Value   ' Val haute
            resultats(5, count) = ws.Cells(i, "AA").Value  ' Unit
        End If
    Next i
    
    wb.Close False
    
    If count = 0 Then
        Call LogInfo("ChercherParametresSYGMA", "0 param tre SYGMA pour : " & codeEqt)
        Exit Function
    End If
    
    Call LogInfo("ChercherParametresSYGMA", count & " param tres SYGMA pour : " & codeEqt)
    ReDim Preserve resultats(1 To 5, 1 To count)
    ChercherParametresSYGMA = resultats
    Exit Function

ErrorHandler:
    Call LogError("ChercherParametresSYGMA", Err.Description, Err.Number)
    On Error Resume Next
    If Not wb Is Nothing Then wb.Close False
End Function



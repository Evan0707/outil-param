' ===================================================
' MODULE: mod_Stats
' Comptage equipements SAL : sans params (PdF), 13E/CAT non vide
' Sans PdF = code present dans fichier SAL ref mais absent du fichier SAL-params
' ===================================================

Sub ComptageEquipements()
    On Error GoTo ErrorHandler

    Application.ScreenUpdating = False
    Application.StatusBar = "Chargement..."

    Dim wsOut As Worksheet
    On Error Resume Next
    Set wsOut = ThisWorkbook.Sheets("Statistiques")
    On Error GoTo ErrorHandler
    If wsOut Is Nothing Then
        Set wsOut = ThisWorkbook.Sheets.Add
        wsOut.Name = "Statistiques"
    End If
    wsOut.Cells.Clear

    wsOut.Cells(1, 1).Value = "Code Equipement"
    wsOut.Cells(1, 2).Value = "Tranche"
    wsOut.Cells(1, 3).Value = "PdF (params SAL)"
    wsOut.Cells(1, 4).Value = "13E"
    wsOut.Cells(1, 5).Value = "CAT"
    wsOut.Rows(1).Font.Bold = True
    wsOut.Rows(1).Interior.Color = RGB(200, 200, 200)

    Dim countTotal As Long, countSansPDF As Long, count13EOuCAT As Long
    Dim outRow As Long: outRow = 2

    Dim tranches(3) As String
    tranches(0) = "0"
    tranches(1) = "1"
    tranches(2) = "2"
    tranches(3) = "9"

    Dim t As Integer
    For t = 0 To 3
        Dim tranche As String: tranche = tranches(t)

        Dim fichierRef As String
        fichierRef = GetCheminFichierSAL_ParTranche(tranche)
        If Dir(fichierRef) = "" Then GoTo SuivantT

        Application.StatusBar = "Tranche " & tranche & "..."

        Dim dictSAL As Object
        Set dictSAL = ChargerCodesSAL(fichierRef)
        If dictSAL.count = 0 Then GoTo SuivantT

        ' Utilise un code representatif pour obtenir les chemins params/13E/CAT de cette tranche
        Dim codeRef As String: codeRef = dictSAL.Keys()(0)

        ' Pre-charger les codes presents dans SAL-params, 13E et CAT
        Dim dictPDF As Object
        Set dictPDF = ChargerCodesSALParams(GetCheminFichierParams(codeRef))
        Dim dict13E As Object
        Set dict13E = ChargerCodesFichierParams(GetCheminFichier13E(codeRef))
        Dim dictCAT As Object
        Set dictCAT = ChargerCodesFichierParams(GetCheminFichierCAT(codeRef))

        Dim codes As Variant: codes = dictSAL.Keys()
        Dim k As Integer
        For k = 0 To dictSAL.count - 1
            Dim codeEqt As String: codeEqt = codes(k)
            countTotal = countTotal + 1

            Dim hasPDF As Boolean: hasPDF = dictPDF.Exists(UCase(Trim(codeEqt)))

            Dim codeST As String: codeST = UCase(Trim(SansTrancheCode(codeEqt)))
            Dim has13E As Boolean: has13E = dict13E.Exists(codeST)
            Dim hasCAT As Boolean: hasCAT = dictCAT.Exists(codeST)

            wsOut.Cells(outRow, 1).Value = codeEqt
            wsOut.Cells(outRow, 2).Value = tranche

            If hasPDF Then
                wsOut.Cells(outRow, 3).Value = "OUI"
                wsOut.Cells(outRow, 3).Font.Color = RGB(0, 150, 0)
            Else
                wsOut.Cells(outRow, 3).Value = "NON"
                wsOut.Cells(outRow, 3).Font.Color = RGB(200, 0, 0)
                countSansPDF = countSansPDF + 1
            End If

            wsOut.Cells(outRow, 4).Value = IIf(has13E, "OUI", "NON")
            wsOut.Cells(outRow, 5).Value = IIf(hasCAT, "OUI", "NON")
            If has13E Then wsOut.Cells(outRow, 4).Font.Color = RGB(0, 150, 0)
            If hasCAT Then wsOut.Cells(outRow, 5).Font.Color = RGB(0, 150, 0)

            If has13E Or hasCAT Then count13EOuCAT = count13EOuCAT + 1

            outRow = outRow + 1
        Next k

SuivantT:
    Next t

    Call FermerFichiersExternes

    ' Resume
    outRow = outRow + 1
    wsOut.Cells(outRow, 1).Value = "TOTAL SAL"
    wsOut.Cells(outRow, 2).Value = countTotal
    wsOut.Rows(outRow).Font.Bold = True

    outRow = outRow + 1
    wsOut.Cells(outRow, 1).Value = "Sans PdF (params SAL)"
    wsOut.Cells(outRow, 1).Font.Bold = True
    wsOut.Cells(outRow, 2).Value = countSansPDF
    wsOut.Cells(outRow, 2).Font.Color = RGB(200, 0, 0)

    outRow = outRow + 1
    wsOut.Cells(outRow, 1).Value = "Avec 13E ou CAT"
    wsOut.Cells(outRow, 1).Font.Bold = True
    wsOut.Cells(outRow, 2).Value = count13EOuCAT
    wsOut.Cells(outRow, 2).Font.Color = RGB(0, 112, 192)

    wsOut.Columns.AutoFit
    wsOut.Activate

    Application.ScreenUpdating = True
    Application.StatusBar = False

    MsgBox "Comptage termin" & Chr(233) & "." & vbCrLf & _
           "Total SAL : " & countTotal & vbCrLf & _
           "Sans PdF : " & countSansPDF & vbCrLf & _
           "Avec 13E ou CAT : " & count13EOuCAT, vbInformation, "Statistiques"
    Exit Sub

ErrorHandler:
    Application.ScreenUpdating = True
    Application.StatusBar = False
    Call LogError("ComptageEquipements", Err.Description, Err.Number)
    MsgBox "Erreur. Consultez les Logs.", vbExclamation
End Sub


' Codes uniques depuis le fichier SAL reference (colonne F, ligne 2+)
Private Function ChargerCodesSAL(fichier As String) As Object
    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")
    dict.CompareMode = 1

    If fichier = "" Or Dir(fichier) = "" Then
        Set ChargerCodesSAL = dict
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

    Dim i As Long, code As String
    For i = 2 To lastRow
        code = UCase(Trim(ws.Cells(i, "F").Value))
        If code <> "" And Not dict.Exists(code) Then
            dict.Add code, code
        End If
    Next i

    wb.Close False
    Set ChargerCodesSAL = dict
End Function


' Codes uniques depuis le fichier SAL-params (colonne B, ligne 2+) - correspondance exacte
Private Function ChargerCodesSALParams(fichier As String) As Object
    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")
    dict.CompareMode = 1

    If fichier = "" Or Dir(fichier) = "" Then
        Set ChargerCodesSALParams = dict
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

    Dim i As Long, code As String
    For i = 2 To lastRow
        code = UCase(Trim(ws.Cells(i, "B").Value))
        If code <> "" And Not dict.Exists(code) Then
            dict.Add code, code
        End If
    Next i

    wb.Close False
    Set ChargerCodesSALParams = dict
End Function


' Codes sans-tranche presents dans un fichier 13E ou CAT (colonne B, ligne 3+)
Private Function ChargerCodesFichierParams(fichier As String) As Object
    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")
    dict.CompareMode = 1

    If fichier = "" Or Dir(fichier) = "" Then
        Set ChargerCodesFichierParams = dict
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

    Dim i As Long, raw As String, codeComp As String
    For i = 3 To lastRow
        raw = ws.Cells(i, "B").Value
        If InStr(raw, "_") > 0 Then
            codeComp = Mid(raw, InStr(raw, "_") + 1)
        Else
            codeComp = SansTrancheCode(raw)
        End If
        codeComp = UCase(Trim(codeComp))
        If codeComp <> "" And Not dict.Exists(codeComp) Then
            dict.Add codeComp, codeComp
        End If
    Next i

    wb.Close False
    Set ChargerCodesFichierParams = dict
End Function



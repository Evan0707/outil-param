' ===================================================
' MODULE: mod_Recherche
' ===================================================

Private m_hist(0 To 9) As String
Private m_nbHist As Integer
' ===================================================
' FAVORIS (persistants - registre Windows)
' ===================================================
Private Const FAV_APP As String = "OutilParam"
Private Const FAV_SECTION As String = "Favoris"
Private Const FAV_KEY As String = "liste"

Sub AjouterHistorique(code As String)
    code = Trim(code)
    If code = "" Then Exit Sub
    Dim i As Integer, j As Integer
    For i = 0 To m_nbHist - 1
        If UCase(m_hist(i)) = UCase(code) Then
            For j = i To m_nbHist - 2
                m_hist(j) = m_hist(j + 1)
            Next j
            m_nbHist = m_nbHist - 1
            Exit For
        End If
    Next i
    Dim top As Integer: top = IIf(m_nbHist < 9, m_nbHist, 9)
    For i = top To 1 Step -1
        m_hist(i) = m_hist(i - 1)
    Next i
    m_hist(0) = code
    If m_nbHist < 10 Then m_nbHist = m_nbHist + 1
End Sub

Function LireHistorique() As Variant
    If m_nbHist = 0 Then LireHistorique = Empty: Exit Function
    Dim result() As String
    ReDim result(0 To m_nbHist - 1)
    Dim i As Integer
    For i = 0 To m_nbHist - 1
        result(i) = m_hist(i)
    Next i
    LireHistorique = result
End Function



Function LireFavoris() As Variant
    Dim liste As String
    liste = GetSetting(FAV_APP, FAV_SECTION, FAV_KEY, "")
    If Trim(liste) = "" Then LireFavoris = Empty: Exit Function
    LireFavoris = Split(liste, "|")
End Function

Function EstFavori(code As String) As Boolean
    Dim fav As Variant: fav = LireFavoris()
    If IsEmpty(fav) Then Exit Function
    Dim i As Long
    For i = 0 To UBound(fav)
        If UCase(Trim(fav(i))) = UCase(Trim(code)) Then
            EstFavori = True: Exit Function
        End If
    Next i
End Function

Sub AjouterFavori(code As String)
    code = Trim(code)
    If code = "" Then Exit Sub
    If EstFavori(code) Then Exit Sub
    Dim liste As String
    liste = GetSetting(FAV_APP, FAV_SECTION, FAV_KEY, "")
    If liste <> "" Then liste = liste & "|"
    SaveSetting FAV_APP, FAV_SECTION, FAV_KEY, liste & code
End Sub

Sub RetirerFavori(code As String)
    code = Trim(code)
    If code = "" Then Exit Sub
    Dim fav As Variant: fav = LireFavoris()
    If IsEmpty(fav) Then Exit Sub
    Dim res As String, i As Long
    For i = 0 To UBound(fav)
        If UCase(Trim(fav(i))) <> UCase(code) And Trim(fav(i)) <> "" Then
            If res <> "" Then res = res & "|"
            res = res & fav(i)
        End If
    Next i
    SaveSetting FAV_APP, FAV_SECTION, FAV_KEY, res
End Sub

' ===================================================
' HELPERS ADO  (lecture sans Workbooks.Open = gain majeur sur reseau)
'
' LireParADO retourne arr(colIdx, rowIdx) en base 0 (format GetRows ADO)
'   row 0    = ligne d-en-tete Excel
'   colIdx   : A=0  B=1  F=5  G=6  I=8  J=9  K=10  L=11  M=12
'              N=13 O=14 P=15 Q=16 R=17 S=18  T=19  V=21  X=23  Z=25  AA=26
'
' Si ACE.OLEDB non disponible => LireParADO = Empty => fallback Workbooks.Open
' ===================================================

Private Function StringConnADO(fichier As String) As String
    Dim ext As String: ext = LCase(Mid(fichier, InStrRev(fichier, ".") + 1))
    Dim xlProps As String
    Select Case ext
        Case "xlsx", "xlsm": xlProps = "Excel 12.0 Xml"
        Case "xlsb":         xlProps = "Excel 12.0"
        Case Else:           xlProps = "Excel 8.0"
    End Select
    StringConnADO = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & fichier & _
                    ";Extended Properties=""" & xlProps & ";HDR=NO;IMEX=1"""
End Function

Private Function OngletADO(conn As Object) As String
    On Error Resume Next
    Dim rs As Object
    Set rs = conn.OpenSchema(20)
    Do While Not rs.EOF
        Dim n As String: n = CStr(rs.Fields("TABLE_NAME").Value)
        If Right(n, 1) = "$" Then OngletADO = n: Exit Do
        rs.MoveNext
    Loop
    If Not rs Is Nothing Then rs.Close
End Function

Private Function LireParADO(fichier As String, nbCols As Integer) As Variant
    On Error GoTo Echec
    Dim conn As Object
    Set conn = CreateObject("ADODB.Connection")
    conn.Open StringConnADO(fichier)

    Dim onglet As String: onglet = OngletADO(conn)
    If onglet = "" Then GoTo Echec

    Dim cols As String, k As Integer
    For k = 1 To nbCols
        If k > 1 Then cols = cols & ","
        cols = cols & "F" & k
    Next k

    Dim rs As Object
    Set rs = CreateObject("ADODB.Recordset")
    rs.Open "SELECT " & cols & " FROM [" & onglet & "]", conn, 3, 1

    If Not rs.EOF Then LireParADO = rs.GetRows()
    rs.Close: conn.Close
    Exit Function
Echec:
    Dim errMsg As String: errMsg = Err.Description & " (" & Err.Number & ")"
    On Error Resume Next
    If Not rs Is Nothing Then rs.Close
    If Not conn Is Nothing Then conn.Close
    Call LogWarning("LireParADO", "Echec ADO [" & Mid(fichier, InStrRev(fichier, "\") + 1) & "] : " & errMsg)
    LireParADO = Empty
End Function

' Fallback Workbooks.Open (si ACE.OLEDB non dispo)
Private Function OuvrirFichierRapide(fichier As String) As Workbook
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
    Application.ScreenUpdating = False
    Set OuvrirFichierRapide = Workbooks.Open(fichier, ReadOnly:=True, UpdateLinks:=0)
    OuvrirFichierRapide.Windows(1).Visible = False
    Application.EnableEvents = True
End Function

' ===================================================
' FORMULAIRES
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

    Call AjouterHistorique(codeEqt)

    Dim frm As New resultats
    frm.SetSources popup.loadSAL, popup.load13E, popup.loadCAT, popup.loadSYGMA
    frm.InitAvecCode codeEqt
    frm.Show

    Exit Sub
ErrorHandler:
    Call LogError("RechercherEquipement", Err.Description, Err.Number)
    Call RestaurerEtatExcel
    MsgBox "Erreur recherche. Consultez les Logs.", vbExclamation
End Sub

' ===================================================
' FONCTIONS DE LECTURE DES FICHIERS SOURCES
' Chaque fonction essaie ADO en premier, fallback Workbooks.Open
' ===================================================

Function EquipementExiste(codeEqt As String, fichier As String) As Boolean
    On Error GoTo ErrorHandler

    If fichier = "" Then
        Call LogWarning("EquipementExiste", "Chemin vide pour : " & codeEqt): Exit Function
    End If
    If Dir(fichier) = "" Then
        Call LogWarning("EquipementExiste", "Introuvable : " & fichier)
        MsgBox "Fichier introuvable : " & fichier, vbExclamation: Exit Function
    End If

    Dim codeUC As String: codeUC = UCase(Trim(codeEqt))
    Dim i As Long

    ' --- ADO ---
    Dim arr As Variant
    arr = LireParADO(fichier, 6)
    If Not IsEmpty(arr) Then
        For i = 1 To UBound(arr, 2)
            If UCase(Trim(CStr(arr(5, i)))) = codeUC Then EquipementExiste = True: Exit Function
        Next i
        Exit Function
    End If

    ' --- Fallback ---
    Dim wb As Workbook
    Set wb = OuvrirFichierRapide(fichier)
    Dim ws As Worksheet: Set ws = wb.Worksheets(1)
    Dim lastRow As Long: lastRow = ws.Cells(ws.Rows.count, 6).End(xlUp).row
    Dim arrWB As Variant
    arrWB = ws.Range(ws.Cells(1, 1), ws.Cells(lastRow, 6)).Value
    wb.Close False: Application.ScreenUpdating = True
    For i = 2 To lastRow
        If UCase(Trim(CStr(arrWB(i, 6)))) = codeUC Then EquipementExiste = True: Exit Function
    Next i
    Exit Function
ErrorHandler:
    Call LogError("EquipementExiste", Err.Description, Err.Number)
    On Error Resume Next
    If Not wb Is Nothing Then wb.Close False
    Application.EnableEvents = True: Application.ScreenUpdating = True
End Function

Function ChercherParametres(codeEqt As String, fichier As String) As Variant
    On Error GoTo ErrorHandler

    If fichier = "" Or Dir(fichier) = "" Then
        Call LogWarning("ChercherParametres", "Introuvable : " & fichier): Exit Function
    End If

    Dim codeUC As String: codeUC = UCase(Trim(codeEqt))
    Dim resultats() As String
    Dim count As Integer: count = 0
    Dim i As Long
    ' Mapping colonnes (0-base ADO) : B=1 G=6 I=8 J=9 K=10 L=11 M=12 N=13 O=14 P=15 Q=16 R=17 S=18 T=19

    ' --- ADO ---
    Dim arr As Variant
    arr = LireParADO(fichier, 20)
    If Not IsEmpty(arr) Then
        ReDim resultats(1 To 13, 1 To UBound(arr, 2))
        For i = 1 To UBound(arr, 2)
            If UCase(Trim(CStr(arr(1, i)))) = codeUC Then
                count = count + 1
                resultats(1, count) = CStr(arr(6, i)):  resultats(2, count) = CStr(arr(8, i))
                resultats(3, count) = CStr(arr(11, i)): resultats(4, count) = CStr(arr(12, i))
                resultats(5, count) = CStr(arr(15, i)): resultats(6, count) = CStr(arr(13, i))
                resultats(7, count) = CStr(arr(14, i)): resultats(8, count) = CStr(arr(16, i))
                resultats(9, count) = CStr(arr(9, i)):  resultats(10, count) = CStr(arr(10, i))
                resultats(11, count) = CStr(arr(17, i)): resultats(12, count) = CStr(arr(18, i))
                resultats(13, count) = CStr(arr(19, i))
            End If
        Next i
        GoTo Finaliser
    End If

    ' --- Fallback ---
    Dim wb As Workbook
    Set wb = OuvrirFichierRapide(fichier)
    Dim ws As Worksheet: Set ws = wb.Worksheets(1)
    Dim lastRow As Long: lastRow = ws.Cells(ws.Rows.count, 2).End(xlUp).row
    Dim arrWB As Variant
    arrWB = ws.Range(ws.Cells(1, 1), ws.Cells(lastRow, 20)).Value
    wb.Close False: Application.ScreenUpdating = True
    ReDim resultats(1 To 13, 1 To lastRow)
    For i = 2 To lastRow
        If UCase(Trim(CStr(arrWB(i, 2)))) = codeUC Then
            count = count + 1
            resultats(1, count) = CStr(arrWB(i, 7)):  resultats(2, count) = CStr(arrWB(i, 9))
            resultats(3, count) = CStr(arrWB(i, 12)): resultats(4, count) = CStr(arrWB(i, 13))
            resultats(5, count) = CStr(arrWB(i, 16)): resultats(6, count) = CStr(arrWB(i, 14))
            resultats(7, count) = CStr(arrWB(i, 15)): resultats(8, count) = CStr(arrWB(i, 17))
            resultats(9, count) = CStr(arrWB(i, 10)): resultats(10, count) = CStr(arrWB(i, 11))
            resultats(11, count) = CStr(arrWB(i, 18)): resultats(12, count) = CStr(arrWB(i, 19))
            resultats(13, count) = CStr(arrWB(i, 20))
        End If
    Next i

Finaliser:
    If count = 0 Then
        Call LogInfo("ChercherParametres", "0 param tre trouv  pour : " & codeEqt): Exit Function
    End If
    Call LogInfo("ChercherParametres", count & " param tres trouv s pour : " & codeEqt)
    ReDim Preserve resultats(1 To 13, 1 To count)
    ChercherParametres = resultats
    Exit Function
ErrorHandler:
    Call LogError("ChercherParametres", Err.Description, Err.Number)
    On Error Resume Next
    If Not wb Is Nothing Then wb.Close False
    Application.EnableEvents = True: Application.ScreenUpdating = True
End Function

Function ChercherTousParametres(codeEqt As String, fichier As String) As Variant
    On Error GoTo ErrorHandler

    If fichier = "" Or Dir(fichier) = "" Then
        Call LogWarning("ChercherTousParametres", "Introuvable (optionnel) : " & fichier): Exit Function
    End If

    Dim codeSansTrancheUC As String: codeSansTrancheUC = UCase(Trim(SansTrancheCode(codeEqt)))
    Dim resultats() As String
    Dim count As Integer: count = 0
    Dim i As Long, raw As String, codeCompare As String

    ' --- ADO ---
    Dim arr As Variant
    arr = LireParADO(fichier, 20)
    If Not IsEmpty(arr) Then
        Call LogInfo("ChercherTousParametres", Mid(fichier, InStrRev(fichier, "\") + 1) & " - ADO - " & UBound(arr, 2) & " lignes")
        ReDim resultats(1 To 13, 1 To UBound(arr, 2))
        For i = 2 To UBound(arr, 2)
            raw = CStr(arr(1, i))
            If InStr(raw, "_") > 0 Then
                codeCompare = Mid(raw, InStr(raw, "_") + 1)
            Else
                codeCompare = SansTrancheCode(raw)
            End If
            If UCase(Trim(codeCompare)) = codeSansTrancheUC Then
                count = count + 1
                resultats(1, count) = CStr(arr(6, i)):  resultats(2, count) = CStr(arr(8, i))
                resultats(3, count) = CStr(arr(11, i)): resultats(4, count) = CStr(arr(12, i))
                resultats(5, count) = CStr(arr(15, i)): resultats(6, count) = CStr(arr(13, i))
                resultats(7, count) = CStr(arr(14, i)): resultats(8, count) = CStr(arr(16, i))
                resultats(9, count) = CStr(arr(9, i)):  resultats(10, count) = CStr(arr(10, i))
                resultats(11, count) = CStr(arr(17, i)): resultats(12, count) = CStr(arr(18, i))
                resultats(13, count) = CStr(arr(19, i))
            End If
        Next i
        GoTo Finaliser
    End If

    ' --- Fallback ---
    Dim wb As Workbook
    Set wb = OuvrirFichierRapide(fichier)
    If wb.Name <> Mid(fichier, InStrRev(fichier, "\") + 1) Then
        Call LogWarning("ChercherTousParametres", "Mauvais fichier : " & wb.Name)
        wb.Close False: Application.ScreenUpdating = True: Exit Function
    End If
    Dim ws As Worksheet: Set ws = wb.Worksheets(1)
    Dim lastRow As Long: lastRow = ws.Cells(ws.Rows.count, 2).End(xlUp).row
    Call LogInfo("ChercherTousParametres", Mid(fichier, InStrRev(fichier, "\") + 1) & " - WB - " & lastRow & " lignes")
    Dim arrWB As Variant
    arrWB = ws.Range(ws.Cells(1, 1), ws.Cells(lastRow, 20)).Value
    wb.Close False: Application.ScreenUpdating = True
    ReDim resultats(1 To 13, 1 To lastRow)
    For i = 3 To lastRow
        raw = CStr(arrWB(i, 2))
        If InStr(raw, "_") > 0 Then
            codeCompare = Mid(raw, InStr(raw, "_") + 1)
        Else
            codeCompare = SansTrancheCode(raw)
        End If
        If UCase(Trim(codeCompare)) = codeSansTrancheUC Then
            count = count + 1
            resultats(1, count) = CStr(arrWB(i, 7)):  resultats(2, count) = CStr(arrWB(i, 9))
            resultats(3, count) = CStr(arrWB(i, 12)): resultats(4, count) = CStr(arrWB(i, 13))
            resultats(5, count) = CStr(arrWB(i, 16)): resultats(6, count) = CStr(arrWB(i, 14))
            resultats(7, count) = CStr(arrWB(i, 15)): resultats(8, count) = CStr(arrWB(i, 17))
            resultats(9, count) = CStr(arrWB(i, 10)): resultats(10, count) = CStr(arrWB(i, 11))
            resultats(11, count) = CStr(arrWB(i, 18)): resultats(12, count) = CStr(arrWB(i, 19))
            resultats(13, count) = CStr(arrWB(i, 20))
        End If
    Next i

Finaliser:
    If count = 0 Then Exit Function
    Call LogInfo("ChercherTousParametres", count & " param tres trouv s dans " & Mid(fichier, InStrRev(fichier, "\") + 1))
    ReDim Preserve resultats(1 To 13, 1 To count)
    ChercherTousParametres = resultats
    Exit Function
ErrorHandler:
    Call LogError("ChercherTousParametres", Err.Description, Err.Number)
    On Error Resume Next
    If Not wb Is Nothing Then wb.Close False
    Application.EnableEvents = True: Application.ScreenUpdating = True
End Function

Function ChercherParametresAlternatifs(nomParam As String, codeEqt As String, chemin13E As String, cheminCAT As String) As Variant
    On Error GoTo ErrorHandler

    Dim chemins(1) As String
    chemins(0) = chemin13E: chemins(1) = cheminCAT
    Dim codeSansTrancheUC As String: codeSansTrancheUC = UCase(Trim(SansTrancheCode(codeEqt)))
    Dim nomParamUC As String: nomParamUC = UCase(Trim(nomParam))

    Dim f As Integer
    For f = 0 To 1
        If chemins(f) = "" Or Dir(chemins(f)) = "" Then GoTo SuivantF

        Dim i As Long, raw As String, codeCompare As String
        Dim res(1 To 13) As String

        ' --- ADO ---
        Dim arr As Variant
        arr = LireParADO(chemins(f), 20)
        If Not IsEmpty(arr) Then
            For i = 2 To UBound(arr, 2)
                raw = CStr(arr(1, i))
                If InStr(raw, "_") > 0 Then
                    codeCompare = Mid(raw, InStr(raw, "_") + 1)
                Else
                    codeCompare = SansTrancheCode(raw)
                End If
                If UCase(Trim(codeCompare)) = codeSansTrancheUC Then
                    If UCase(Trim(CStr(arr(8, i)))) = nomParamUC Then
                        res(1) = CStr(arr(6, i)):  res(2) = CStr(arr(8, i))
                        res(3) = CStr(arr(11, i)): res(4) = CStr(arr(12, i))
                        res(5) = CStr(arr(15, i)): res(6) = CStr(arr(13, i))
                        res(7) = CStr(arr(14, i)): res(8) = CStr(arr(16, i))
                        res(9) = CStr(arr(9, i)):  res(10) = CStr(arr(10, i))
                        res(11) = CStr(arr(17, i)): res(12) = CStr(arr(18, i))
                        res(13) = CStr(arr(19, i))
                        ChercherParametresAlternatifs = res: Exit Function
                    End If
                End If
            Next i
            GoTo SuivantF
        End If

        ' --- Fallback ---
        Dim wb As Workbook
        On Error Resume Next
        Set wb = OuvrirFichierRapide(chemins(f))
        On Error GoTo ErrorHandler
        If wb Is Nothing Then GoTo SuivantF
        Dim ws As Worksheet: Set ws = wb.Worksheets(1)
        Dim lastRow As Long: lastRow = ws.Cells(ws.Rows.count, 2).End(xlUp).row
        Dim arrWB As Variant
        arrWB = ws.Range(ws.Cells(1, 1), ws.Cells(lastRow, 20)).Value
        wb.Close False: Application.ScreenUpdating = True
        For i = 3 To lastRow
            raw = CStr(arrWB(i, 2))
            If InStr(raw, "_") > 0 Then
                codeCompare = Mid(raw, InStr(raw, "_") + 1)
            Else
                codeCompare = SansTrancheCode(raw)
            End If
            If UCase(Trim(codeCompare)) = codeSansTrancheUC Then
                If UCase(Trim(CStr(arrWB(i, 9)))) = nomParamUC Then
                    res(1) = CStr(arrWB(i, 7)):  res(2) = CStr(arrWB(i, 9))
                    res(3) = CStr(arrWB(i, 12)): res(4) = CStr(arrWB(i, 13))
                    res(5) = CStr(arrWB(i, 16)): res(6) = CStr(arrWB(i, 14))
                    res(7) = CStr(arrWB(i, 15)): res(8) = CStr(arrWB(i, 17))
                    res(9) = CStr(arrWB(i, 10)): res(10) = CStr(arrWB(i, 11))
                    res(11) = CStr(arrWB(i, 18)): res(12) = CStr(arrWB(i, 19))
                    res(13) = CStr(arrWB(i, 20))
                    ChercherParametresAlternatifs = res: Exit Function
                End If
            End If
        Next i
SuivantF:
    Next f
    Exit Function
ErrorHandler:
    Call LogError("ChercherParametresAlternatifs", Err.Description, Err.Number)
    On Error Resume Next
    If Not wb Is Nothing Then wb.Close False
    Application.EnableEvents = True: Application.ScreenUpdating = True
End Function

Function ChercherParametresSYGMA(codeEqt As String, fichier As String) As Variant
    On Error GoTo ErrorHandler
    If fichier = "" Or Dir(fichier) = "" Then
        Call LogWarning("ChercherParametresSYGMA", "SYGMA introuvable : " & fichier): Exit Function
    End If

    Dim codeSansTranche As String: codeSansTranche = UCase(Trim(SansTrancheCode(codeEqt)))
    Dim codeUC As String: codeUC = UCase(Trim(codeEqt))
    Dim resultats() As String
    Dim count As Integer: count = 0
    Dim i As Long, codeCell As String
    ' Mapping (0-base ADO) : A=0  Q=16  V=21  X=23  Z=25  AA=26

    ' --- ADO ---
    Dim arr As Variant
    arr = LireParADO(fichier, 27)
    If Not IsEmpty(arr) Then
        ReDim resultats(1 To 5, 1 To UBound(arr, 2))
        For i = 1 To UBound(arr, 2)
            codeCell = UCase(Trim(CStr(arr(0, i))))
            If codeCell = codeUC Or UCase(Trim(SansTrancheCode(codeCell))) = codeSansTranche Then
                count = count + 1
                resultats(1, count) = CStr(arr(16, i)) ' Q
                resultats(2, count) = CStr(arr(21, i)) ' V
                resultats(3, count) = CStr(arr(23, i)) ' X
                resultats(4, count) = CStr(arr(25, i)) ' Z
                resultats(5, count) = CStr(arr(26, i)) ' AA
            End If
        Next i
        GoTo Finaliser
    End If

    ' --- Fallback ---
    Dim wb As Workbook
    Set wb = OuvrirFichierRapide(fichier)
    Dim ws As Worksheet: Set ws = wb.Worksheets(1)
    Dim lastRow As Long: lastRow = ws.Cells(ws.Rows.count, 1).End(xlUp).row
    Dim arrWB As Variant
    arrWB = ws.Range(ws.Cells(1, 1), ws.Cells(lastRow, 27)).Value
    wb.Close False: Application.ScreenUpdating = True
    ReDim resultats(1 To 5, 1 To lastRow)
    For i = 2 To lastRow
        codeCell = UCase(Trim(CStr(arrWB(i, 1))))
        If codeCell = codeUC Or UCase(Trim(SansTrancheCode(codeCell))) = codeSansTranche Then
            count = count + 1
            resultats(1, count) = CStr(arrWB(i, 17)): resultats(2, count) = CStr(arrWB(i, 22))
            resultats(3, count) = CStr(arrWB(i, 24)): resultats(4, count) = CStr(arrWB(i, 26))
            resultats(5, count) = CStr(arrWB(i, 27))
        End If
    Next i

Finaliser:
    If count = 0 Then
        Call LogInfo("ChercherParametresSYGMA", "0 param tre SYGMA pour : " & codeEqt): Exit Function
    End If
    Call LogInfo("ChercherParametresSYGMA", count & " param tres SYGMA pour : " & codeEqt)
    ReDim Preserve resultats(1 To 5, 1 To count)
    ChercherParametresSYGMA = resultats
    Exit Function

ErrorHandler:
    Call LogError("ChercherParametresSYGMA", Err.Description, Err.Number)
    On Error Resume Next
    If Not wb Is Nothing Then wb.Close False
    Application.EnableEvents = True: Application.ScreenUpdating = True
End Function





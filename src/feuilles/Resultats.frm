Option Explicit

Private WithEvents m_btnRecherche As MSForms.CommandButton
Private m_loadSAL As Boolean
Private m_load13E As Boolean
Private m_loadCAT As Boolean
Private m_loadSYGMA As Boolean
Private m_CodeEqt As String
Private m_HTMLSAL As String
Private m_HTML13E As String
Private m_HTMLCAT As String
Private m_HTMLSYGMA As String

' API Windows pour taille ecran reelle + DPI
#If VBA7 Then
    Private Declare PtrSafe Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long
    Private Declare PtrSafe Function GetDC Lib "user32" (ByVal hwnd As LongPtr) As LongPtr
    Private Declare PtrSafe Function GetDeviceCaps Lib "gdi32" (ByVal hdc As LongPtr, ByVal nIndex As Long) As Long
    Private Declare PtrSafe Function ReleaseDC Lib "user32" (ByVal hwnd As LongPtr, ByVal hdc As LongPtr) As Long
    Private Declare PtrSafe Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As LongPtr
    Private Declare PtrSafe Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As LongPtr, ByVal nIndex As Long) As Long
    Private Declare PtrSafe Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As LongPtr, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
    Private Declare PtrSafe Function SetWindowPos Lib "user32" (ByVal hwnd As LongPtr, ByVal hWndInsertAfter As LongPtr, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal uFlags As Long) As Long
#Else
    Private Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long
    Private Declare Function GetDC Lib "user32" (ByVal hwnd As Long) As Long
    Private Declare Function GetDeviceCaps Lib "gdi32" (ByVal hdc As Long, ByVal nIndex As Long) As Long
    Private Declare Function ReleaseDC Lib "user32" (ByVal hwnd As Long, ByVal hdc As Long) As Long
    Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
    Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
    Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
    Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal uFlags As Long) As Long
#End If
Private Const SM_CXSCREEN As Long = 0
Private Const SM_CYSCREEN As Long = 1
Private Const LOGPIXELSX As Long = 88
Private Const GWL_STYLE As Long = -16
Private Const WS_MAXIMIZEBOX As Long = &H10000
Private Const WS_THICKFRAME As Long = &H40000
Private Const SWP_NOMOVE As Long = &H2
Private Const SWP_NOSIZE As Long = &H1
Private Const SWP_NOZORDER As Long = &H4
Private Const SWP_FRAMECHANGED As Long = &H20

Private Sub TailleEcranPoints(ByRef w As Single, ByRef h As Single)
    Dim px As Long: px = GetSystemMetrics(SM_CXSCREEN)
    Dim py As Long: py = GetSystemMetrics(SM_CYSCREEN)
    Dim dpi As Long
    #If VBA7 Then
        Dim dc As LongPtr: dc = GetDC(0)
        dpi = GetDeviceCaps(dc, LOGPIXELSX)
        ReleaseDC 0, dc
    #Else
        Dim dc As Long: dc = GetDC(0)
        dpi = GetDeviceCaps(dc, LOGPIXELSX)
        ReleaseDC 0, dc
    #End If
    If dpi = 0 Then dpi = 96
    w = px * 72! / dpi
    h = py * 72! / dpi
End Sub



' INITIALISATION

Public Sub InitAvecCode(codeEqt As String)
    On Error GoTo ErrorHandler
    m_CodeEqt = codeEqt
    lblCode.Caption = " quipement : " & m_CodeEqt
    lblCode.Font.Bold = True
    lblCode.Font.Size = 14
    lblCode.TextAlign = fmTextAlignCenter
    Call ChargerDonnees
    Exit Sub
ErrorHandler:
    Call LogError("InitAvecCode", Err.Description, Err.Number)
End Sub


Private Sub UserForm_Initialize()
    On Error GoTo ErrorHandler

    Me.Caption = "R" & Chr(233) & "sultats " & Chr(233) & "quipement"
    mpOnglets.Pages(0).Caption = "SAL"
    mpOnglets.Pages(1).Caption = "13E"
    mpOnglets.Pages(2).Caption = "CAT"
    mpOnglets.Pages(3).Caption = "SYGMA"

    ' Taille reelle de l'ecran (pixels -> points selon DPI)
    Dim w As Single, h As Single
    Call TailleEcranPoints(w, h)
    w = w - 10
    h = h - 40
    Me.Width = w
    Me.Height = h

    Call AjusterControles(w, h)

    m_loadSAL = True
    m_load13E = True
    m_loadCAT = True
    m_loadSYGMA = True

    Set m_btnRecherche = Me.Controls.Add("Forms.CommandButton.1", "btnRecherche")
    With m_btnRecherche
        .Caption = "Nouvelle recherche"
        .Width = 140
        .Height = 26
        .Left = Me.InsideWidth - (m_btnRecherche.Width + 10)
        .top = m_btnRecherche.Height - 20
        .Font.Size = 10
    End With

    Exit Sub
ErrorHandler:
    Call LogError("UserForm_Initialize", Err.Description, Err.Number)
End Sub

Private Sub AjusterControles(w As Single, h As Single)
    mpOnglets.Width = w
    mpOnglets.Height = h - 100
    wbTableau.Width = mpOnglets.Width - 10
    wbTableau.Height = mpOnglets.Height - 25
    If Not m_btnRecherche Is Nothing Then
        m_btnRecherche.top = m_btnRecherche.Height - 20
        m_btnRecherche.Left = Me.InsideWidth - (m_btnRecherche.Width + 10)
    End If
End Sub

Private Sub UserForm_Activate()
    ' Active le bouton maximiser natif Windows dans la barre de titre
    #If VBA7 Then
        Dim hwnd As LongPtr
    #Else
        Dim hwnd As Long
    #End If
    hwnd = FindWindow("ThunderRT6DFrame", Me.Caption)
    If hwnd = 0 Then hwnd = FindWindow("ThunderDFrame", Me.Caption)
    If hwnd <> 0 Then
        Dim style As Long
        style = GetWindowLong(hwnd, GWL_STYLE)
        style = style Or WS_MAXIMIZEBOX Or WS_THICKFRAME
        SetWindowLong hwnd, GWL_STYLE, style
        SetWindowPos hwnd, 0, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE Or SWP_NOZORDER Or SWP_FRAMECHANGED
    End If
End Sub

Private Sub UserForm_Resize()
    On Error Resume Next
    Call AjusterControles(Me.Width, Me.Height)
End Sub

' CHARGEMENT DONN ES

Private Sub ChargerDonnees()
    On Error GoTo ErrorHandler
    Application.ScreenUpdating = False

    Dim loadSAL As Boolean: loadSAL = m_loadSAL
    Dim load13E As Boolean: load13E = m_load13E
    Dim loadCAT As Boolean: loadCAT = m_loadCAT
    Dim loadSYGMA As Boolean: loadSYGMA = m_loadSYGMA

    Dim results13E As Variant
    If load13E Then
        results13E = ChercherTousParametres(m_CodeEqt, GetCheminFichier13E(m_CodeEqt))
        Call FermerFichiersExternes
    End If

    Dim resultsCAT As Variant
    If loadCAT Then
        resultsCAT = ChercherTousParametres(m_CodeEqt, GetCheminFichierCAT(m_CodeEqt))
        Call FermerFichiersExternes
    End If

    Dim resultsSAL As Variant
    If loadSAL Then
        resultsSAL = ChercherParametres(m_CodeEqt, GetCheminFichierParams(m_CodeEqt))
        m_HTMLSAL = GenererHTMLSALAvecTooltip(resultsSAL, results13E, resultsCAT)
        Call FermerFichiersExternes
    Else
        m_HTMLSAL = "<html><body><p style='color:gray;font-style:italic'>SAL non charg </p></body></html>"
    End If

    If load13E Then
        m_HTML13E = GenererHTMLTableau(results13E, "13E", "#00B050", "#f0fff0", "#d0ffd0", _
        Array("Num s q.", "Param tre", "Val Normale", "Unit  mesure", "Seuil TB", "Seuil B", "Seuil H", "Seuil TH", "Val REF", "Unit  REF", "Commentaire 1", "Commentaire 2", "Commentaire 3"))
    Else
        m_HTML13E = "<html><body><p style='color:gray;font-style:italic'>13E non charg </p></body></html>"
    End If

    If loadCAT Then
        m_HTMLCAT = GenererHTMLTableau(resultsCAT, "CAT", "#FF6600", "#fff5f0", "#ffd0b0", _
        Array("Num s q.", "Param tre", "Val Normale", "Unit  mesure", "Seuil TB", "Seuil B", "Seuil H", "Seuil TH", "Val REF", "Unit  REF", "Commentaire 1", "Commentaire 2", "Commentaire 3"))
    Else
        m_HTMLCAT = "<html><body><p style='color:gray;font-style:italic'>CAT non charg </p></body></html>"
    End If

    If loadSYGMA Then
        Dim resultsSYGMA As Variant
        resultsSYGMA = ChercherParametresSYGMA(m_CodeEqt, GetCheminFichierSYGMA())
        m_HTMLSYGMA = GenererHTMLTableau(resultsSYGMA, "SYGMA", "#6600CC", "#f5f0ff", "#e0d0ff", _
        Array("Param tre", "Val Basse", "Val Normale", "Val Haute", "Unit "))
    Else
        m_HTMLSYGMA = "<html><body><p style='color:gray;font-style:italic'>SYGMA non charg </p></body></html>"
    End If

    Application.ScreenUpdating = True
    Call AfficherHTML(m_HTMLSAL)
    Exit Sub
ErrorHandler:
    Application.ScreenUpdating = True
    Call LogError("ChargerDonnees", Err.Description, Err.Number)
End Sub
' NAVIGATION ONGLETS

Private Sub mpOnglets_Change()
    On Error GoTo ErrorHandler
    Select Case mpOnglets.Value
        Case 0: Call AfficherHTML(m_HTMLSAL)
        Case 1: Call AfficherHTML(m_HTML13E)
        Case 2: Call AfficherHTML(m_HTMLCAT)
        Case 3: Call AfficherHTML(m_HTMLSYGMA)
    End Select
    Exit Sub
ErrorHandler:
    Call LogError("mpOnglets_Change", Err.Description, Err.Number)
End Sub

' AFFICHAGE HTML

Private Sub AfficherHTML(html As String)
    On Error GoTo ErrorHandler
    
    Dim cheminTemp As String
    cheminTemp = Environ("TEMP") & "\edf_tableau.html"
    
    Dim fileNum As Integer
    fileNum = FreeFile
    Open cheminTemp For Output As #fileNum
    Print #fileNum, html
    Close #fileNum
    
    wbTableau.Navigate cheminTemp
    
    Dim t As Double: t = Timer
    Do While wbTableau.ReadyState <> 4
        DoEvents
        If Timer - t > 5 Then Exit Do
    Loop
    Exit Sub
ErrorHandler:
    Call LogError("AfficherHTML", Err.Description, Err.Number)
End Sub

' G N RATION HTML SAL AVEC TOOLTIP

Private Function GenererHTMLSALAvecTooltip(resultsSAL As Variant, results13E As Variant, resultsCAT As Variant) As String
    Dim html As String
    
    Dim css As String
    css = "body{font-family:Calibri,Arial;font-size:12px;margin:5px;}"
    css = css & "table{border-collapse:collapse;width:100%;}"
    css = css & "th{background-color:#1C6EA4;color:white;padding:6px 8px;text-align:left;border:1px solid #ccc;}"
    css = css & "td{padding:5px 8px;border:1px solid #ddd;}"
    css = css & "tr:nth-child(even){background-color:#f2f7ff;}"
    css = css & "tr.sal-row:hover{background-color:#dce9ff;cursor:pointer;}"
    css = css & ".tooltip{display:none;position:fixed;background:#fff;border:2px solid #1C6EA4;"
    css = css & "padding:10px;border-radius:6px;box-shadow:3px 3px 10px rgba(0,0,0,0.3);"
    css = css & "z-index:9999;min-width:450px;font-size:11px;pointer-events:none;}"
    css = css & ".tooltip table{width:100%;margin-bottom:6px;border-collapse:collapse;}"
    css = css & ".tooltip th{background:#555;color:white;padding:3px 6px;font-size:11px;border:1px solid #999;}"
    css = css & ".tooltip td{padding:3px 6px;border:1px solid #ddd;}"
    css = css & ".src-label{font-weight:bold;display:block;margin:6px 0 3px 0;font-size:12px;}"
    css = css & ".src-13e{color:#00B050;}"
    css = css & ".src-cat{color:#FF6600;}"
    css = css & ".diff{color:#F04D00;font-weight:bold;}"
    css = css & ".no-data{color:gray;font-style:italic;}"
    css = css & ".tip-title{font-size:13px;font-weight:bold;border-bottom:1px solid #ddd;padding-bottom:4px;margin-bottom:6px;}"
    
    Dim js As String
    js = "function showTip(e,id){"
    js = js & "var t=document.getElementById(id);"
    js = js & "if(!t)return;"
    js = js & "t.style.display='block';"
    js = js & "positionTip(e,t);}"
    js = js & "function hideTip(id){"
    js = js & "var t=document.getElementById(id);"
    js = js & "if(t)t.style.display='none';}"
    js = js & "function moveTip(e,id){"
    js = js & "var t=document.getElementById(id);"
    js = js & "if(t&&t.style.display=='block')positionTip(e,t);}"
    js = js & "function positionTip(e,t){"
    js = js & "var x=e.clientX+15,y=e.clientY+15;"
    js = js & "var w=t.offsetWidth,h=t.offsetHeight;"
    js = js & "if(x+w>window.innerWidth)x=e.clientX-w-5;"
    js = js & "if(y+h>window.innerHeight)y=e.clientY-h-5;"
    js = js & "t.style.left=x+'px';t.style.top=y+'px';}"
    
    html = "<html><head>"
    html = html & "<style>" & css & "</style>"
    html = html & "<script>" & js & "</script>"
    html = html & "</head><body>"
    
    If IsEmpty(resultsSAL) Then
        html = html & "<p style='color:gray;font-style:italic'>Aucun param tre SAL trouv .</p>"
        GenererHTMLSALAvecTooltip = html & "</body></html>"
        Exit Function
    End If
    
    ' Tableau SAL
    html = html & "<table>"
    html = html & "<tr><th>Num s&eacute;q.</th><th>Param&egrave;tre</th><th>Val Normale</th><th>Unit&eacute; mesure</th><th>Seuil TB</th><th>Seuil B</th><th>Seuil H</th><th>Seuil TH</th><th>Val REF</th><th>Unit&eacute; REF</th><th>Commentaire 1</th><th>Commentaire 2</th><th>Commentaire 3</th></tr>"
    
    Dim tooltips As String
    tooltips = ""
    
    Dim salIdx() As Integer
    salIdx = TrierIndicesParCol1(resultsSAL)
    Dim i As Integer
    For i = 1 To UBound(resultsSAL, 2)
        Dim realI As Integer: realI = salIdx(i)
        Dim nomParam As String
        nomParam = CStr(resultsSAL(2, realI))
        Dim tipId As String
        tipId = "tip" & i

        ' Ligne SAL
        Dim evts As String
        evts = " class='sal-row'"
        evts = evts & " onmouseover='showTip(event,""" & tipId & """)'"
        evts = evts & " onmouseout='hideTip(""" & tipId & """)'"
        evts = evts & " onmousemove='moveTip(event,""" & tipId & """)'"

        html = html & "<tr" & evts & ">"
        html = html & "<td>" & CStr(resultsSAL(1, realI)) & "</td>"
        html = html & "<td>" & nomParam & "</td>"
        html = html & "<td>" & CStr(resultsSAL(3, realI)) & "</td>"
        html = html & "<td>" & CStr(resultsSAL(4, realI)) & "</td>"
        html = html & "<td>" & CStr(resultsSAL(5, realI)) & "</td>"
        html = html & "<td>" & CStr(resultsSAL(6, realI)) & "</td>"
        html = html & "<td>" & CStr(resultsSAL(7, realI)) & "</td>"
        html = html & "<td>" & CStr(resultsSAL(8, realI)) & "</td>"
        html = html & "<td>" & CStr(resultsSAL(9, realI)) & "</td>"
        html = html & "<td>" & CStr(resultsSAL(10, realI)) & "</td>"
        html = html & "<td>" & CStr(resultsSAL(11, realI)) & "</td>"
        html = html & "<td>" & CStr(resultsSAL(12, realI)) & "</td>"
        html = html & "<td>" & CStr(resultsSAL(13, realI)) & "</td>"
        html = html & "</tr>"
        
        ' Construire tooltip (hors table)
        Dim tip As String
        tip = "<div id='" & tipId & "' class='tooltip'>"
        tip = tip & "<div class='tip-title'>" & nomParam & "</div>"
        
        ' 13E
        Dim idx13E As Integer
        idx13E = TrouverIndexParamLocal(UCase(Trim(nomParam)), results13E)
        tip = tip & "<span class='src-label src-13e'>&#128196; 13E</span>"
        If idx13E > 0 Then
            tip = tip & "<table><tr>"
            tip = tip & "<th>Seuil TB</th><th>Seuil B</th><th>Seuil H</th><th>Seuil TH</th><th>Val Normale</th>"
            tip = tip & "</tr><tr>"
            tip = tip & "<td class='" & DiffClass(CStr(resultsSAL(5, realI)), CStr(results13E(5, idx13E))) & "'>" & CStr(results13E(5, idx13E)) & "</td>"
            tip = tip & "<td class='" & DiffClass(CStr(resultsSAL(6, realI)), CStr(results13E(6, idx13E))) & "'>" & CStr(results13E(6, idx13E)) & "</td>"
            tip = tip & "<td class='" & DiffClass(CStr(resultsSAL(7, realI)), CStr(results13E(7, idx13E))) & "'>" & CStr(results13E(7, idx13E)) & "</td>"
            tip = tip & "<td class='" & DiffClass(CStr(resultsSAL(8, realI)), CStr(results13E(8, idx13E))) & "'>" & CStr(results13E(8, idx13E)) & "</td>"
            tip = tip & "<td class='" & DiffClass(CStr(resultsSAL(3, realI)), CStr(results13E(3, idx13E))) & "'>" & CStr(results13E(3, idx13E)) & "</td>"
            tip = tip & "</tr></table>"
        Else
            tip = tip & "<span class='no-data'>Non trouv  dans 13E</span>"
        End If
        
        ' CAT
        Dim idxCAT As Integer
        idxCAT = TrouverIndexParamLocal(UCase(Trim(nomParam)), resultsCAT)
        tip = tip & "<span class='src-label src-cat'>&#128196; CAT</span>"
        If idxCAT > 0 Then
            tip = tip & "<table><tr>"
            tip = tip & "<th>Seuil TB</th><th>Seuil B</th><th>Seuil H</th><th>Seuil TH</th><th>Val Normale</th>"
            tip = tip & "</tr><tr>"
            tip = tip & "<td class='" & DiffClass(CStr(resultsSAL(5, realI)), CStr(resultsCAT(5, idxCAT))) & "'>" & CStr(resultsCAT(5, idxCAT)) & "</td>"
            tip = tip & "<td class='" & DiffClass(CStr(resultsSAL(6, realI)), CStr(resultsCAT(6, idxCAT))) & "'>" & CStr(resultsCAT(6, idxCAT)) & "</td>"
            tip = tip & "<td class='" & DiffClass(CStr(resultsSAL(7, realI)), CStr(resultsCAT(7, idxCAT))) & "'>" & CStr(resultsCAT(7, idxCAT)) & "</td>"
            tip = tip & "<td class='" & DiffClass(CStr(resultsSAL(8, realI)), CStr(resultsCAT(8, idxCAT))) & "'>" & CStr(resultsCAT(8, idxCAT)) & "</td>"
            tip = tip & "<td class='" & DiffClass(CStr(resultsSAL(3, realI)), CStr(resultsCAT(3, idxCAT))) & "'>" & CStr(resultsCAT(3, idxCAT)) & "</td>"
            tip = tip & "</tr></table>"
        Else
            tip = tip & "<span class='no-data'>Non trouv  dans CAT</span>"
        End If
        
        tip = tip & "</div>"
        tooltips = tooltips & tip  ' Accumuler les tooltips
    Next i
    
    html = html & "</table>"
    html = html & "<p style='color:gray;font-size:11px'>" & UBound(resultsSAL, 2) & " param tre(s)   survolez une ligne pour comparer</p>"
    
    ' Ajouter TOUS les tooltips   la fin du body (hors table)
    html = html & tooltips
    html = html & "</body></html>"
    
    GenererHTMLSALAvecTooltip = html
End Function

' G N RATION HTML G N RIQUE

Private Function GenererHTMLTableau(resultats As Variant, source As String, couleurHeader As String, couleurPair As String, couleurHover As String, entetes As Variant) As String
    Dim html As String
    Dim css As String
    css = "body{font-family:Calibri,Arial;font-size:12px;margin:5px;}"
    css = css & "table{border-collapse:collapse;width:100%;}"
    css = css & "th{background-color:" & couleurHeader & ";color:white;padding:6px 8px;text-align:left;border:1px solid #ccc;}"
    css = css & "td{padding:5px 8px;border:1px solid #ddd;}"
    css = css & "tr:nth-child(even){background-color:" & couleurPair & ";}"
    css = css & "tr:hover{background-color:" & couleurHover & ";}"
    
    html = "<html><head><style>" & css & "</style></head><body>"
    
    If IsEmpty(resultats) Then
        html = html & "<p style='color:gray;font-style:italic'>Aucun param tre " & source & " trouv .</p>"
    Else
        html = html & "<table><tr>"
        Dim k As Integer
        For k = 0 To UBound(entetes)
            html = html & "<th>" & entetes(k) & "</th>"
        Next k
        html = html & "</tr>"
        
        Dim nbCols As Integer
        nbCols = UBound(resultats, 1)
        
        Dim tblIdx() As Integer
        tblIdx = TrierIndicesParCol1(resultats)
        Dim i As Integer
        For i = 1 To UBound(resultats, 2)
            Dim realIdx As Integer: realIdx = tblIdx(i)
            html = html & "<tr>"
            Dim c As Integer
            For c = 1 To nbCols
                html = html & "<td>" & CStr(resultats(c, realIdx)) & "</td>"
            Next c
            html = html & "</tr>"
        Next i
        
        html = html & "</table>"
        html = html & "<p style='color:gray;font-size:11px'>" & UBound(resultats, 2) & " param tre(s)</p>"
    End If
    
    html = html & "</body></html>"
    GenererHTMLTableau = html
End Function

' HELPERS

Private Function DiffClass(valSAL As String, valAutre As String) As String
    If valSAL <> "" And valAutre <> "" And valSAL <> valAutre Then
        DiffClass = "diff"
    Else
        DiffClass = ""
    End If
End Function

Private Function TrouverIndexParamLocal(nomParam As String, resultats As Variant) As Integer
    If IsEmpty(resultats) Then
        TrouverIndexParamLocal = 0
        Exit Function
    End If
    Dim i As Integer
    For i = 1 To UBound(resultats, 2)
        If UCase(Trim(CStr(resultats(2, i)))) = UCase(Trim(nomParam)) Then
            TrouverIndexParamLocal = i
            Exit Function
        End If
    Next i
    TrouverIndexParamLocal = 0
End Function

Private Function TrierIndicesParCol1(arr As Variant) As Integer()
    Dim n As Integer: n = UBound(arr, 2)
    Dim idx() As Integer
    ReDim idx(1 To n)
    Dim ii As Integer
    For ii = 1 To n
        idx(ii) = ii
    Next ii
    Dim jj As Integer, tmp As Integer
    For ii = 1 To n - 1
        For jj = 1 To n - ii
            Dim va As String: va = CStr(arr(1, idx(jj)))
            Dim vb As String: vb = CStr(arr(1, idx(jj + 1)))
            Dim doSwap As Boolean: doSwap = False
            If IsNumeric(va) And IsNumeric(vb) Then
                If CDbl(va) > CDbl(vb) Then doSwap = True
            Else
                If va > vb Then doSwap = True
            End If
            If doSwap Then
                tmp = idx(jj)
                idx(jj) = idx(jj + 1)
                idx(jj + 1) = tmp
            End If
        Next jj
    Next ii
    TrierIndicesParCol1 = idx
End Function

' BOUTONS

Public Sub SetSources(sal As Boolean, e13 As Boolean, cat As Boolean, sygma As Boolean)
    m_loadSAL = sal
    m_load13E = e13
    m_loadCAT = cat
    m_loadSYGMA = sygma
End Sub

Private Sub m_btnRecherche_Click()
    Dim frm As New FrmRecherche
    frm.SetCode m_CodeEqt
    frm.Show
    If Not frm.Confirmed Then Exit Sub
    Dim newCode As String
    newCode = Trim(frm.codeEqt)
    If newCode = "" Then Exit Sub
    If Not EquipementExiste(newCode, GetCheminFichier(newCode)) Then
        MsgBox Chr(201) & "quipement non trouv" & Chr(233) & " : " & newCode, vbInformation
        Exit Sub
    End If
    m_loadSAL = frm.loadSAL
    m_load13E = frm.load13E
    m_loadCAT = frm.loadCAT
    m_loadSYGMA = frm.loadSYGMA
    m_CodeEqt = newCode
    lblCode.Caption = Chr(201) & "quipement : " & m_CodeEqt
    mpOnglets.Value = 0
    Call ChargerDonnees
End Sub

Private Sub btnFermer_Click()
    Unload Me
End Sub

Private Sub btnExporter_Click()
    On Error GoTo ErrorHandler
    Dim resultats As Variant
    resultats = ChercherParametres(m_CodeEqt, GetCheminFichierParams(m_CodeEqt))
    Call AfficherResultats(m_CodeEqt, resultats)
    MsgBox "Export  dans la feuille 'R sultats'", vbInformation
    Exit Sub
ErrorHandler:
    Call LogError("btnExporter_Click", Err.Description, Err.Number)
    MsgBox "Erreur export. Consultez les Logs.", vbExclamation
End Sub




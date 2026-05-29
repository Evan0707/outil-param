VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Resultats 
   Caption         =   "UserForm1"
   ClientHeight    =   9440
   ClientLeft      =   110
   ClientTop       =   450
   ClientWidth     =   15780
   OleObjectBlob   =   "Resultats.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Resultats"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private m_CodeEqt As String
Private m_HTMLSAL As String
Private m_HTML13E As String
Private m_HTMLCAT As String
Private m_HTMLSYGMA As String

' INITIALISATION

Public Sub InitAvecCode(codeEqt As String)
    On Error GoTo ErrorHandler
    m_CodeEqt = codeEqt
    lblCode.Caption = "Équipement : " & m_CodeEqt
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
    Me.Width = 800
    Me.Height = 500
    Me.Caption = "Résultats équipement"
    mpOnglets.Pages(0).Caption = "SAL"
    mpOnglets.Pages(1).Caption = "13E"
    mpOnglets.Pages(2).Caption = "CAT"
    mpOnglets.Pages(3).Caption = "SYGMA"
    Exit Sub
ErrorHandler:
    Call LogError("UserForm_Initialize", Err.Description, Err.Number)
End Sub

' CHARGEMENT DONNÉES

Private Sub ChargerDonnees()
    On Error GoTo ErrorHandler
    Call LogInfo("ChargerDonnees", "Début pour : " & m_CodeEqt)
    
    ' Charger 13E et CAT d'abord (nécessaire pour les tooltips SAL)
    Dim results13E As Variant
    results13E = ChercherTousParametres(m_CodeEqt, GetCheminFichier13E(m_CodeEqt))
    Call FermerFichiersExternes
    
    Dim resultsCAT As Variant
    resultsCAT = ChercherTousParametres(m_CodeEqt, GetCheminFichierCAT(m_CodeEqt))
    Call FermerFichiersExternes
    
    ' SAL avec tooltips comparatifs
    Dim resultsSAL As Variant
    resultsSAL = ChercherParametres(m_CodeEqt, GetCheminFichierParams(m_CodeEqt))
    m_HTMLSAL = GenererHTMLSALAvecTooltip(resultsSAL, results13E, resultsCAT)
    Call FermerFichiersExternes
    
    ' 13E
    m_HTML13E = GenererHTMLTableau(results13E, "13E", "#00B050", "#f0fff0", "#d0ffd0", _
        Array("Paramètre", "Seuil B", "Seuil H", "Seuil TB", "Seuil TH", "Val Normale", "Unité"))
    
    ' CAT
    m_HTMLCAT = GenererHTMLTableau(resultsCAT, "CAT", "#FF6600", "#fff5f0", "#ffd0b0", _
        Array("Paramètre", "Seuil B", "Seuil H", "Seuil TB", "Seuil TH", "Val Normale", "Unité"))
    
    ' SYGMA
    Dim resultsSYGMA As Variant
    resultsSYGMA = ChercherParametresSYGMA(m_CodeEqt, GetCheminFichierSYGMA())
    m_HTMLSYGMA = GenererHTMLTableau(resultsSYGMA, "SYGMA", "#6600CC", "#f5f0ff", "#e0d0ff", _
        Array("Paramètre", "Val Basse", "Val Normale", "Val Haute", "Unité"))
    
    ' Afficher SAL par défaut
    Call AfficherHTML(m_HTMLSAL)
    Call LogInfo("ChargerDonnees", "Terminé")
    Exit Sub
ErrorHandler:
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

' GÉNÉRATION HTML SAL AVEC TOOLTIP

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
        html = html & "<p style='color:gray;font-style:italic'>Aucun paramètre SAL trouvé.</p>"
        GenererHTMLSALAvecTooltip = html & "</body></html>"
        Exit Function
    End If
    
    ' Tableau SAL
    html = html & "<table>"
    html = html & "<tr><th>Paramètre</th><th>Seuil B</th><th>Seuil H</th><th>Seuil TB</th><th>Seuil TH</th><th>Val Normale</th><th>Unité</th></tr>"
    
    Dim tooltips As String
    tooltips = ""
    
    Dim i As Integer
    For i = 1 To UBound(resultsSAL, 2)
        Dim nomParam As String
        nomParam = CStr(resultsSAL(1, i))
        Dim tipId As String
        tipId = "tip" & i
        
        ' Ligne SAL
        Dim evts As String
        evts = " class='sal-row'"
        evts = evts & " onmouseover='showTip(event,""" & tipId & """)'"
        evts = evts & " onmouseout='hideTip(""" & tipId & """)'"
        evts = evts & " onmousemove='moveTip(event,""" & tipId & """)'"
        
        html = html & "<tr" & evts & ">"
        html = html & "<td>" & nomParam & "</td>"
        html = html & "<td>" & CStr(resultsSAL(2, i)) & "</td>"
        html = html & "<td>" & CStr(resultsSAL(3, i)) & "</td>"
        html = html & "<td>" & CStr(resultsSAL(4, i)) & "</td>"
        html = html & "<td>" & CStr(resultsSAL(5, i)) & "</td>"
        html = html & "<td>" & CStr(resultsSAL(6, i)) & "</td>"
        html = html & "<td>" & CStr(resultsSAL(7, i)) & "</td>"
        html = html & "</tr>"
        
        ' Construire tooltip (hors table)
        Dim tip As String
        tip = "<div id='" & tipId & "' class='tooltip'>"
        tip = tip & "<div class='tip-title'>" & nomParam & "</div>"
        
        ' 13E
        Dim idx13E As Integer
        idx13E = TrouverIndexParamLocal(UCase(Trim(nomParam)), results13E)
        tip = tip & "<span class='src-label src-13e'>? 13E</span>"
        If idx13E > 0 Then
            tip = tip & "<table><tr>"
            tip = tip & "<th>Seuil B</th><th>Seuil H</th><th>Seuil TB</th><th>Seuil TH</th><th>Val Normale</th>"
            tip = tip & "</tr><tr>"
            tip = tip & "<td class='" & DiffClass(CStr(resultsSAL(2, i)), CStr(results13E(2, idx13E))) & "'>" & CStr(results13E(2, idx13E)) & "</td>"
            tip = tip & "<td class='" & DiffClass(CStr(resultsSAL(3, i)), CStr(results13E(3, idx13E))) & "'>" & CStr(results13E(3, idx13E)) & "</td>"
            tip = tip & "<td class='" & DiffClass(CStr(resultsSAL(4, i)), CStr(results13E(4, idx13E))) & "'>" & CStr(results13E(4, idx13E)) & "</td>"
            tip = tip & "<td class='" & DiffClass(CStr(resultsSAL(5, i)), CStr(results13E(5, idx13E))) & "'>" & CStr(results13E(5, idx13E)) & "</td>"
            tip = tip & "<td class='" & DiffClass(CStr(resultsSAL(6, i)), CStr(results13E(6, idx13E))) & "'>" & CStr(results13E(6, idx13E)) & "</td>"
            tip = tip & "</tr></table>"
        Else
            tip = tip & "<span class='no-data'>Non trouvé dans 13E</span>"
        End If
        
        ' CAT
        Dim idxCAT As Integer
        idxCAT = TrouverIndexParamLocal(UCase(Trim(nomParam)), resultsCAT)
        tip = tip & "<span class='src-label src-cat'>? CAT</span>"
        If idxCAT > 0 Then
            tip = tip & "<table><tr>"
            tip = tip & "<th>Seuil B</th><th>Seuil H</th><th>Seuil TB</th><th>Seuil TH</th><th>Val Normale</th>"
            tip = tip & "</tr><tr>"
            tip = tip & "<td class='" & DiffClass(CStr(resultsSAL(2, i)), CStr(resultsCAT(2, idxCAT))) & "'>" & CStr(resultsCAT(2, idxCAT)) & "</td>"
            tip = tip & "<td class='" & DiffClass(CStr(resultsSAL(3, i)), CStr(resultsCAT(3, idxCAT))) & "'>" & CStr(resultsCAT(3, idxCAT)) & "</td>"
            tip = tip & "<td class='" & DiffClass(CStr(resultsSAL(4, i)), CStr(resultsCAT(4, idxCAT))) & "'>" & CStr(resultsCAT(4, idxCAT)) & "</td>"
            tip = tip & "<td class='" & DiffClass(CStr(resultsSAL(5, i)), CStr(resultsCAT(5, idxCAT))) & "'>" & CStr(resultsCAT(5, idxCAT)) & "</td>"
            tip = tip & "<td class='" & DiffClass(CStr(resultsSAL(6, i)), CStr(resultsCAT(6, idxCAT))) & "'>" & CStr(resultsCAT(6, idxCAT)) & "</td>"
            tip = tip & "</tr></table>"
        Else
            tip = tip & "<span class='no-data'>Non trouvé dans CAT</span>"
        End If
        
        tip = tip & "</div>"
        tooltips = tooltips & tip  ' Accumuler les tooltips
    Next i
    
    html = html & "</table>"
    html = html & "<p style='color:gray;font-size:11px'>" & UBound(resultsSAL, 2) & " paramètre(s) — survolez une ligne pour comparer</p>"
    
    ' Ajouter TOUS les tooltips à la fin du body (hors table)
    html = html & tooltips
    html = html & "</body></html>"
    
    GenererHTMLSALAvecTooltip = html
End Function

' GÉNÉRATION HTML GÉNÉRIQUE

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
        html = html & "<p style='color:gray;font-style:italic'>Aucun paramètre " & source & " trouvé.</p>"
    Else
        html = html & "<table><tr>"
        Dim k As Integer
        For k = 0 To UBound(entetes)
            html = html & "<th>" & entetes(k) & "</th>"
        Next k
        html = html & "</tr>"
        
        Dim nbCols As Integer
        nbCols = UBound(resultats, 1)
        
        Dim i As Integer
        For i = 1 To UBound(resultats, 2)
            html = html & "<tr>"
            Dim c As Integer
            For c = 1 To nbCols
                html = html & "<td>" & CStr(resultats(c, i)) & "</td>"
            Next c
            html = html & "</tr>"
        Next i
        
        html = html & "</table>"
        html = html & "<p style='color:gray;font-size:11px'>" & UBound(resultats, 2) & " paramètre(s)</p>"
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
        If UCase(Trim(CStr(resultats(1, i)))) = UCase(Trim(nomParam)) Then
            TrouverIndexParamLocal = i
            Exit Function
        End If
    Next i
    TrouverIndexParamLocal = 0
End Function

' BOUTONS

Private Sub btnFermer_Click()
    Unload Me
End Sub

Private Sub btnExporter_Click()
    On Error GoTo ErrorHandler
    Dim resultats As Variant
    resultats = ChercherParametres(m_CodeEqt, GetCheminFichierParams(m_CodeEqt))
    Call AfficherResultats(m_CodeEqt, resultats)
    MsgBox "Exporté dans la feuille 'Résultats'", vbInformation
    Exit Sub
ErrorHandler:
    Call LogError("btnExporter_Click", Err.Description, Err.Number)
    MsgBox "Erreur export. Consultez les Logs.", vbExclamation
End Sub


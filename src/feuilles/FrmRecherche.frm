
Option Explicit

Public Confirmed As Boolean
Public codeEqt As String
Public loadSAL As Boolean
Public load13E As Boolean
Public loadCAT As Boolean
Public loadSYGMA As Boolean

Private WithEvents m_btnOK As MSForms.CommandButton
Private WithEvents m_btnCancel As MSForms.CommandButton
Private WithEvents m_btnFav As MSForms.CommandButton
Private WithEvents m_txtCode As MSForms.ComboBox
Private m_chkSAL As MSForms.CheckBox
Private m_chk13E As MSForms.CheckBox
Private m_chkCAT As MSForms.CheckBox
Private m_chkSYGMA As MSForms.CheckBox

Private Sub UserForm_Initialize()
    Me.Caption = "Recherche " & Chr(233) & "quipement"
    Me.Width = 300
    Me.Height = 205
    Me.BackColor = RGB(250, 250, 250)

    Dim lbl As MSForms.Label
    Set lbl = Me.Controls.Add("Forms.Label.1")
    With lbl
        .Caption = "Code " & Chr(233) & "quipement :"
        .Left = 10: .top = 12: .Width = 150: .Height = 15
        .Font.Size = 8
        .BackStyle = 0
        .ForeColor = RGB(60, 60, 60)
    End With

    Set m_txtCode = Me.Controls.Add("Forms.ComboBox.1", "txtCode")
    With m_txtCode
        .Left = 10: .top = 29: .Width = 232: .Height = 18
        .Font.Size = 10
        .SpecialEffect = 0
        .BackColor = RGB(255, 255, 255)
        .style = 0
    End With
    
    Set m_btnFav = Me.Controls.Add("Forms.CommandButton.1", "btnFav")
    With m_btnFav
        .Caption = ChrW(9734)
        .Left = 246: .top = 29: .Width = 26: .Height = 18
        .Font.Size = 11
        .BackColor = RGB(255, 255, 255)
        .BackStyle = 1
        .ForeColor = RGB(120, 120, 120)
        .ControlTipText = Chr(201) & "pingler / retirer des favoris"
    End With

    Dim lblSrc As MSForms.Label
    Set lblSrc = Me.Controls.Add("Forms.Label.1")
    With lblSrc
        .Caption = "Sources :"
        .Left = 10: .top = 57: .Width = 58: .Height = 15
        .Font.Size = 8
        .BackStyle = 0
        .ForeColor = RGB(60, 60, 60)
    End With

    Set m_chkSAL = Me.Controls.Add("Forms.CheckBox.1", "chkSAL")
    With m_chkSAL: .Caption = "SAL": .Value = True
        .Left = 70: .top = 55: .Width = 46: .Height = 18: .Font.Size = 8
        .SpecialEffect = 0: .BackStyle = 0: .BackColor = RGB(250, 250, 250)
        .ForeColor = RGB(60, 60, 60)
    End With
    Set m_chk13E = Me.Controls.Add("Forms.CheckBox.1", "chk13E")
    With m_chk13E: .Caption = "13E": .Value = True
        .Left = 118: .top = 55: .Width = 46: .Height = 18: .Font.Size = 8
        .SpecialEffect = 0: .BackStyle = 0: .BackColor = RGB(250, 250, 250)
        .ForeColor = RGB(60, 60, 60)
    End With
    Set m_chkCAT = Me.Controls.Add("Forms.CheckBox.1", "chkCAT")
    With m_chkCAT: .Caption = "CAT": .Value = True
        .Left = 166: .top = 55: .Width = 46: .Height = 18: .Font.Size = 8
        .SpecialEffect = 0: .BackStyle = 0: .BackColor = RGB(250, 250, 250)
        .ForeColor = RGB(60, 60, 60)
    End With
    Set m_chkSYGMA = Me.Controls.Add("Forms.CheckBox.1", "chkSYGMA")
    With m_chkSYGMA: .Caption = "SYGMA": .Value = True
        .Left = 214: .top = 55: .Width = 60: .Height = 18: .Font.Size = 8
        .SpecialEffect = 0: .BackStyle = 0: .BackColor = RGB(250, 250, 250)
        .ForeColor = RGB(60, 60, 60)
    End With

    Dim sep As MSForms.Label
    Set sep = Me.Controls.Add("Forms.Label.1")
    With sep
        .Caption = ""
        .Left = 10: .top = 77: .Width = 262: .Height = 1
        .BackColor = RGB(220, 220, 220)
        .BackStyle = 1
        .SpecialEffect = 0
    End With

    Set m_btnOK = Me.Controls.Add("Forms.CommandButton.1", "btnOK")
    With m_btnOK
        .Caption = "Rechercher"
        .Left = 55: .top = 86: .Width = 110: .Height = 26
        .Default = True: .Font.Size = 10
        .BackColor = RGB(255, 255, 255)
        .BackStyle = 1
        .ForeColor = RGB(40, 40, 40)
    End With

    Set m_btnCancel = Me.Controls.Add("Forms.CommandButton.1", "btnCancel")
    With m_btnCancel
        .Caption = "Annuler"
        .Left = 175: .top = 86: .Width = 90: .Height = 26
        .Cancel = True: .Font.Size = 10
        .BackColor = RGB(255, 255, 255)
        .BackStyle = 1
        .ForeColor = RGB(100, 100, 100)
    End With

    ' Remplit favoris + historique tout seul (aucun appel externe necessaire)
    Call RemplirListe
    Call MajEtoile
End Sub

' Remplit la liste deroulante : favoris (avec etoile) puis historique
Private Sub RemplirListe()
    On Error Resume Next
    m_txtCode.Clear
    Dim fav As Variant
    fav = LireFavoris()
    Dim i As Long
    If Not IsEmpty(fav) Then
        For i = 0 To UBound(fav)
            If Trim(CStr(fav(i))) <> "" Then m_txtCode.AddItem ChrW(9733) & " " & fav(i)
        Next i
    End If
    Dim hist As Variant
    hist = LireHistorique()
    If Not IsEmpty(hist) Then
        For i = 0 To UBound(hist)
            If Trim(CStr(hist(i))) <> "" And Not EstFavori(CStr(hist(i))) Then
                m_txtCode.AddItem hist(i)
            End If
        Next i
    End If
End Sub

' Retire l'etoile prefixe d'un libelle de la liste
Private Function NettoyerCode(s As String) As String
    s = Trim(s)
    If Len(s) > 0 Then
        If Left(s, 1) = ChrW(9733) Then s = Trim(Mid(s, 2))
    End If
    NettoyerCode = s
End Function

' Met a jour l'icone etoile selon que le code courant est favori ou non
Private Sub MajEtoile()
    On Error Resume Next
    If EstFavori(NettoyerCode(m_txtCode.Text)) Then
        m_btnFav.Caption = ChrW(9733)
        m_btnFav.ForeColor = RGB(230, 170, 0)
    Else
        m_btnFav.Caption = ChrW(9734)
        m_btnFav.ForeColor = RGB(120, 120, 120)
    End If
End Sub

Private Sub m_txtCode_Change()
    Call MajEtoile
End Sub
Private Sub m_btnFav_Click()
    Dim code As String
    code = NettoyerCode(m_txtCode.Text)
    If code = "" Then Exit Sub
    If EstFavori(code) Then
        Call RetirerFavori(code)
    Else
        Call AjouterFavori(code)
    End If
    Call RemplirListe
    m_txtCode.Text = code
    Call MajEtoile
End Sub
Public Sub SetCode(code As String)
    If Not m_txtCode Is Nothing Then m_txtCode.Text = code
End Sub

Private Sub m_btnOK_Click()
    codeEqt = NettoyerCode(m_txtCode.Text)
    loadSAL = m_chkSAL.Value
    load13E = m_chk13E.Value
    loadCAT = m_chkCAT.Value
    loadSYGMA = m_chkSYGMA.Value
    Confirmed = True
    Me.Hide
End Sub

Private Sub m_btnCancel_Click()
    Confirmed = False
    Me.Hide
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode = 0 Then
        Confirmed = False
        Cancel = 1
        Me.Hide
    End If
End Sub


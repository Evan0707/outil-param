
Option Explicit

Public Confirmed As Boolean
Public codeEqt As String
Public loadSAL As Boolean
Public load13E As Boolean
Public loadCAT As Boolean
Public loadSYGMA As Boolean

Private WithEvents m_btnOK As MSForms.CommandButton
Private WithEvents m_btnCancel As MSForms.CommandButton
Private m_txtCode As MSForms.TextBox
Private m_chkSAL As MSForms.CheckBox
Private m_chk13E As MSForms.CheckBox
Private m_chkCAT As MSForms.CheckBox
Private m_chkSYGMA As MSForms.CheckBox

Private Sub UserForm_Initialize()
    Me.Caption = "Recherche " & Chr(233) & "quipement"
    Me.Width = 300
    Me.Height = 195

    Dim lbl As MSForms.Label
    Set lbl = Me.Controls.Add("Forms.Label.1")
    With lbl
        .Caption = "Code " & Chr(233) & "quipement :"
        .Left = 10: .top = 12: .Width = 150: .Height = 18
        .Font.Size = 10
    End With

    Set m_txtCode = Me.Controls.Add("Forms.TextBox.1", "txtCode")
    With m_txtCode
        .Left = 10: .top = 32: .Width = 262: .Height = 24
        .Font.Size = 11
    End With

    Dim lblSrc As MSForms.Label
    Set lblSrc = Me.Controls.Add("Forms.Label.1")
    With lblSrc
        .Caption = "Sources :"
        .Left = 10: .top = 65: .Width = 60: .Height = 16
        .Font.Size = 9
    End With

    Set m_chkSAL = Me.Controls.Add("Forms.CheckBox.1", "chkSAL")
    With m_chkSAL: .Caption = "SAL": .Value = True
        .Left = 72: .top = 63: .Width = 48: .Height = 18: .Font.Size = 9
    End With
    Set m_chk13E = Me.Controls.Add("Forms.CheckBox.1", "chk13E")
    With m_chk13E: .Caption = "13E": .Value = True
        .Left = 124: .top = 63: .Width = 48: .Height = 18: .Font.Size = 9
    End With
    Set m_chkCAT = Me.Controls.Add("Forms.CheckBox.1", "chkCAT")
    With m_chkCAT: .Caption = "CAT": .Value = True
        .Left = 176: .top = 63: .Width = 48: .Height = 18: .Font.Size = 9
    End With
    Set m_chkSYGMA = Me.Controls.Add("Forms.CheckBox.1", "chkSYGMA")
    With m_chkSYGMA: .Caption = "SYGMA": .Value = True
        .Left = 228: .top = 63: .Width = 60: .Height = 18: .Font.Size = 9
    End With

    Set m_btnOK = Me.Controls.Add("Forms.CommandButton.1", "btnOK")
    With m_btnOK
        .Caption = "Rechercher"
        .Left = 55: .top = 92: .Width = 110: .Height = 26
        .Default = True: .Font.Size = 10
    End With

    Set m_btnCancel = Me.Controls.Add("Forms.CommandButton.1", "btnCancel")
    With m_btnCancel
        .Caption = "Annuler"
        .Left = 175: .top = 92: .Width = 90: .Height = 26
        .Cancel = True: .Font.Size = 10
    End With
End Sub

Public Sub SetCode(code As String)
    If Not m_txtCode Is Nothing Then m_txtCode.Text = code
End Sub

Private Sub m_btnOK_Click()
    codeEqt = Trim(m_txtCode.Text)
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


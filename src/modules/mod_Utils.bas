' ===================================================
' MODULE: mod_Utils
' Fonctions utilitaires (tranche, code, listes)
' ===================================================

Function SansTrancheCode(codeEqt As String) As String
    If Left(codeEqt, 1) = "." Or Left(codeEqt, 2) = "12" Then
        SansTrancheCode = Mid(codeEqt, 3)
    ElseIf Left(codeEqt, 1) Like "[0-9]" Then
        SansTrancheCode = Mid(codeEqt, 2)
    Else
        SansTrancheCode = codeEqt
    End If
End Function

Function GetTranche(codeEqt As String) As String
    If Left(codeEqt, 1) = "." Then
        GetTranche = "*"
    ElseIf Left(codeEqt, 2) = "12" Then
        GetTranche = "*"
    ElseIf Left(codeEqt, 1) Like "[0-9]" Then
        GetTranche = Left(codeEqt, 1)
    Else
        GetTranche = "*"
    End If
End Function

Function EstDansListe(valeur As String, liste() As String) As Boolean
    Dim i As Integer
    For i = LBound(liste) To UBound(liste)
        If liste(i) = valeur Then
            EstDansListe = True
            Exit Function
        End If
    Next i
    EstDansListe = False
End Function

Sub FermerFichiersExternes()
    Dim wbTemp As Workbook
    For Each wbTemp In Workbooks
        If wbTemp.Name <> ThisWorkbook.Name Then wbTemp.Close False
    Next wbTemp
End Sub

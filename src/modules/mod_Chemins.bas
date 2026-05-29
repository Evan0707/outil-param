Attribute VB_Name = "mod_Chemins"
' ===================================================
' MODULE: mod_Chemins
' Génération des chemins de fichiers selon la tranche
' Chemins RELATIFS basés sur ThisWorkbook.Path
' ===================================================

' Retourne le chemin absolu depuis un chemin relatif Config
Private Function ResoudreChemin(dossierConfig As String) As String
    Dim base As String
    base = ThisWorkbook.Path & "\"
    
    ' Si le chemin dans Config est déjà absolu, on le retourne tel quel
    If Left(dossierConfig, 1) = "\" Or Mid(dossierConfig, 2, 1) = ":" Then
        ResoudreChemin = dossierConfig
    Else
        ResoudreChemin = base & dossierConfig
    End If
    
    ' S'assurer que le chemin se termine par \
    If Right(ResoudreChemin, 1) <> "\" Then
        ResoudreChemin = ResoudreChemin & "\"
    End If
End Function

Function GetCheminFichier(codeEqt As String) As String
    Dim dossier As String: dossier = ResoudreChemin(GetConfig("DossierRecherche"))
    Dim tranche As String: tranche = GetTranche(codeEqt)
    Select Case tranche
        Case "0", "*": GetCheminFichier = dossier & GetConfig("FichierSAL_T0")
        Case "1": GetCheminFichier = dossier & GetConfig("FichierSAL_T1")
        Case "2": GetCheminFichier = dossier & GetConfig("FichierSAL_T2")
        Case "9": GetCheminFichier = dossier & GetConfig("FichierSAL_T9")
        Case Else: GetCheminFichier = ""
    End Select
End Function

Function GetCheminFichierParams(codeEqt As String) As String
    Dim dossier As String: dossier = ResoudreChemin(GetConfig("DossierParametres"))
    Dim tranche As String: tranche = GetTranche(codeEqt)
    Select Case tranche
        Case "0", "*": GetCheminFichierParams = dossier & GetConfig("FichierParams_T0")
        Case "1": GetCheminFichierParams = dossier & GetConfig("FichierParams_T1")
        Case "2": GetCheminFichierParams = dossier & GetConfig("FichierParams_T2")
        Case "9": GetCheminFichierParams = dossier & GetConfig("FichierParams_T9")
        Case Else: GetCheminFichierParams = ""
    End Select
End Function

Function GetCheminFichier13E(codeEqt As String) As String
    Dim dossier As String: dossier = ResoudreChemin(GetConfig("DossierParametres"))
    Dim tranche As String: tranche = GetTranche(codeEqt)
    Select Case tranche
        Case "0", "*": GetCheminFichier13E = dossier & GetConfig("Fichier13E_T0")
        Case "1": GetCheminFichier13E = dossier & GetConfig("Fichier13E_T1")
        Case "2": GetCheminFichier13E = dossier & GetConfig("Fichier13E_T2")
        Case "9": GetCheminFichier13E = dossier & GetConfig("Fichier13E_T9")
        Case Else: GetCheminFichier13E = ""
    End Select
End Function

Function GetCheminFichierCAT(codeEqt As String) As String
    Dim dossier As String: dossier = ResoudreChemin(GetConfig("DossierParametres"))
    Dim tranche As String: tranche = GetTranche(codeEqt)
    Select Case tranche
        Case "0", "*": GetCheminFichierCAT = dossier & GetConfig("FichierCAT_T0")
        Case "1": GetCheminFichierCAT = dossier & GetConfig("FichierCAT_T1")
        Case "2": GetCheminFichierCAT = dossier & GetConfig("FichierCAT_T2")
        Case "9": GetCheminFichierCAT = dossier & GetConfig("FichierCAT_T9")
        Case Else: GetCheminFichierCAT = ""
    End Select
End Function

Function GetCheminReleve() As String
    GetCheminReleve = ResoudreChemin(GetConfig("DossierReleves")) & GetConfig("FichierReleve")
End Function

Function GetCheminFichierSYGMA() As String
    GetCheminFichierSYGMA = ResoudreChemin(GetConfig("DossierParametres")) & GetConfig("FichierSYGMA")
End Function

Attribute VB_Name = "Module1"
Sub CreerBoutonsPrincipal()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(1) ' Ou remplacez par le nom de votre feuille principale
    
    ' Supprimer les anciens boutons
    Dim btn As Button
    For Each btn In ws.Buttons
        btn.Delete
    Next btn
    
    ' ===== BOUTON RECHERCHER =====
    Dim btnRecherche As Button
    Set btnRecherche = ws.Buttons.Add(50, 50, 200, 35)
    With btnRecherche
        .Caption = "Rechercher un equipement"
        .OnAction = "mod_Recherche.RechercherEquipement"
        .Font.Size = 11
        .Font.Bold = True
    End With
    ws.Shapes(btnRecherche.Name).Fill.ForeColor.RGB = RGB(140, 175, 255)
    ws.Shapes(btnRecherche.Name).TextFrame.Characters.Font.Color = RGB(0, 0, 0)
    
    ' ===== BOUTON VOIR LOGS =====
    Dim btnLogs As Button
    Set btnLogs = ws.Buttons.Add(50, 100, 200, 35)
    With btnLogs
        .Caption = "Voir les logs"
        .OnAction = "AfficherLogs"
        .Font.Size = 11
    End With
    
    ' ===== BOUTON VIDER LOGS =====
    Dim btnViderLogs As Button
    Set btnViderLogs = ws.Buttons.Add(270, 100, 200, 35)
    With btnViderLogs
        .Caption = "Vider les logs"
        .OnAction = "ViderLogs"
        .Font.Size = 11
    End With
    
    MsgBox "Boutons crees !", vbInformation
End Sub

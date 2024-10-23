Attribute VB_Name = "MIntranet"
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Rôle                    : MODULE DE GESTION DE L'INTRANET
' Nom                    : MIntranet.bas
' Date de création : 19/12/2002
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

'--- déclarations obligatoires ---
Option Explicit

'--- options générales ---
Option Base 1
DefVar A-Z

'--- constantes privées ---

'--- constantes publiques ---
Public NePasNaviguerMaintenant As Boolean       'indique qu'il ne faut pas naviguer à l'instant x
Public AdresseDeDepart As String                          'adresse de départ pour le navigateur Web

'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Rôle      : Lancement du navigateur avec le site intranet de l'entreprise
' Entrées :
' Retours :
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Sub LancementIntranet()
    
    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next

    '--- déclaration ---

    '--- affectation sur l'adresse de site de l'entreprise ---
    AdresseDeDepart = "http:\\didier"
            
    '--- lancement de la navigation ---
    With OccFSynoptique
        If Len(AdresseDeDepart) > 0 Then
            
            .CBAdresses.Text = AdresseDeDepart
            .CBAdresses.AddItem .CBAdresses.Text
    
            '--- essayer de naviguer jusqu'à l'adresse de départ ---
            .TimerNavigateurWeb.Enabled = True
            .WBNavigateurWeb.Navigate AdresseDeDepart
        
        End If
    End With

End Sub

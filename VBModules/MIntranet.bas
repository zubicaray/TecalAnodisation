Attribute VB_Name = "MIntranet"
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' R�le                    : MODULE DE GESTION DE L'INTRANET
' Nom                    : MIntranet.bas
' Date de cr�ation : 19/12/2002
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

'--- d�clarations obligatoires ---
Option Explicit

'--- options g�n�rales ---
Option Base 1
DefVar A-Z

'--- constantes priv�es ---

'--- constantes publiques ---
Public NePasNaviguerMaintenant As Boolean       'indique qu'il ne faut pas naviguer � l'instant x
Public AdresseDeDepart As String                          'adresse de d�part pour le navigateur Web

'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' R�le      : Lancement du navigateur avec le site intranet de l'entreprise
' Entr�es :
' Retours :
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Sub LancementIntranet()
    
    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next

    '--- d�claration ---

    '--- affectation sur l'adresse de site de l'entreprise ---
    AdresseDeDepart = "http:\\didier"
            
    '--- lancement de la navigation ---
    With OccFSynoptique
        If Len(AdresseDeDepart) > 0 Then
            
            .CBAdresses.Text = AdresseDeDepart
            .CBAdresses.AddItem .CBAdresses.Text
    
            '--- essayer de naviguer jusqu'� l'adresse de d�part ---
            .TimerNavigateurWeb.Enabled = True
            .WBNavigateurWeb.Navigate AdresseDeDepart
        
        End If
    End With

End Sub

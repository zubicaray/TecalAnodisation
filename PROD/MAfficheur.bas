Attribute VB_Name = "MAfficheur"
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Rôle                    : MODULE GERANT L'AFFICHEUR A LEDS ROUGE
' Nom                    : MAfficheur.bas
' Date de création : 05/09/2011
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

'--- déclarations obligatoires ---
Option Explicit

'--- options générales ---
Option Base 1
DefVar A-Z

'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Rôle      : Acquitte les alarmes (coupe le gyrophare et le klaxon)
' Entrées : NomVariableInterne -> Nom de la variable interne exemple "A"
'                         TexteAAfficher -> Le texte à afficher
' Retours :
' Détails  :
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function MessageAfficheur(ByVal NomVariableInterne As String, _
                                                        ByVal TexteAAfficher As String) As Double
    
    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next

    '--- déclaration ---
    Dim ValeurRetour As Double

    '--- affectation de la variable ---
    ValeurRetour = OccFPrincipale.RTUpdateManager1.UpdateTextVariable(NomVariableInterne, TexteAAfficher, "COLOR_DEFAULT", "STYLE_DEFAULT")
    
    '--- affichage de la variable concernée ---
    MessageAfficheur = OccFPrincipale.RTUpdateManager1.SetRunSequence(NomVariableInterne)

End Function


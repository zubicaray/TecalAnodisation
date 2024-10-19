Attribute VB_Name = "MAfficheur"
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' R�le                    : MODULE GERANT L'AFFICHEUR A LEDS ROUGE
' Nom                    : MAfficheur.bas
' Date de cr�ation : 05/09/2011
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

'--- d�clarations obligatoires ---
Option Explicit

'--- options g�n�rales ---
Option Base 1
DefVar A-Z

'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' R�le      : Acquitte les alarmes (coupe le gyrophare et le klaxon)
' Entr�es : NomVariableInterne -> Nom de la variable interne exemple "A"
'                         TexteAAfficher -> Le texte � afficher
' Retours :
' D�tails  :
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function MessageAfficheur(ByVal NomVariableInterne As String, _
                                                        ByVal TexteAAfficher As String) As Double
    
    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next

    '--- d�claration ---
    Dim ValeurRetour As Double

    '--- affectation de la variable ---
    ValeurRetour = OccFPrincipale.RTUpdateManager1.UpdateTextVariable(NomVariableInterne, TexteAAfficher, "COLOR_DEFAULT", "STYLE_DEFAULT")
    
    '--- affichage de la variable concern�e ---
    MessageAfficheur = OccFPrincipale.RTUpdateManager1.SetRunSequence(NomVariableInterne)

End Function


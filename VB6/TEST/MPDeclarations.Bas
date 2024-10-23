Attribute VB_Name = "MPDeclarationsAPI"
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Rôle                    : MODULE DES DECLARATIONS PUBLIQUES
' Nom                    : MPDeclarations.bas
' Date de création : 26/03/1999
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

'--- déclarations obligatoires ---
Option Explicit

'--- options générales ---
Option Base 1
DefVar A-Z

'--- types publiques ---

'*** PROCEDURES ***

'--- pause en millièmes de secondes ---
Public Declare Sub Sleep Lib "kernel32" (ByVal dwmilliseconds As Long)

'*** FONCTIONS ***

'---- fonctions Win32 pour extraire le nom de l'ordinateur et de l'utilisateur ---
Declare Function GetComputerName Lib "kernel32" Alias "GetComputerNameA" (ByVal lpBuffer As String, nSize As Long) As Long
Declare Function GetUserName Lib "advapi32.dll" Alias "GetUserNameA" (ByVal lpBuffer As String, nSize As Long) As Long

'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Rôle      : Appel de la routine API pour extraire le nom de l'ordinateur
' Entrées :
' Retours : API_NomOrdinateur -> Le nom de l'ordinateur sinon ""
' Détails  :
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function API_NomOrdinateur() As String

    '--- aiguillage en cas d'erreur ---
    On Error Resume Next

    '--- constantes privées ---
    Const LONGUEUR_MAXIMALE_NOM_ORDINATEUR As Long = 16
    
    '--- déclaration ---
    Dim EmplacementvbNullChar As Long
    Dim NomOrdinateur As String * LONGUEUR_MAXIMALE_NOM_ORDINATEUR
    
    '--- affectation ---
    NomOrdinateur = String(LONGUEUR_MAXIMALE_NOM_ORDINATEUR, vbNullChar)

    '--- appel de la routine ---
    If GetComputerName(NomOrdinateur, LONGUEUR_MAXIMALE_NOM_ORDINATEUR) <> 0 Then
        EmplacementvbNullChar = InStr(NomOrdinateur, vbNullChar)
        If EmplacementvbNullChar > 1 Then
            API_NomOrdinateur = Left(NomOrdinateur, EmplacementvbNullChar - 1)
        Else
            API_NomOrdinateur = ""
        End If
    Else
        API_NomOrdinateur = ""
    End If
    
End Function

'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Rôle      : Appel de la routine API pour extraire le nom de l'utilisateur de l'ordinateur
' Entrées :
' Retours : API_NomUtilisateur -> Le nom de l'utilisateur de l'ordinateur sinon ""
' Détails  :
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function API_NomUtilisateur() As String

    '--- aiguillage en cas d'erreur ---
    On Error Resume Next

    '--- constantes privées ---
    Const LONGUEUR_MAXIMALE_NOM_UTILISATEUR As Long = 16
    
    '--- déclaration ---
    Dim EmplacementvbNullChar As Long
    Dim NomUtilisateur As String * LONGUEUR_MAXIMALE_NOM_UTILISATEUR
    
    '--- affectation ---
    NomUtilisateur = String(LONGUEUR_MAXIMALE_NOM_UTILISATEUR, vbNullChar)

    '--- appel de la routine ---
    If GetUserName(NomUtilisateur, LONGUEUR_MAXIMALE_NOM_UTILISATEUR) <> 0 Then
        EmplacementvbNullChar = InStr(NomUtilisateur, vbNullChar)
        If EmplacementvbNullChar > 1 Then
            API_NomUtilisateur = Left(NomUtilisateur, EmplacementvbNullChar - 1)
        Else
            API_NomUtilisateur = ""
        End If
    Else
        API_NomUtilisateur = ""
    End If
    
End Function









Attribute VB_Name = "MConstantesInternes"
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Rôle                    : MODULE DES CONSTANTES INTERNES A TOUS LES PROGRAMMES
' Nom                    : MConstantesInternes.bas
' Date de création : 26/03/1999
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

'--- déclarations obligatoires ---
Option Explicit

'--- options générales ---
Option Base 1
DefVar A-Z

'*** CONSTANTES NUMERIQUES ***

'--- couleurs ---
Public Enum COULEURS
    
    BLANC = &HFFFFFF
    NOIR = &H0&
    
    GRIS_SYSTEME = &H8000000F
    GRIS_1 = &HE0E0E0
    GRIS_2 = &HC0C0C0
    GRIS_3 = &H808080
    GRIS_4 = &H404040

    ROUGE_0 = &HD5D5FF
    ROUGE_1 = &HC0C0FF
    ROUGE_2 = &H8080FF
    ROUGE_3 = &HFF&
    ROUGE_4 = &HC0&
    ROUGE_5 = &H80&
    ROUGE_6 = &H40&

    ORANGE_0 = &HD5EAFF
    ORANGE_1 = &HC0E0FF
    ORANGE_2 = &H80C0FF
    ORANGE_3 = &H80FF&
    ORANGE_4 = &H40C0&
    ORANGE_5 = &H4080&
    ORANGE_6 = &H404080

    JAUNE_0 = &HD5FFFF
    JAUNE_1 = &HC0FFFF
    JAUNE_2 = &H80FFFF
    JAUNE_3 = &HFFFF&
    JAUNE_4 = &HC0C0&
    JAUNE_5 = &H8080&
    JAUNE_6 = &H4040&

    VERT_0 = &HD5FFD5
    VERT_1 = &HC0FFC0
    VERT_2 = &H80FF80
    VERT_3 = &HFF00&
    VERT_4 = &HC000&
    VERT_5 = &H8000&
    VERT_6 = &H4000&

    CYAN_0 = &HFFFFD5
    CYAN_1 = &HFFFFC0
    CYAN_2 = &HFFFF80
    CYAN_3 = &HFFFF00
    CYAN_4 = &HC0C000
    CYAN_5 = &H808000
    CYAN_6 = &H404000

    BLEU_0 = &HFFD5D5
    BLEU_1 = &HFFC0C0
    BLEU_2 = &HFF8080
    BLEU_3 = &HFF0000
    BLEU_4 = &HC00000
    BLEU_5 = &H800000
    BLEU_6 = &H400000

    MAGENTA_0 = &HFFD5FF
    MAGENTA_1 = &HFFC0FF
    MAGENTA_2 = &HFF80FF
    MAGENTA_3 = &HFF00FF
    MAGENTA_4 = &HC000C0
    MAGENTA_5 = &H800080
    MAGENTA_6 = &H400040

End Enum

Public Enum MOIS_ANNEE
    JANVIER = 1
    FEVRIER = 2
    MARS = 3
    AVRIL = 4
    MAI = 5
    JUIN = 6
    JUILLET = 7
    AOUT = 8
    SEPTEMBRE = 9
    OCTOBRE = 10
    NOVEMBRE = 11
    DECEMBRE = 12
End Enum

Public Enum JOURS_SEMAINE
    LUNDI = 1
    MARDI = 2
    MERCREDI = 3
    JEUDI = 4
    VENDREDI = 5
    SAMEDI = 6
    DIMANCHE = 7
End Enum

'--- codes ASCII particuliers ---
Public Const CODE_ASCII_DOLLAR As Integer = 36
Public Const CODE_ASCII_PHI As Integer = 216
Public Const CODE_ASCII_EURO As Integer = 128

'*** CONSTANTES CHAINES ***
    
'--- épaisseur moyenne d'un caractère (en twips) de la police MS sans Serif 8 gras ---
Public Const EPAISSEUR_CARACTERE As Integer = 140

'--- espaces ---
Public Const UN_ESPACE As String * 1 = " "


Attribute VB_Name = "MMathematiques"
'---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Rôle                    : MODULE DES FORMULES MATHEMATIQUES ET DE CONVERSIONS
' Nom                    : MMathematiques.bas
' Date de création : 26/03/1999
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

'--- déclarations obligatoires ---
Option Explicit

'--- options générales ---
Option Base 1
DefVar A-Z

'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Rôle      : Calcul le factoriel d'un nombre
' Retours : Le factoriel du nombre
' Détails  : C'est le produit des n premiers nombres entiers naturels hormis le zéro
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function Fact(Nombre) As Double
    
    '--- aiguillage en cas d'erreur ---
    On Error Resume Next

    '--- déclaration ---
    Dim a

    '--- affectation ---
    Fact = 1

    '--- calcul ---
    For a = Nombre To 1 Step -1
        Fact = Fact * a
    Next a

End Function

'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Rôle      : Conversion d'un temps en un nombre décimale
' Entrées : TempsEnHeuresMinutes -> Temps en heures, minutes à convertir (format Date)
' Retours :            CDateEnDecimale -> Nombre décimale retournée en heures, minutes
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function CDateEnDecimale(ByVal TempsEnHeuresMinutes As Date) As Double
    
    '--- aiguillage en cas d'erreur ---
    On Error GoTo GestionErreurs
   
    '--- constantes privées ---
    Const TITRE_MESSAGES As String = "Conversion d'un format DATE en nombre décimale"
        
    '--- calcul de la valeur décimale ---
    CDateEnDecimale = Hour(TempsEnHeuresMinutes) + Minute(TempsEnHeuresMinutes) / 60
    
    Exit Function

'--- gestion des erreurs ---
GestionErreurs:
      
    '--- affichage du message d'erreur ---
    MessageErreur TITRE_MESSAGES, Err.Description, Err.Number

End Function

'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Rôle      : Conversion d'un temps sous forme de chaine (99 heures 59 minutes possible) en un nombre décimale
' Entrées : TempsEnHeuresMinutes  -> Temps en heures, minutes à convertir (format String)
'                 ValiditeReponse               -> FALSE = Réponse non valide
'                                                               TRUE  = Réponse valide
' Retours : CChaineDateEnDecimale -> Nombre décimale retournée en heures, minutes
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function CChaineTempsEnDecimale(ByVal TempsEnHeuresMinutes As String, _
                                                                       ByRef ValiditeReponse As Boolean) As Double
    
    '--- aiguillage en cas d'erreur ---
    On Error GoTo GestionErreurs
   
    '--- constantes privées ---
    Const TITRE_MESSAGES As String = "Conversion d'une chaine au format temps en nombre décimale"
        
    '--- déclaration ---
    Dim LesHeures As Double, _
           LesMinutes As Double
    
    '--- affectation ---
    ValiditeReponse = False
    CChaineTempsEnDecimale = 0
    
    '--- calcul de la valeur décimale ---
    If Len(TempsEnHeuresMinutes) = 5 Then
        If IsNumeric(Left(TempsEnHeuresMinutes, 2)) = True And IsNumeric(Right(TempsEnHeuresMinutes, 2)) = True Then
            
            '--- affectation ---
            LesHeures = CDbl(Left(TempsEnHeuresMinutes, 2))
            LesMinutes = CDbl(Right(TempsEnHeuresMinutes, 2))
            
            '--- calcul ---
            If LesMinutes >= 0 And LesMinutes <= 59 Then
                CChaineTempsEnDecimale = LesHeures + LesMinutes / 60
                ValiditeReponse = True
            End If
    
        End If
    End If
    
    Exit Function

'--- gestion des erreurs ---
GestionErreurs:
      
    '--- affichage du message d'erreur ---
    MessageErreur TITRE_MESSAGES, Err.Description, Err.Number

End Function

'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Rôle      : Conversion d'un nombre décimale au format DATE
' Entrées : TempsEnDecimale -> Temps en décimale
' Retours : CDecimaleEnDate  -> Date retournée heures, minutes (format DATE)
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function CDecimaleEnDate(ByVal TempsEnDecimale As Double) As Date
    
    
    '--- aiguillage en cas d'erreur ---
    On Error GoTo GestionErreurs
   
    '--- constantes privées ---
    Const TITRE_MESSAGES As String = "Conversion d'un nombre décimale au format DATE"
        
    '--- calcul de la date ---
    If TempsEnDecimale >= 0 And TempsEnDecimale <= 23.99 Then
        CDecimaleEnDate = CDate(CStr(Int(TempsEnDecimale)) & ":" & _
                                                    CStr(Round((TempsEnDecimale - Int(TempsEnDecimale)) * 60)) & _
                                                    ":00")
    End If
    
    Exit Function

'--- gestion des erreurs ---
GestionErreurs:
      
    '--- affichage du message d'erreur ---
    MessageErreur TITRE_MESSAGES, Err.Description, Err.Number

End Function

'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Rôle      : Conversion d'un nombre décimale en chaine de caractères
' Entrées : TempsEnDecimale    -> Temps en décimale
' Retours : CDecimaleEnChaine -> Chaine retournée heures, minutes "hh:mm" (format STRING)
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function CDecimaleEnChaine(ByVal TempsEnDecimale As Double) As String
    
    '--- aiguillage en cas d'erreur ---
    On Error GoTo GestionErreurs
   
    '--- constantes privées ---
    Const TITRE_MESSAGES As String = "Conversion d'un nombre décimale en chaine"
        
    '--- calcul de la date ---
    CDecimaleEnChaine = Right("00" & CStr(Int(TempsEnDecimale)), 2) & ":" & _
                                         Right("00" & CStr(Round((TempsEnDecimale - Int(TempsEnDecimale)) * 60)), 2)
    
    Exit Function

'--- gestion des erreurs ---
GestionErreurs:
      
    '--- affichage du message d'erreur ---
    MessageErreur TITRE_MESSAGES, Err.Description, Err.Number

End Function

'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Rôle      : Conversion d'un temps en décimale
' Entrées : TempsEnHeuresMinutes -> Temps en heures, minutes à convertir (format Date)
'                              TypeConversion -> FALSE = heure admise de 00:00 à 23:59
'                                                              TRUE  = heure admise de 00:00 à 99:59
' Retours :              CHeureDecimale -> Heures, minutes en décimale
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function CHeureDecimale(ByVal TempsEnHeuresMinutes As String, _
                                                      ByRef TypeConversion As Boolean) As Double
    
    '--- aiguillage en cas d'erreur ---
    On Error Resume Next

    '--- déclaration ---
    Dim ReponseValide As Boolean
    Dim TempsDecimale As Double
    
    If TypeConversion = False Then
    
        '--- contrôle des valeurs et affichage en décimale (maxi 23 heures 59 minutes) ---
        If IsDate(TempsEnHeuresMinutes) = True Then
            CHeureDecimale = CDbl(FormatNumber(CDateEnDecimale(CDate(TempsEnHeuresMinutes)), 2))
        Else
            CHeureDecimale = 0
        End If
    
    Else
        
        '--- contrôle des valeurs et affichage en décimale (maxi 99 heures 59 minutes) ---
        TempsDecimale = CChaineTempsEnDecimale(TempsEnHeuresMinutes, ReponseValide)
        If ReponseValide = True Then
            CHeureDecimale = CDbl(FormatNumber(TempsDecimale, 2))
        Else
            CHeureDecimale = 0
        End If
    
    End If
    
End Function


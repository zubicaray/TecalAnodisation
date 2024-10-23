Attribute VB_Name = "MMathematiques"
'---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' R�le                    : MODULE DES FORMULES MATHEMATIQUES ET DE CONVERSIONS
' Nom                    : MMathematiques.bas
' Date de cr�ation : 26/03/1999
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

'--- d�clarations obligatoires ---
Option Explicit

'--- options g�n�rales ---
Option Base 1
DefVar A-Z

'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' R�le      : Calcul le factoriel d'un nombre
' Retours : Le factoriel du nombre
' D�tails  : C'est le produit des n premiers nombres entiers naturels hormis le z�ro
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function Fact(Nombre) As Double
    
    '--- aiguillage en cas d'erreur ---
    On Error Resume Next

    '--- d�claration ---
    Dim a

    '--- affectation ---
    Fact = 1

    '--- calcul ---
    For a = Nombre To 1 Step -1
        Fact = Fact * a
    Next a

End Function

'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' R�le      : Conversion d'un temps en un nombre d�cimale
' Entr�es : TempsEnHeuresMinutes -> Temps en heures, minutes � convertir (format Date)
' Retours :            CDateEnDecimale -> Nombre d�cimale retourn�e en heures, minutes
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function CDateEnDecimale(ByVal TempsEnHeuresMinutes As Date) As Double
    
    '--- aiguillage en cas d'erreur ---
    On Error GoTo GestionErreurs
   
    '--- constantes priv�es ---
    Const TITRE_MESSAGES As String = "Conversion d'un format DATE en nombre d�cimale"
        
    '--- calcul de la valeur d�cimale ---
    CDateEnDecimale = Hour(TempsEnHeuresMinutes) + Minute(TempsEnHeuresMinutes) / 60
    
    Exit Function

'--- gestion des erreurs ---
GestionErreurs:
      
    '--- affichage du message d'erreur ---
    MessageErreur TITRE_MESSAGES, Err.Description, Err.Number

End Function

'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' R�le      : Conversion d'un temps sous forme de chaine (99 heures 59 minutes possible) en un nombre d�cimale
' Entr�es : TempsEnHeuresMinutes  -> Temps en heures, minutes � convertir (format String)
'                 ValiditeReponse               -> FALSE = R�ponse non valide
'                                                               TRUE  = R�ponse valide
' Retours : CChaineDateEnDecimale -> Nombre d�cimale retourn�e en heures, minutes
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function CChaineTempsEnDecimale(ByVal TempsEnHeuresMinutes As String, _
                                                                       ByRef ValiditeReponse As Boolean) As Double
    
    '--- aiguillage en cas d'erreur ---
    On Error GoTo GestionErreurs
   
    '--- constantes priv�es ---
    Const TITRE_MESSAGES As String = "Conversion d'une chaine au format temps en nombre d�cimale"
        
    '--- d�claration ---
    Dim LesHeures As Double, _
           LesMinutes As Double
    
    '--- affectation ---
    ValiditeReponse = False
    CChaineTempsEnDecimale = 0
    
    '--- calcul de la valeur d�cimale ---
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
' R�le      : Conversion d'un nombre d�cimale au format DATE
' Entr�es : TempsEnDecimale -> Temps en d�cimale
' Retours : CDecimaleEnDate  -> Date retourn�e heures, minutes (format DATE)
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function CDecimaleEnDate(ByVal TempsEnDecimale As Double) As Date
    
    
    '--- aiguillage en cas d'erreur ---
    On Error GoTo GestionErreurs
   
    '--- constantes priv�es ---
    Const TITRE_MESSAGES As String = "Conversion d'un nombre d�cimale au format DATE"
        
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
' R�le      : Conversion d'un nombre d�cimale en chaine de caract�res
' Entr�es : TempsEnDecimale    -> Temps en d�cimale
' Retours : CDecimaleEnChaine -> Chaine retourn�e heures, minutes "hh:mm" (format STRING)
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function CDecimaleEnChaine(ByVal TempsEnDecimale As Double) As String
    
    '--- aiguillage en cas d'erreur ---
    On Error GoTo GestionErreurs
   
    '--- constantes priv�es ---
    Const TITRE_MESSAGES As String = "Conversion d'un nombre d�cimale en chaine"
        
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
' R�le      : Conversion d'un temps en d�cimale
' Entr�es : TempsEnHeuresMinutes -> Temps en heures, minutes � convertir (format Date)
'                              TypeConversion -> FALSE = heure admise de 00:00 � 23:59
'                                                              TRUE  = heure admise de 00:00 � 99:59
' Retours :              CHeureDecimale -> Heures, minutes en d�cimale
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function CHeureDecimale(ByVal TempsEnHeuresMinutes As String, _
                                                      ByRef TypeConversion As Boolean) As Double
    
    '--- aiguillage en cas d'erreur ---
    On Error Resume Next

    '--- d�claration ---
    Dim ReponseValide As Boolean
    Dim TempsDecimale As Double
    
    If TypeConversion = False Then
    
        '--- contr�le des valeurs et affichage en d�cimale (maxi 23 heures 59 minutes) ---
        If IsDate(TempsEnHeuresMinutes) = True Then
            CHeureDecimale = CDbl(FormatNumber(CDateEnDecimale(CDate(TempsEnHeuresMinutes)), 2))
        Else
            CHeureDecimale = 0
        End If
    
    Else
        
        '--- contr�le des valeurs et affichage en d�cimale (maxi 99 heures 59 minutes) ---
        TempsDecimale = CChaineTempsEnDecimale(TempsEnHeuresMinutes, ReponseValide)
        If ReponseValide = True Then
            CHeureDecimale = CDbl(FormatNumber(TempsDecimale, 2))
        Else
            CHeureDecimale = 0
        End If
    
    End If
    
End Function


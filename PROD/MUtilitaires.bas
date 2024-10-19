Attribute VB_Name = "MUtilitaires"
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' R�le                    : MODULE DES UTILITAIRES
' Nom                    : MUtilitaires.bas
' Date de cr�ation : 26/03/1999
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

'--- d�clarations obligatoires ---
Option Explicit

'--- options g�n�rales ---
Option Base 1
DefVar A-Z



Public Sub LogCharges(ByVal msg As String)

        'SZ 2023
        Exit Sub
        Dim nUnit As Integer
        nUnit = FreeFile
        ' This assumes write access to the directory containing the program '
        ' You will need to choose another directory if this is not possible '
        Dim str As String
        
       
        str = Format(Now, "yyyymmdd")
        Open App.Path & "\" & str & "_DETAILS_CHARGES.txt" For Append As nUnit
        ' For Append As nUnit
        Print #nUnit, "  " & msg
        Close nUnit
 
     

End Sub
Public Sub Log(ByVal msg As String, Optional toPrint As Boolean = True)

    
    If toPrint = True And ShowLog Then
     
        Dim nUnit As Integer
        nUnit = FreeFile
        ' This assumes write access to the directory containing the program '
        ' You will need to choose another directory if this is not possible '
        Dim str As String
        str = Format(Now, "yyyymmdd")
        Open App.Path & "\" & str & "_LOG.txt" For Append As nUnit
        ' For Append As nUnit
        Print #nUnit, Format$(Now)
        Print #nUnit, "  " & msg
        Print #nUnit, " --------------------------------------- " '& Format$(Now)
        Close nUnit
    End If
     

End Sub

Public Sub LogPourCPO(ByVal msg As String)

     
        Dim nUnit As Integer
        nUnit = FreeFile
        ' This assumes write access to the directory containing the program '
        ' You will need to choose another directory if this is not possible '
        Dim str As String
        str = Format(Now, "yyyymmdd")
        Open App.Path & "\" & "CPO_LOG_" & str & ".txt" For Append As nUnit
        ' For Append As nUnit
        Print #nUnit, Format$(Now)
        Print #nUnit, "  " & msg
        Print #nUnit, " --------------------------------------- " '& Format$(Now)
        Close nUnit
     

End Sub

Public Function FolderExists(sFullPath As String) As Boolean
    Dim myFSO As Object
    Set myFSO = CreateObject("Scripting.FileSystemObject")
    FolderExists = myFSO.FolderExists(sFullPath)
    
End Function


'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' R�le      : Permet la modification par le code d'un contr�le li� (CORRIGE UN BUG ENORME DE VB)
' Entr�es : LeControleAModifier -> Repr�sente le contr�le � mofifier
'                Enregistrement         -> Enregistrement de la table
'                ValeurDuChamp        -> Valeur du champ
' Retours : LeControleAModifier -> Repr�sente le contr�le mofifi� (passage par r�f�rence)
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Sub ModifieControleLie(ByRef LeControleAModifier As Control, _
                                                  ByVal Enregistrement As ADODB.Recordset, _
                                                  ByVal ValeurDuChamp As Variant)

    
    'ATTENTION cette fonction peut �tre remplacer par exemple par :
    'ADODCFournisseurs.Recordset(TBAdrPrincipaleCodePostal.DataField).Value = "F-"
    'ou
    'ADODCFournisseurs.Recordset("AdrPrincipaleCodePostal").Value = "F-"
    
    '--- aiguillage en cas d'erreur ---
    On Error Resume Next
    
    '--- d�claration ---
    Dim NomChampDonneeLie As Variant

    '--- modification ---
    With LeControleAModifier
        'If Not (Enregistrement.BOF And Enregistrement.EOF) Then
            Enregistrement(.DataField) = ValeurDuChamp
            NomChampDonneeLie = .DataField
            .DataField = ""
            .DataField = NomChampDonneeLie
        'End If
    End With
    
End Sub

'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' R�le      : Permet l'affichage format�e d'un nombre (si 0 alors affichage en fonction du format fourni)
' Entr�es : LeControleConcerne -> Repr�sente le contr�le � mofifier
'                      ValeurDuNombre -> Valeur du nombre
'                        FormatSouhaite -> Format souhait� (fonction de l'instruction format)
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Sub AffichageNombre(ByRef LeControleConcerne As Control, _
                                              ByVal ValeurDuNombre As Variant, _
                                              ByVal FormatSouhaite As String)
    
    '--- aiguillage en cas d'erreur ---
    On Error Resume Next
    
    '--- d�claration ---
    
    '--- modification ---
    With LeControleConcerne
        
        If TypeOf LeControleConcerne Is Label Then
            
            '--- cas de l'outil Label ---
            If IsNumeric(ValeurDuNombre) = True Then
                .Caption = Format(ValeurDuNombre, FormatSouhaite)
            Else
                .Caption = Format(0, FormatSouhaite)
            End If
        
        ElseIf TypeOf LeControleConcerne Is TextBox Then
            
            '--- cas de l'outil TextBox ---
            If IsNumeric(ValeurDuNombre) = True Then
                .Text = Format(ValeurDuNombre, FormatSouhaite)
            Else
                .Text = Format(0, FormatSouhaite)
            End If
        
        Else
        End If
    
    End With

End Sub

'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' R�le      : Permet de retourner une valeur num�rique d'un contr�le
' Entr�es :       LeControleConcerne -> Repr�sente le contr�le � v�rifier
'                 ValeurSiNonNumerique -> Valeur � retourner si le contr�le n'est pas num�rique
' Retours :          ControleSiNombre -> Valeur num�rique du contr�le ou ValeurSiNonNumerique si non num�rique
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function ControleSiNombre(ByRef LeControleConcerne As Control, _
                                                        ByVal ValeurSiNonNumerique As Variant) As Variant

    
    '--- aiguillage en cas d'erreur ---
    On Error Resume Next
    
    '--- d�claration ---
    
    '--- modification ---
    With LeControleConcerne
        
        If TypeOf LeControleConcerne Is Label Then
            
            '--- cas de l'outil Label ---
            If IsNumeric(LeControleConcerne.Caption) = True Then
                ControleSiNombre = CDbl(LeControleConcerne.Caption)
            Else
                ControleSiNombre = ValeurSiNonNumerique
            End If
        
        ElseIf TypeOf LeControleConcerne Is TextBox Then
            
            '--- cas de l'outil TextBox ---
            If IsNumeric(LeControleConcerne.Text) = True Then
                ControleSiNombre = CDbl(LeControleConcerne.Text)
            Else
                ControleSiNombre = ValeurSiNonNumerique
            End If
        
        Else
            
            '--- tous les autres cas ---
            ControleSiNombre = ValeurSiNonNumerique
        
        End If
    End With

End Function

'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' R�le      : Formate un r�pertoire avec divers contr�les
' Entr�es : Repertoire -> Repertoire � formater
' Retours : FormatRepertoire = R�pertoire contr�l� et format�
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Function FormateRepertoire(ByVal Repertoire As String) As String
    
    '--- aiguillage en cas d'erreur ---
    On Error Resume Next
   
    '--- affectation ---
    FormateRepertoire = ""
    
    '--- affectation ---
    Repertoire = Trim(Repertoire)
    
    '--- contr�le chaine vide ---
    If Repertoire = "" Then Exit Function
    
    '--- contr�le de la forme du chemin ---
    If Len(Repertoire) >= 3 Then
        If Right(Repertoire, 1) <> "\" Then Repertoire = Repertoire & "\"
        FormateRepertoire = Repertoire
    End If

End Function

'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' R�le      : Permet le codage d'un mot de passe
' Entr�es : MotDePasseACoder -> Mot de passe � coder
' Retours : CodeMotDePasse = Mot de passe cod�
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function CodeMotDePasse(ByVal MotDePasseACoder As String) As String
    
    '--- aiguillage en cas d'erreur ---
    On Error Resume Next

    '--- d�claration ---
    Dim a As Integer
    
    '--- affectation ---
    CodeMotDePasse = ""
    
    If MotDePasseACoder = "" Then
        Exit Function
    Else
        For a = 1 To Len(MotDePasseACoder)
            CodeMotDePasse = CodeMotDePasse & Chr(Asc(Mid(MotDePasseACoder, a, 1)) + 10)
        Next a
    End If
    
End Function

'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' R�le      : Permet le d�codage d'un mot de passe cod�
' Entr�es : MotDePasseADecoder -> Mot de passe � d�coder
' Retours : DecodeMotDePasse = Mot de passe d�cod�
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function DecodeMotDePasse(ByVal MotDePasseADecoder As String) As String
    
    '--- aiguillage en cas d'erreur ---
    On Error Resume Next

    '--- d�claration ---
    Dim a As Integer
    
    '--- affectation ---
    DecodeMotDePasse = ""
    
    If MotDePasseADecoder = "" Then
        Exit Function
    Else
        For a = 1 To Len(MotDePasseADecoder)
            DecodeMotDePasse = DecodeMotDePasse & Chr(Asc(Mid(MotDePasseADecoder, a, 1)) - 10)
        Next a
    End If
    
End Function

'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' R�le      : Conversion d'un temps en une chaine de caract�res exploitable
' Entr�es : TempsEnSecondes -> Temps en secondes � convertir
' Retours :                   CTemps -> Chaine retourn�e en jours, heures, minutes, secondes
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function CTemps(ByVal TempsEnSecondes As Double) As String
    
    '--- aiguillage en cas d'erreur ---
    On Error Resume Next
   
    '--- d�claration ---
    Dim TempsNegatif As Boolean
    Dim NbrJours As Double, _
            NbrHeures As Double, _
            NbrMinutes As Double, _
            NbrSecondes As Double, _
            Reste As Double
        
    '--- controle du signe du temps ---
    If TempsEnSecondes < 0 Then
        TempsEnSecondes = Abs(TempsEnSecondes)
        TempsNegatif = True
    End If
    
    '---  calcul des valeurs ---
    NbrJours = Int(TempsEnSecondes / 86400#)
    Reste = TempsEnSecondes Mod 86400#
    NbrHeures = Int(Reste / 3600#)
    Reste = Reste Mod 3600#
    NbrMinutes = Int(Reste / 60#)
    NbrSecondes = Reste Mod 60#

    '--- affectation ---
    If TempsNegatif = True Then
        CTemps = "-"
    Else
        CTemps = ""
    End If

    '--- construction de la chaine ---
    If NbrJours > 0 Then CTemps = Format(NbrJours, " 00 ") & "j"
    CTemps = Trim(CTemps & Format(NbrHeures, " 00:") & _
                                               Format(NbrMinutes, "00:") & _
                                               Format(NbrSecondes, "00"))

End Function

'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' R�le      : Conversion d'un temps en une chaine de caract�res exploitable
' Entr�es : TempsEnSecondes -> Temps en secondes � convertir
' Retours : CTemps -> Chaine retourn�e en heures, minutes, secondes 99:59:59 possible
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function CTemps2(ByVal TempsEnSecondes As Double) As String
    
    '--- aiguillage en cas d'erreur ---
    On Error Resume Next
   
    '--- d�claration ---
    Dim TempsNegatif As Boolean
    Dim NbrHeures As Double, _
            NbrMinutes As Double, _
            NbrSecondes As Double, _
            Reste As Double
            
    '--- controle du signe du temps ---
    If TempsEnSecondes < 0 Then
        TempsEnSecondes = Abs(TempsEnSecondes)
        TempsNegatif = True
    End If
    
    '---  calcul des valeurs ---
    NbrHeures = Int(TempsEnSecondes / 3600#)
    Reste = TempsEnSecondes Mod 3600#
    NbrMinutes = Int(Reste / 60#)
    NbrSecondes = Reste Mod 60#

    '--- affectation ---
    If TempsNegatif = True Then
        CTemps2 = "-"
    Else
        CTemps2 = ""
    End If

    '--- construction de la chaine ---
    CTemps2 = Trim(CTemps2 & Format(NbrHeures, " #00:") & _
                                                   Format(NbrMinutes, "00:") & _
                                                   Format(NbrSecondes, "00"))

End Function

'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' R�le      : Conversion d'un temps texte (00:00 ou 00:00:00) en un temps en secondes
' Entr�es :                        TempsTexte -> Temps en format texte 00:00 ou 00:00:00
' Retours : CTempsTexteEnSecondes -> Valeur en secondes du temps texte
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function CTempsTexteEnSecondes(ByVal TempsTexte As String) As Long
    
    '--- aiguillage en cas d'erreur ---
    On Error Resume Next
   
    '--- d�claration ---
    Dim Secondes As Long, _
            Minutes  As Long, _
            Heures As Long
            
    '--- affectation ---
    CTempsTexteEnSecondes = 0
    
    '--- sortie directe chaine vide ---
    If TempsTexte = "" Then Exit Function
    
    Select Case Len(TempsTexte)
    
        Case 5
            '--- cas du temps 00:00 ---
            If IsNumeric(Left(TempsTexte, 2)) = True Then
                Minutes = CLng(Left(TempsTexte, 2))
            End If
            If IsNumeric(Right(TempsTexte, 2)) = True Then
                Secondes = CLng(Right(TempsTexte, 2))
            End If
            CTempsTexteEnSecondes = Minutes * 60 + Secondes
        
        Case 7
            '--- cas du temps 0:00:00 ---
            If IsNumeric(Left(TempsTexte, 1)) = True Then
                Heures = CLng(Left(TempsTexte, 1))
            End If
            If IsNumeric(Mid(TempsTexte, 3, 2)) = True Then
                Minutes = CLng(Mid(TempsTexte, 3, 2))
            End If
            If IsNumeric(Right(TempsTexte, 2)) = True Then
                Secondes = CLng(Right(TempsTexte, 2))
            End If
            CTempsTexteEnSecondes = Heures * 3600 + Minutes * 60 + Secondes
        
        Case 8
            '--- cas du temps 00:00:00 ---
            If IsNumeric(Left(TempsTexte, 2)) = True Then
                Heures = CLng(Left(TempsTexte, 2))
            End If
            If IsNumeric(Mid(TempsTexte, 4, 2)) = True Then
                Minutes = CLng(Mid(TempsTexte, 4, 2))
            End If
            If IsNumeric(Right(TempsTexte, 2)) = True Then
                Secondes = CLng(Right(TempsTexte, 2))
            End If
            CTempsTexteEnSecondes = Heures * 3600 + Minutes * 60 + Secondes
        
        Case Else
    
    End Select
    
End Function

'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' R�le        : V�rifie l'existence d'un fichier
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function FileExist(ByVal NomFichier As String) As Boolean
    
    '--- aiguillage en cas d'erreur ---
    On Error Resume Next
   
    '--- d�claration ---
    Dim Resultat
    
    '---  recherche du fichier ---
    'Err = 53 -> fichier introuvable
    'Err = 76 -> chemin d'acc�s introuvable
    'Resultat = GetAttr(NomFichier)
    'If Err = 53 Or Err = 76 Then
    '    FileExist = False
    'Else
    '    FileExist = True
    'End If
    
    Resultat = GetAttr(NomFichier)
    If Err = 0 Then
        FileExist = True
    Else
        FileExist = False
    End If
    
End Function

'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' R�le        : Retourne la valeur de variable moins 1
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function Pred(ByVal Variable As Variant)
    
    '--- aiguillage en cas d'erreur ---
    On Error Resume Next
    
    '--- soustraction ---
    Pred = Variable - 1

End Function

'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' R�le        : Retourne la valeur de variable plus 1
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function Succ(ByVal Variable As Variant)
    
    '--- aiguillage en cas d'erreur ---
    On Error Resume Next
   
    '--- addition ---
    Succ = Variable + 1

End Function

'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' R�le        : Effectue l'op�ration variable = variable + 1
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Sub Inc(ByRef Variable As Variant)
    
    '--- aiguillage en cas d'erreur ---
    On Error Resume Next
    
    '--- addition ---
    Variable = Variable + 1

End Sub

'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' R�le        : Permet de convertir un texte contenant la lettre Y en Phi
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function CTextePhi(ByVal Texte As String) As String
       
    '--- aiguillage en cas d'erreur ---
    On Error Resume Next
 
    '--- d�claration ---
    Dim PositionCaractere As Long
    Dim CaractereRecherche As String * 1
    
    '--- analyse en fonction de l'�tat ---
    CaractereRecherche = Chr(221)
    If Len(Texte) > 0 Then
        Do
            PositionCaractere = InStr(Texte, CaractereRecherche)
            If PositionCaractere > 0 Then
                Mid(Texte, PositionCaractere, 1) = Chr(CODE_ASCII_PHI)    'Code ASCII de PHI = 216
            End If
        Loop Until PositionCaractere <= 0
    End If

    '--- valeur de retour ---
    CTextePhi = Texte

End Function

'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' R�le        : Permet de centrer une fenetre sur l'�cran et d'afficher le nom du programme
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Sub Centrefenetre(ByRef fenetre As Form, _
                                        Optional ByVal NomDuProgramme As String = "", _
                                        Optional ByVal AjoutSurAxeX As Long = 0, _
                                        Optional ByVal AjoutSurAxeY As Long = 0)
    
    '--- aiguillage en cas d'erreur ---
    On Error Resume Next
    
    With fenetre
        
        '--- titre de la fenetre ---
        If NomDuProgramme <> "" Then
            .Caption = NomDuProgramme & .Caption
        End If
        
        '--- valeurs pour obtenir le centrage ---
        If .MDIChild = True Then
            .Left = (OccFPrincipale.ScaleWidth - .Width) \ 2 + AjoutSurAxeX
            .Top = (OccFPrincipale.ScaleHeight - .Height) \ 2 + AjoutSurAxeY
        Else
            .Left = (Screen.Width - .Width) \ 2 + AjoutSurAxeX
            .Top = (Screen.Height - .Height) \ 2 + AjoutSurAxeY
        End If
        
    End With

End Sub

'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' R�le        : Effectue l'op�ration variable = variable - 1
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Sub Dec(Variable)
    
    '--- aiguillage en cas d'erreur ---
    On Error Resume Next
    
    '--- soustraction ---
    Variable = Variable - 1

End Sub

'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' R�le      : Attendre Xsecondes sans bloquer des �v�nements
' Entr�es : TempsEnSecondes -> Temps en secondes de la pause
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Sub Pause(ByVal TempsEnSecondes As Long)
    
    '--- aiguillage en cas d'erreur ---
    On Error Resume Next
   
    '--- d�claration ---
    Dim MemDateDepart  As Date                      'm�moire de la date de d�part pour compter le temps
        
    '--- affectation ---
    MemDateDepart = Now
    
    '--- contr�le du temps ---
    Do
        DoEvents
    Loop Until DateDiff("s", MemDateDepart, Now) >= TempsEnSecondes

End Sub

'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' R�le        : Conversion d'une chaine num�rique en h�xad�cimal (4 caract�res)
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function CHex(ByVal ChaineNumerique As String) As String
    
    '--- aiguillage en cas d'erreur ---
    On Error Resume Next

    '--- d�claration ---
    Dim Nombre As Long
    
    '--- controle du format ---
    If IsNumeric(ChaineNumerique) = True Then
    
        '--- affectation ---
        Nombre = CLng(ChaineNumerique)
        
        '--- controle du nombre ---
        If Nombre < 0 Then
            If Nombre >= -32768 Then                    'cas d'un INTEGER
                Nombre = Abs(Nombre) + 32767
            Else
                CHex = "ERREUR"
                Exit Function
            End If
        End If
        
        '--- conversion ---
        CHex = Right("0000" & Hex$(Nombre), 4)
    
    Else
        CHex = "ERREUR"
    End If

End Function

'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' R�le        : Extrait un bit d'une chaine binaire (16 bits)
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function Bit(ByVal ChaineBinaire As String, ByVal EmplacementBit As Integer) As Integer
    
    '--- aiguillage en cas d'erreur ---
    On Error Resume Next

    '--- contr�le ---
    If ChaineBinaire = "ERREUR" Or Len(ChaineBinaire) <> 16 Then
        Exit Function
    End If
    
    '--- extraction ---
    Bit = CInt(Mid(ChaineBinaire, 16 - EmplacementBit, 1))

End Function

'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' R�le        : Conversion d'une chaine num�rique en binaire (16 bits)
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function CBin(ByVal ChaineNumerique As String) As String
    
    '--- aiguillage en cas d'erreur ---
    On Error Resume Next

    '--- d�claration ---
    Dim ChaineHexa As String * 4
    Static MemPassage As Boolean
    Static TCorrespondances(0 To 15) As String * 4

    If MemPassage = False Then

        '--- remplissage du tableau ---
        TCorrespondances(0) = "0000"
        TCorrespondances(1) = "0001"
        TCorrespondances(2) = "0010"
        TCorrespondances(3) = "0011"
        TCorrespondances(4) = "0100"
        TCorrespondances(5) = "0101"
        TCorrespondances(6) = "0110"
        TCorrespondances(7) = "0111"
        TCorrespondances(8) = "1000"
        TCorrespondances(9) = "1001"
        TCorrespondances(10) = "1010"
        TCorrespondances(11) = "1011"
        TCorrespondances(12) = "1100"
        TCorrespondances(13) = "1101"
        TCorrespondances(14) = "1110"
        TCorrespondances(15) = "1111"
        
        '--- affectation ---
        MemPassage = True

    End If
    
    '--- conversion en hexad�cimale et contr�le ---
    ChaineHexa = CHex(ChaineNumerique)
    If ChaineHexa = Left("ERREUR", 4) Then
        CBin = "ERREUR"
        Exit Function
    End If
        
    '--- conversion en binaire ---
    CBin = TCorrespondances(CByte("&h" & Left(ChaineHexa, 1)))
    CBin = CBin & TCorrespondances(CByte("&h" & Mid(ChaineHexa, 2, 1)))
    CBin = CBin & TCorrespondances(CByte("&h" & Mid(ChaineHexa, 3, 1)))
    CBin = CBin & TCorrespondances(CByte("&h" & Right(ChaineHexa, 1)))

End Function

'---- retourne le mot bas d'un mot long ---
Function LOWORD(ByVal l As Long) As Integer
    On Error Resume Next
    LOWORD = l And &HFFFF&
End Function

'---- retourne le mot haut d'un mot long ---
Function HIWORD(ByVal l As Long) As Integer
    On Error Resume Next
    HIWORD = CInt((l And &HFFFF0000) / 65536)
End Function

'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' R�le      : Filtre la frappe d'une touche (code ASCII uniquement) au clavier en fonction du format des donn�es
' Entr�es :         CodeASCIITouche -> Code ASCII de la touche frapp�e
'                 FormatDonneesChoisi -> Format de donn�es choisi
' Retours :         CodeASCIITouche -> Code ASCII de la touche filtr�e
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Sub FiltreToucheASCII(ByRef CodeASCIITouche As Integer, _
                                                ByVal FormatDonneesChoisi As Long, _
                                                Optional ByVal NbrCaracteresMaxi As Variant)
    
    '--- aiguillage en cas d'erreur ---
    On Error Resume Next
    
    '--- d�claration ---
    Dim PositionCurseur As Integer
    Dim Texte As String
    Dim ControleEdition As Control
            
    '--- affectation ---
    Set ControleEdition = Screen.ActiveControl
    
    '--- limiter le maximum de caract�res ---
    If IsMissing(NbrCaracteresMaxi) = False Then
        If TypeOf ControleEdition Is TextBox Or TypeOf ControleEdition Is MaskEdBox Then
            ControleEdition.MaxLength = NbrCaracteresMaxi
        End If
    End If
    
    '--- analyse de la touche frapp�e ---
    Select Case FormatDonneesChoisi
                
        Case DONNEES.D_GENERALE
            '--- tous les caract�res ---
            Select Case CodeASCIITouche
                Case CODE_ASCII_DOLLAR: CodeASCIITouche = CODE_ASCII_PHI
                Case vbKeyReturn: CodeASCIITouche = 0
                Case Else
            End Select
                    
        Case DONNEES.D_GENERALE_MINUSCULES
            '--- tous les caract�res en minuscules ---
            Select Case CodeASCIITouche
                Case CODE_ASCII_DOLLAR: CodeASCIITouche = CODE_ASCII_PHI
                Case vbKeyReturn: CodeASCIITouche = 0
                Case Else: CodeASCIITouche = Asc(LCase(Chr(CodeASCIITouche)))
            End Select
        
        Case DONNEES.D_GENERALE_MAJUSCULES
            '--- tous les caract�res en majuscules ---
            Select Case CodeASCIITouche
                Case CODE_ASCII_DOLLAR: CodeASCIITouche = CODE_ASCII_PHI
                Case vbKeyReturn: CodeASCIITouche = 0
                Case Else: CodeASCIITouche = Asc(UCase(Chr(CodeASCIITouche)))
            End Select
    
        Case DONNEES.D_TEXTE
            '--- lettres de a � z en minuscules et majuscules ---
            Select Case CodeASCIITouche
                Case CODE_ASCII_DOLLAR: CodeASCIITouche = CODE_ASCII_PHI
                Case Asc(Chr(CODE_ASCII_PHI)), _
                         vbKeySpace, vbKeyBack, _
                         Asc("a") To Asc("z"), _
                         vbKeyA To vbKeyZ, _
                         Asc("�") To Asc("�")
                Case Else: CodeASCIITouche = 0
            End Select
        
        Case DONNEES.D_TEXTE_MINUSCULES
            '--- lettres de a � z en minuscules ---
            Select Case CodeASCIITouche
                Case CODE_ASCII_DOLLAR: CodeASCIITouche = CODE_ASCII_PHI
                Case Asc(Chr(CODE_ASCII_PHI)), _
                         vbKeySpace, vbKeyBack, _
                         Asc("a") To Asc("z"), _
                         vbKeyA To vbKeyZ, _
                         Asc("�") To Asc("�"), _
                         vbKey0 To vbKey9, _
                            CodeASCIITouche = Asc(LCase(Chr(CodeASCIITouche)))
                Case Else: CodeASCIITouche = 0
            End Select
        
        Case DONNEES.D_TEXTE_MINUSCULES_NUMERIQUES
            '--- lettres de a � z en minuscules ou touches num�riques ---
            Select Case CodeASCIITouche
                Case CODE_ASCII_DOLLAR: CodeASCIITouche = CODE_ASCII_PHI
                Case Asc("."): CodeASCIITouche = Asc(",")
                Case Asc(Chr(CODE_ASCII_PHI)), _
                         vbKeySpace, vbKeyBack, _
                         Asc("a") To Asc("z"), _
                         vbKeyA To vbKeyZ, _
                         Asc("�") To Asc("�"), _
                         vbKey0 To vbKey9, Asc(","), _
                            CodeASCIITouche = Asc(LCase(Chr(CodeASCIITouche)))
                Case Else: CodeASCIITouche = 0
            End Select
        
        Case DONNEES.D_TEXTE_MAJUSCULES
            '--- lettres de a � z en majuscules ---
            Select Case CodeASCIITouche
                Case CODE_ASCII_DOLLAR: CodeASCIITouche = CODE_ASCII_PHI
                Case Asc(Chr(CODE_ASCII_PHI)), _
                         vbKeySpace, vbKeyBack, _
                         Asc("a") To Asc("z"), _
                         vbKeyA To vbKeyZ, _
                         Asc("�") To Asc("�")
                            CodeASCIITouche = Asc(UCase(Chr(CodeASCIITouche)))
                Case Else: CodeASCIITouche = 0
            End Select
        
        Case DONNEES.D_TEXTE_MAJUSCULES_NUMERIQUES
            '--- lettres de a � z en majuscules ou touches num�riques ------
            Select Case CodeASCIITouche
                Case CODE_ASCII_DOLLAR: CodeASCIITouche = CODE_ASCII_PHI
                Case Asc(Chr(CODE_ASCII_PHI)), _
                         vbKeySpace, vbKeyBack, _
                         Asc("a") To Asc("z"), _
                         vbKeyA To vbKeyZ, _
                         Asc("�") To Asc("�"), _
                         vbKey0 To vbKey9, Asc(","), _
                            CodeASCIITouche = Asc(UCase(Chr(CodeASCIITouche)))
                Case Else: CodeASCIITouche = 0
            End Select
                    
        Case DONNEES.D_NBR_NATURELS
            '--- touches num�riques sans d�cimale positif (de 0 � x) ---
            Select Case CodeASCIITouche
                Case vbKey0 To vbKey9, vbKeyBack
                Case Else: CodeASCIITouche = 0
            End Select
 
        Case DONNEES.D_NBR_RELATIFS
            '--- touches num�riques sans d�cimale (de -x � +x) ---
            Select Case CodeASCIITouche
                Case vbKey0 To vbKey9, Asc("-"), vbKeyBack
                Case Else: CodeASCIITouche = 0
            End Select
        
        Case DONNEES.D_NBR_REELS
            '--- touches num�riques avec d�cimale (de -x,x... � + x,xx...) ---
            Select Case CodeASCIITouche
                Case Asc("."): CodeASCIITouche = Asc(",")
                Case vbKey0 To vbKey9, Asc(","), Asc("-"), vbKeyBack
                Case Else: CodeASCIITouche = 0
            End Select
        
        Case DONNEES.D_NBR_REELS_POSITIFS
            '--- touches num�riques avec d�cimale (de 0 � + x,xx...) ---
            Select Case CodeASCIITouche
                Case Asc("."): CodeASCIITouche = Asc(",")
                Case vbKey0 To vbKey9, Asc(","), vbKeyBack
                Case Else: CodeASCIITouche = 0
            End Select
            
        Case DONNEES.D_HEURE_HHMM
            '--- format heure HH:MM ---
            If TypeOf ControleEdition Is TextBox Then ControleEdition.MaxLength = 5
            Select Case CodeASCIITouche
                Case Asc("."): CodeASCIITouche = Asc(":")
                Case Asc("0") To Asc("9"), Asc(":"), vbKeyBack
                Case Else: CodeASCIITouche = 0
            End Select
        
        Case DONNEES.D_HEURE_HHMMSS
            '--- format heure HH:MM:SS ---
            If TypeOf ControleEdition Is TextBox Then ControleEdition.MaxLength = 8
            Select Case CodeASCIITouche
                Case Asc("."): CodeASCIITouche = Asc(":")
                Case Asc("0") To Asc("9"), Asc(":"), vbKeyBack
                Case Else: CodeASCIITouche = 0
            End Select
        
        Case DONNEES.D_DATE_JJMMAAAA
            '--- format date JJ/MM/AAAA ---
            If TypeOf ControleEdition Is TextBox Then ControleEdition.MaxLength = 10
            Select Case CodeASCIITouche
                Case Asc("0") To Asc("9"), Asc("/"), vbKeyBack
                Case Else: CodeASCIITouche = 0
            End Select

        Case DONNEES.D_TELEPHONE
            '--- format t�l�phone ---
            Select Case CodeASCIITouche
                Case Asc("0") To Asc("9"), Asc("-"), vbKeyBack
                Case Else: CodeASCIITouche = 0
            End Select
        
        Case DONNEES.D_FAX
            '--- format fax ---
            Select Case CodeASCIITouche
                Case Asc("0") To Asc("9"), Asc("-"), vbKeyBack
                Case Else: CodeASCIITouche = 0
            End Select
        
        Case DONNEES.D_CODE_CLIENT, DONNEES.D_CODE_FOURNISSEUR
            '--- code client, fournisseur ---
            Select Case CodeASCIITouche
                Case Asc("0") To Asc("9"), vbKeyBack, _
                         Asc("a") To Asc("z"), _
                         vbKeyA To vbKeyZ
                            CodeASCIITouche = Asc(UCase(Chr(CodeASCIITouche)))
                Case Else: CodeASCIITouche = 0
            End Select
    
        Case DONNEES.D_TYPE_DE_PRIX
            '--- type de prix (U ou E en majuscules) ---
            Select Case CodeASCIITouche
                Case Asc("e"), Asc("u"), vbKeyE, vbKeyU, vbKeyBack
                    CodeASCIITouche = Asc(UCase(Chr(CodeASCIITouche)))
                Case Else: CodeASCIITouche = 0
            End Select
    
        Case DONNEES.D_JOUR_OU_NUIT
            '--- format nuit ou jour (J ou N en majuscules) ---
            Select Case CodeASCIITouche
                Case Asc("j"), Asc("n"), vbKeyJ, vbKeyN, vbKeyBack
                    CodeASCIITouche = Asc(UCase(Chr(CodeASCIITouche)))
                Case Else: CodeASCIITouche = 0
            End Select

        Case DONNEES.D_AVEC_JUMELAGE
            '--- format avec jumelage (Espace ou D (double)) ---
            Select Case CodeASCIITouche
                Case vbKeySpace, Asc("d"), vbKeyD, vbKeyBack
                    CodeASCIITouche = Asc(UCase(Chr(CodeASCIITouche)))
                Case Else: CodeASCIITouche = 0
            End Select
        
        Case DONNEES.D_MANU_AUTO
            '--- format manu auto (A ou M en majuscules) ---
            Select Case CodeASCIITouche
                Case Asc("a"), Asc("m"), vbKeyA, vbKeyM, vbKeyBack
                    CodeASCIITouche = Asc(UCase(Chr(CodeASCIITouche)))
                Case Else: CodeASCIITouche = 0
            End Select
        
        Case DONNEES.D_CODE_POSTAL
            '--- code postal ---
            Select Case CodeASCIITouche
                Case Asc("0") To Asc("9"), Asc("-"), vbKeyBack, _
                         Asc("a") To Asc("z"), _
                         vbKeyA To vbKeyZ
                            CodeASCIITouche = Asc(UCase(Chr(CodeASCIITouche)))
                Case Else: CodeASCIITouche = 0
            End Select
        
        Case DONNEES.D_SECURITE_SOCIALE
            '--- s�curit� sociale ---
            Select Case CodeASCIITouche
                Case vbKey0 To vbKey9, vbKeySpace, vbKeyBack
                Case Else: CodeASCIITouche = 0
            End Select
        
        Case Else

    End Select

    '--- analyse du mode de surfrappe ---
    'Debug.Print CodeASCIITouche
    If CodeASCIITouche > 0 And ModeSurFrappe = True Then
    
        If TypeOf ControleEdition Is TextBox Or _
           TypeOf ControleEdition Is MaskEdBox Or _
           TypeOf ControleEdition Is DataCombo Then
        
            '--- s�paration des touches ---
            Select Case CodeASCIITouche
                
                Case vbKeyReturn, vbKeyBack, vbKeyDelete, vbKeyClear, _
                         vbKeyLeft, vbKeyUp, vbKeyRight, vbKeyDown, _
                         vbKeyPageUp, vbKeyPageDown, vbKeyEnd, vbKeyHome
                
                Case Else
                    '--- remplacer le caract�re ---
                    With ControleEdition
                        
                        '--- affectation ---
                        PositionCurseur = .SelStart
                        Texte = .Text
                        
                        '--- remplacement de caract�re ---
                        If PositionCurseur < Len(Texte) Then
                            PositionCurseur = Succ(PositionCurseur)
                            Mid(Texte, PositionCurseur, 1) = Chr(CodeASCIITouche)
                            .Text = Texte
                            .SelStart = PositionCurseur
                            .Refresh
                            CodeASCIITouche = 0
                        End If
                    
                    End With
            
            End Select
    
        End If
    
    End If
    
End Sub

'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' R�le      : Filtre la frappe d'une touche de fonction par rapport au contr�le actif
' Entr�es :                 CodeTouche -> Code de la touche frapp�e
'                                     CodeShift -> Code Shift de la touche frapp�e
' Retours :                 CodeTouche -> Code de la touche filtr�e
'                                     CodeShift -> Code Shift de la touche filtr�e
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Sub FiltreToucheFonction(ByRef CodeTouche As Integer, _
                                                     ByRef CodeShift As Integer)
    
    '--- aiguillage en cas d'erreur ---
    On Error Resume Next

    '--- d�claration ---
    Static PremierPassage As Boolean
   
    If TypeOf Screen.ActiveControl Is TextBox Or _
       TypeOf Screen.ActiveControl Is MaskEdBox Then
                    
        '--- analyse de la touche frapp�e ---
        Select Case CodeTouche
            Case vbKeyInsert
                If CodeShift = 0 Then
                    ModeSurFrappe = Not (ModeSurFrappe)
                End If
            Case vbKeyDown, vbKeyReturn: SendKeys "{tab}": CodeTouche = 0
            Case vbKeyUp: SendKeys "+{tab}": CodeTouche = 0
            Case Else
        End Select

    ElseIf TypeOf Screen.ActiveControl Is DataCombo Then
    
        '--- analyse de la touche frapp�e ---
        'Select Case CodeTouche
        '    Case vbKeyF12
                'Screen.ActiveControl.Text = Screen.ActiveControl.Tag
                'If PremierPassage = False Then
                '    SendKeys "{F12}"
                '    PremierPassage = True
                'Else
                '    PremierPassage = False
                'End If
        '        CodeTouche = 0
        '    Case Else
        'End Select
    
    ElseIf TypeOf Screen.ActiveControl Is RichTextBox Then
    
        '--- pas d'action ---
    
    Else
    
    End If
    
End Sub

'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' R�le      : Change la couleur de fond des boutons en fonction du d�placement de la souris dans l'�cran
' Entr�es : ObjetConcerne -> Indique l'objet concern� par la modication de couleur de fond
'                                             Si l'objet est absent on r�tabli la couleur de fond de base
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Sub ChangeCouleursBoutons2(ByRef fenetre As Form, _
                                                              Optional ByRef ObjetConcerne As Variant)

    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next
    
    '--- constantes priv�es ---
    Const COULEUR_DE_SELECTION As Long = COULEURS.CYAN_1
    
    '--- d�claration ---
    Static MemCouleurFond As Long
    Dim UnControle As Control
    
    '--- affectation ---
    If IsEmpty(MemCouleurFond) = True Then
        MemCouleurFond = COULEURS.GRIS_SYSTEME
    End If
    
    If IsMissing(ObjetConcerne) = True Then
        
        '--- changement de la couleur ---
        For Each UnControle In fenetre.Controls
            If TypeOf UnControle Is CommandButton Then
                If UnControle.BackColor = COULEUR_DE_SELECTION Then
                    UnControle.BackColor = MemCouleurFond
                End If
            End If
        Next
    
    Else
    
        '--- changement de la couleur ---
        If TypeOf ObjetConcerne Is CommandButton Then
            If ObjetConcerne.Enabled = True And ObjetConcerne.Visible = True And ObjetConcerne.BackColor <> COULEUR_DE_SELECTION Then
                MemCouleurFond = ObjetConcerne.BackColor
                ObjetConcerne.BackColor = COULEUR_DE_SELECTION
            End If
        End If
    
    End If
    
End Sub

'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' R�le      : Affiche un num�ro et un message refl�tant une erreur
' Entr�es : TitreMessage       -> Titre du message
'                NumErreur            -> Num�ro de l'erreur
'                DescriptionErreur -> Description de l'erreur
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Function MessageErreur(ByVal TitreMessage As String, _
                                         ByVal DescriptionErreur As String, _
                                         Optional ByVal NumErreur As Long = 0) As String
    
    '--- aiguillage en cas d'erreur ---
    On Error Resume Next
   
    Call Log(DescriptionErreur)
    '--- affichage du message d'erreur ---
    Screen.MousePointer = vbDefault
    If NumErreur = 0 Then
        Bidon = AppelFenetre(F_MESSAGE, _
                                          TitreMessage, _
                                          DescriptionErreur, _
                                          TYPES_MESSAGES.T_ATTENTION, _
                                          TYPES_BOUTONS.T_CONFIRMER, _
                                          EMPLACEMENT_FOCUS.E_SUR_CONFIRMER)
    Else
        Bidon = AppelFenetre(F_MESSAGE, _
                                          TitreMessage, _
                                          vbCrLf & "Erreur n� " & NumErreur & vbCrLf & vbCrLf & DescriptionErreur, _
                                          TYPES_MESSAGES.T_ATTENTION, _
                                          TYPES_BOUTONS.T_CONFIRMER, _
                                          EMPLACEMENT_FOCUS.E_SUR_CONFIRMER)
    End If

End Function

'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' R�le      : Permet l'affichage d'un texte avec un anti-rebond si le texte d�j� afficher est le m�me
' Entr�es : LeControleConcerne -> Repr�sente le contr�le � mofifier
'                                        Texte -> Texte � afficher
'                             CouleurPlan -> Couleur de premier plan du texte
'                            CouleurFond -> Couleur de fond du texte
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Sub AffichageTexte(ByRef LeControleConcerne As Control, _
                                           ByVal Texte As String, _
                                           Optional ByVal CouleurFond As Variant, _
                                           Optional ByVal CouleurPlan As Variant)
    
    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next
    
    '--- d�claration ---
    
    '--- modification ---
    With LeControleConcerne
        
        If TypeOf LeControleConcerne Is Label Then
            
            '--- couleur de fond ---
            If IsMissing(CouleurFond) = False Then
                If .BackColor <> CouleurFond Then
                    .BackColor = CouleurFond
                End If
            End If
            
            '--- couleur de premier plan ---
            If IsMissing(CouleurPlan) = False Then
                If .ForeColor <> CouleurPlan Then
                    .ForeColor = CouleurPlan
                End If
            End If
            
            '--- cas de l'outil Label ---
            If .Caption <> Texte Then
                .Caption = Texte
                .Refresh
            End If
            
        ElseIf TypeOf LeControleConcerne Is TextBox Then
        
            '--- couleur de fond ---
            If IsMissing(CouleurFond) = False Then
                If .BackColor <> CouleurFond Then
                    .BackColor = CouleurFond
                End If
            End If
            
            '--- couleur de premier plan ---
            If IsMissing(CouleurPlan) = False Then
                If .ForeColor <> CouleurPlan Then
                    .ForeColor = CouleurPlan
                End If
            End If
            
            '--- cas de l'outil TexBox ---
            If .Text <> Texte Then
                .Text = Texte
            End If
        
        ElseIf TypeOf LeControleConcerne Is MSHFlexGrid Or _
                  TypeOf LeControleConcerne Is VSFlexGrid Then
            
            '--- couleur de fond ---
            If IsMissing(CouleurFond) = False Then
                If .CellBackColor <> CouleurFond Then
                    .CellBackColor = CouleurFond
                End If
            End If
            
            '--- couleur de premier plan ---
            If IsMissing(CouleurPlan) = False Then
                If .CellForeColor <> CouleurPlan Then
                    .CellForeColor = CouleurPlan
                End If
            End If
            
            '--- cas de l'outil MSHFlexGrid or VSFlexGrid ---
            If .Text <> Texte Then
                .Text = Texte
            End If
        
        End If
    
    End With

End Sub

'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' R�le      : Permet l'affichage d'un texte par la matrice avec un anti-rebond si le texte d�j� afficher est le m�me
' Entr�es : LeControleConcerne -> Repr�sente le contr�le � mofifier
'                                        Texte -> Texte � afficher
'                                 NumLigne -> Repr�sente un num�ro de ligne
'                             NumColonne-> Repr�sente un num�ro de colonne
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Sub AffichageTexteMatrice(ByRef LeControleConcerne As Control, _
                                                       ByVal NumLigne As Long, _
                                                       ByVal NumColonne As Long, _
                                                       ByVal Texte As String)
    
    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next
    
    '--- d�claration ---
    
    '--- modification ---
    With LeControleConcerne
        
        If TypeOf LeControleConcerne Is MSHFlexGrid Or _
           TypeOf LeControleConcerne Is VSFlexGrid Then
            
            '--- cas de l'outil MSHFlexGrid or VSFlexGrid ---
            If .TextMatrix(NumLigne, NumColonne) <> Texte Then
                .TextMatrix(NumLigne, NumColonne) = Texte
            End If
        
        End If
    
    End With

End Sub

'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' R�le      : Conversion d'un temps en une chaine de caract�res exploitable
' Entr�es : TempsEnSecondes -> Temps en secondes � convertir
' Retours :                  CTemps -> Chaine retourn�e en minutes, secondes (99:99 possible)
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function CTemps3(ByVal TempsEnSecondes As Double) As String
    
    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next
   
    '--- d�claration ---
    Dim TempsNegatif As Boolean
    Dim NbrMinutes As Double, _
            NbrSecondes As Double

    '--- contr�le du signe du temps ---
    If TempsEnSecondes < 0 Then
        TempsEnSecondes = Abs(TempsEnSecondes)
        TempsNegatif = True
    End If
    
    '---  calcul des valeurs ---
    NbrMinutes = Int(TempsEnSecondes / 60#)
    NbrSecondes = TempsEnSecondes Mod 60#

    '--- affectation ---
    If TempsNegatif = True Then
        CTemps3 = "-"
    Else
        CTemps3 = ""
    End If

    '--- construction de la chaine ---
    CTemps3 = Trim(CTemps3 & Format(NbrMinutes, "00:") & _
                                                   Format(NbrSecondes, "00"))

End Function

'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' R�le      : Indique si un nombre est paire
' Entr�es : Nombre -> le nombre � contr�ler
' Retours :
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function Even(ByRef Nombre As Variant) As Boolean
    
    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next

    '--- contr�le du nombre ---
    If IsNumeric(Nombre) = True Then
        If CDbl(Nombre) - Int(CCur(Nombre)) = 0 Then
            Select Case Right(CStr(Nombre), 1)
                Case "0", "2", "4", "6", "8": Even = True
                Case Else: Even = False
            End Select
        Else
            '--- lancement de l'erreur "type imcompatible" ---
            Error 13
        End If
    Else
        '--- lancement de l'erreur "type imcompatible" ---
        Error 13
    End If

End Function

'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' R�le      : Indique si un nombre est impaire
' Entr�es : Nombre -> le nombre � contr�ler
' Retours :
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function Odd(ByRef Nombre As Variant) As Boolean
    
    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next
    
    '--- contr�le du nombre ---
    If IsNumeric(Nombre) = True Then
        If CDbl(Nombre) - Int(CCur(Nombre)) = 0 Then
            Select Case Right(CStr(Nombre), 1)
                Case "1", "3", "5", "7", "9": Odd = True
                Case Else: Odd = False
            End Select
        Else
            '--- lancement de l'erreur "type imcompatible" ---
            Error 13
        End If
    Else
        '--- lancement de l'erreur "type imcompatible" ---
        Error 13
    End If

End Function


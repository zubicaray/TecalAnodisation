Attribute VB_Name = "MGestionClipper"
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' R�le                    : MODULE GERANT L'ACCES A CLIPPER
' Nom                    : MGestionClipper.bas
' Date de cr�ation : 31/01/2012
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

'--- d�clarations obligatoires ---
Option Explicit

'--- options g�n�rales ---
Option Base 1
DefVar A-Z

'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' R�le      : Enregistrement des bains pour renseigner les fiches d'atelier pour CLIPPER
' Entr�es :                                      NumCharge -> Num�ro de la charge concern�e
' Retours : EnregistrementBainsPourCLIPPER -> "" = pas d'incident sinon num�ro de l'erreur
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function EnregistrementBainsPourCLIPPER(ByVal NumCharge As Integer) As String
    
    '--- aiguillage en cas d'erreurs ---
    On Error GoTo GestionErreurs

    '--- constantes priv�es ---

    '--- d�claration ---
    Dim a As Integer                                                       'pour les boucles FOR...NEXT
    Dim b As Integer                                                       'pour les boucles FOR...NEXT
    Dim NumFic As Integer, _
            NbrFichesAtelier As Integer
    Dim NbrPieces As Double
    Dim TempsDecimale As Double
    Dim TempsDecimaleTexte As String                        'temps d�cimale en texte
    Dim NumFicheAtelier As String
    Dim CoFrais As String                                               'centre de frais
    Dim BainAvecJumelage As String                             'chaine contenant le nombre de fiche d'atelier sinon un "+"
    Dim NumFicheProduction As String                          'n� de la fiche de production
    Dim DateEntreeEnLigne As Date, _
            DateArriveeAuDechargement As Date, _
            DateModification As Date, _
            DateEntreePoste As Date, _
            DateSortiePoste As Date
    Dim ChaineEnvoi As String                                       'chaine � envoyer pour CLIPPER

    '----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    
    '--- affectation ---
    EnregistrementBainsPourCLIPPER = ""
    DateModification = Now
    
    '----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    
    With TEtatsCharges(NumCharge)

        '----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

        '--- affectation des dates ---
        DateEntreeEnLigne = .DateEntreeEnLigne
        DateArriveeAuDechargement = .DateArriveeAuDechargement
    
        '----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    
        '--- calcul du temps d�cimale ---
        TempsDecimale = CDbl(DateDiff("s", DateEntreeEnLigne, DateArriveeAuDechargement)) / 3600#
    
        '----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

        '--- calcul du nombre de fiches d'atelier ---
        For a = LBound(.TDetailsCharges()) To UBound(.TDetailsCharges())
            With .TDetailsCharges(a)
                NumFicheAtelier = .NumCommandeInterne
                If NumFicheAtelier <> "" Then
                    Inc NbrFichesAtelier
                Else
                    Exit For
                End If
            End With
        Next a
    
        '----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
            
        '--- affectation du jumelage ---
        Select Case NbrFichesAtelier
            Case 1: BainAvecJumelage = 0
            Case Else: BainAvecJumelage = 1
        End Select
        
        '----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
        
        '--- affectation ---
        NumFic = FreeFile(1)
    
        '--- ouverture du fichier ---
        Open RepFicClipper & FIC_BAINS_ANODISATION For Append Shared As #NumFic

        '----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
        
        '--- lecture des d�tails de la charge ---
        For a = LBound(.TDetailsCharges()) To UBound(.TDetailsCharges())
                
            '--- affectation ---
            NumFicheAtelier = .TDetailsCharges(a).NumCommandeInterne
                
            If NumFicheAtelier <> "" Then

                '--- extraction du nombre de pi�ces ---
                NbrPieces = .TDetailsCharges(a).NbrPieces

                '--- enregistrement des d�tails des fiches de production ---
                For b = LBound(.TDetailsFichesProduction()) To UBound(.TDetailsFichesProduction())
                
                    With .TDetailsFichesProduction(b)
                    
                        If .NumPoste <> 0 Then
                   
                            '--- recherche du centre de frais ---
                            Bidon = RechercheCentreDeFraisAnodisation(NumPoste:=.NumPoste, CoFrais:=CoFrais)
                                                                              
                            If CoFrais <> "" Then
            
                                '--- donn�es � enregistrer ---
                                '1 : N� Phase de gamme (GACLEUNIK)
                                '2:  Code employ�
                                '3 : Heure D�but du bain (HH :MM :SS)
                                '4 : Heure de Fin du bain (HH :MM :SS)
                                '5 : Temps pass� (unit� de temps)
                                '6 : Quantit� de pi�ces r�alis�es
                                '7 : Date de pointage (JJ/MM/AAAA)
                                '8 : Centre de frais (bain sur lequel les pi�ces ont �t� trait�es)
                                '9 : N� ALEA (BAIN AVEC JUMELAGE)
                                
                                '--- calcul du temps d�cimale dans le poste ---
                                DateEntreePoste = .DateEntreePoste
                                DateSortiePoste = .DateSortiePoste
                                TempsDecimale = CDbl(DateDiff("s", DateEntreePoste, DateSortiePoste)) / 3600#
                                
                                '--- remplacement de la virgule par un point dans le temps d�cimale ---
                                TempsDecimaleTexte = Trim(CStr(TempsDecimale))
                                TempsDecimaleTexte = Replace(TempsDecimaleTexte, ",", ".")
                                
                                '--- construction de la chaine d'envoi ---
                                ChaineEnvoi = NumFicheAtelier & ";" & _
                                                        "BAIN" & ";" & _
                                                        Format(DateEntreePoste, "hhnnss") & ";" & _
                                                        Format(DateSortiePoste, "hhnnss") & ";" & _
                                                        TempsDecimaleTexte & ";" & _
                                                        NbrPieces & ";" & _
                                                        Format(DateEntreePoste, "dd/mm/yyyy") & ";" & _
                                                        CoFrais & ";" & _
                                                        BainAvecJumelage
                                
                                '--- enregistrement dans le fichier ---
                                Print #NumFic, ChaineEnvoi
                    
                            End If
    
                        Else
                        
                            '--- sortie directe si plus de fiche poste ---
                            Exit For
                        
                        End If
                        
                    End With
                
                Next b
                    
            Else
                            
                '--- sortie directe si plus de fiche d'atelier ---
                Exit For
                   
            End If
           
        Next a

    End With
    
    '--- fermeture du fichier ---
    Close #NumFic
    
    Exit Function

GestionErreurs:
    
    '--- for�age de la fermeture du fichier ---
    If NumFic > 0 Then Close #NumFic
    
    '--- valeur de retour ---
    EnregistrementBainsPourCLIPPER = CStr(Err.Number)
    
End Function

'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' R�le      : Extrait un enregistrement de la table des phases de CLIPPER
' Entr�es :     Enregistrement -> Enregistrement de la table des phases de CLIPPER
' Retours : TEnrFichesAtelier -> Tableau contenant l'image d'un enregistrement de la table des
'                                                  phases de CLIPPER
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Sub ExtraitEnrPhasesClipper(ByVal Enregistrement As ADODB.Recordset, _
                                                           ByRef TEnrPhasesClipper As EnrPhasesClipper)

    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next

    With TEnrPhasesClipper
            
        '--- extraction de l'enregistrement ---
        .GaCLeUnik = C_Nullite_Champ(Enregistrement, "GACLEUNIK", 0)
        .CoFrais = C_Nullite_Champ(Enregistrement, "COFRAIS", "")
        .CoCli = C_Nullite_Champ(Enregistrement, "COCLI", "")
        .NomClient = C_Nullite_Champ(Enregistrement, "NOMCLIENT", "")
        .Piece = C_Nullite_Champ(Enregistrement, "PIECE", "")
        .QteAf = C_Nullite_Champ(Enregistrement, "QTEAF", 0)
        .Desa1 = C_Nullite_Champ(Enregistrement, "DESA1", "")
        .DateLance = C_Nullite_Champ(Enregistrement, "DATE_LANCE", "")
        .Matiere = C_Nullite_Champ(Enregistrement, "MATIERE", "")
        .GamObs = C_Nullite_Champ(Enregistrement, "GAMOBS", "")
        .NumGamme = C_Nullite_Champ(Enregistrement, "NUMGAMME", "")
        .Naf = C_Nullite_Champ(Enregistrement, "NAF", "")
        
    End With
        
End Sub


Function TrimLeadingZeros(value)
    TrimLeadingZeros = value
    While Left(TrimLeadingZeros, 1) = "0" And TrimLeadingZeros <> "0"
        TrimLeadingZeros = Mid(TrimLeadingZeros, 2)
    Wend
End Function
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' R�le      : Recherche une phase CLIPPER
' Entr�es :             NumPhaseClipper -> Num�ro de la phase CLIPPER (GACLEUNIK)
' Retours : RecherchePhasesClipper -> TROUVE          = Enregistrement trouv� ou valid�
'                                                               NON_TROUVE = Recherche non trouv�e ou abandonn�e
'                                                                                          autres valeurs = N� du message d'erreur
'                                                               ATTENTION -> Les caract�ristiques de l'enregistrement se trouve dans la
'                                                                                       m�moire temporaire
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function RecherchePhasesClipper(ByVal NumPhaseClipper As Variant) As String
    
    '--- aiguillage en cas d'erreurs ---
    On Error GoTo GestionErreurs
    
    '--- d�claration ---
    Dim Requete As String
    Dim ConnexionBDNickelSQL As New ADODB.Connection
    Dim Enregistrement As New ADODB.Recordset
    
    Dim gac As Integer
    
    
    NumPhaseClipper = TrimLeadingZeros(NumPhaseClipper)
    'Call Log("NumPhaseClipper" & NumPhaseClipper)
    
    '--- contr�le ---
    If IsNumeric(NumPhaseClipper) = False Then
        RecherchePhasesClipper = NON_TROUVE
        Exit Function
    End If
    Screen.MousePointer = vbHourglass
    '--- ouverture de la connexion � la base de donn�es ANODISATION en SQL ---
    With ConnexionBDNickelSQL
        .ConnectionString = PARAMETRES_CONNEXION_BD_CLIPPER_HF
        .CursorLocation = adUseServer
        .Mode = adModeReadWrite
        .ConnectionTimeout = 2     'X secondes d'attente de connexion avant de lancer un message d'erreur
        .Open
    End With
    
    '--- recherche ---
    With Enregistrement
        'TMP
        '--- lancement de la requ�te ---
        Requete = "SELECT  CP.Compl�ment as NUMGAMME,'' AS GAMOBS,G.GACLEUNIK,A.NAF,A.COCLI,A.PIECE,C.NOM AS NOMCLIENT,A.QTEAF, A.DESA+  A.DESA2 +A.DESA3  AS DESA1 " & _
                "FROM  ((AFFAIRE  A INNER JOIN  CLIENT C   ON C.COCLI = A.COCLI ) " & _
                "INNER JOIN GAMME G on  G.NAF=A.NAF ) " & _
                "LEFT JOIN COMPLEMS  CP on  CP.Cl�unik=G.GACLEUNIK   and COPAR='GACPL01'   WHERE G.GACLEUNIK='" & NumPhaseClipper & "'"

                
        Requete = "SELECT  CP1.Compl�ment as NUMGAMME, CP2.Compl�ment as MATIERE,'' AS GAMOBS,G.GACLEUNIK,A.NAF,A.COCLI,A.PIECE,C.NOM AS NOMCLIENT,A.QTEAF, A.DESA+  A.DESA2 +A.DESA3  AS DESA1" & _
            " FROM  (((AFFAIRE  A INNER JOIN  CLIENT C   ON C.COCLI = A.COCLI )" & _
              " INNER JOIN GAMME G on  G.NAF=A.NAF ) " & _
            " LEFT JOIN COMPLEMS  CP1 on  CP1.Cl�unik=G.GACLEUNIK and CP1.COPAR='GACPL01')" & _
            " LEFT JOIN COMPLEMS  CP2 on  CP2.Cl�unik=G.GACLEUNIK and CP2.COPAR='GACPL02'" & _
            " Where g.GaCLeUnik = '" & NumPhaseClipper & "'"
                
                
                
        .CursorLocation = adUseServer
        .MaxRecords = 1
        .Open Requete, ConnexionBDNickelSQL, adOpenStatic, adLockReadOnly, adCmdText
        
        If .BOF = False And .EOF = False Then
        
            '--- pointer le premier enregistrement ---
            .MoveFirst
        
            '--- analyse apr�s recherche ---
            If .BOF = False And .EOF = False Then
                ExtraitEnrPhasesClipper Enregistrement, TTempEnrPhasesClipper
                RecherchePhasesClipper = TROUVE
            Else
                RecherchePhasesClipper = NON_TROUVE
            End If
                
        Else
            
            '--- affectation ---
            RecherchePhasesClipper = NON_TROUVE
        
        End If
       
    End With
    Screen.MousePointer = vbNormal
    '--- fermeture de l'enregistrement / connexion ---
    Enregistrement.Close
    Set Enregistrement = Nothing
    ConnexionBDNickelSQL.Close
    Set ConnexionBDNickelSQL = Nothing
    
    Exit Function

'--- gestion des erreurs ---
GestionErreurs:
    
    '--- valeur de retour ---
    RecherchePhasesClipper = NON_TROUVE
    MsgBox (CStr(Err.Description))
    
    '--- fermeture de l'enregistrement / connexion ---
    On Error Resume Next
    Enregistrement.Close
    Set Enregistrement = Nothing
    ConnexionBDNickelSQL.Close
    Set ConnexionBDNickelSQL = Nothing

End Function

'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' R�le      : Recherche le centre de frais correspondant au num�ro de poste de la ligne d'ANODISATION
' Entr�es :                                            NumPoste -> Num�ro du poste
' Retours : RechercheCentreDeFraisAnodisation -> TROUVE          = Enregistrement trouv� ou valid�
'                                                                                 NON_TROUVE = Recherche non trouv�e ou abandonn�e
'                                                                                                             autres valeurs = N� du message d'erreur
'                                                             COFRAIS -> Centre de frais
'                                                         ATTENTION -> Les caract�ristiques de l'enregistrement se trouve dans la
'                                                                                 m�moire temporaire
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function RechercheCentreDeFraisAnodisation(ByVal NumPoste As Variant, _
                                                                                      ByRef CoFrais As String) As String
    
    '--- aiguillage en cas d'erreurs ---
    On Error GoTo GestionErreurs
    
    '--- d�claration ---
    Dim Requete As String
    Dim ConnexionBDAnodisationSQL As New ADODB.Connection
    Dim Enregistrement As New ADODB.Recordset

    '--- contr�le ---
    If IsNumeric(NumPoste) = False Then
        RechercheCentreDeFraisAnodisation = NON_TROUVE
        Exit Function
    End If
    
    '--- ouverture de la connexion � la base de donn�es ANODISATION en SQL ---
    With ConnexionBDAnodisationSQL
        .ConnectionString = PARAMETRES_CONNEXION_BD_ANODISATION_SQL
        .CursorLocation = adUseServer
        .Mode = adModeReadWrite
        .ConnectionTimeout = 2     'X secondes d'attente de connexion avant de lancer un message d'erreur
        .Open
    End With
    
    '--- recherche ---
    With Enregistrement

        '--- lancement de la requ�te ---
        Requete = "SELECT CorrespondanceClipperAnodisation.* FROM CorrespondanceClipperAnodisation WHERE (NumPoste = " & NumPoste & ")"
        .CursorLocation = adUseServer
        .MaxRecords = 1
        .Open Requete, ConnexionBDAnodisationSQL, adOpenStatic, adLockReadOnly, adCmdText
        
        If .BOF = False And .EOF = False Then
        
            '--- pointer le premier enregistrement ---
            .MoveFirst
        
            '--- analyse apr�s recherche ---
            If .BOF = False And .EOF = False Then
                CoFrais = Enregistrement("COFRAIS").value
                RechercheCentreDeFraisAnodisation = TROUVE
            Else
                RechercheCentreDeFraisAnodisation = NON_TROUVE
            End If
                
        Else
            
            '--- affectation ---
            RechercheCentreDeFraisAnodisation = NON_TROUVE
        
        End If
       
    End With
    
    '--- fermeture de l'enregistrement / connexion ---
    Enregistrement.Close
    Set Enregistrement = Nothing
    ConnexionBDAnodisationSQL.Close
    Set ConnexionBDAnodisationSQL = Nothing
    
    Exit Function

'--- gestion des erreurs ---
GestionErreurs:
    
    '--- valeur de retour ---
    RechercheCentreDeFraisAnodisation = CStr(Err.Number)
    
    '--- fermeture de l'enregistrement / connexion ---
    On Error Resume Next
    Enregistrement.Close
    Set Enregistrement = Nothing
    ConnexionBDAnodisationSQL.Close
    Set ConnexionBDAnodisationSQL = Nothing

End Function



Attribute VB_Name = "MBaseDeDonnees"
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' R�le                    : MODULE DE GESTION DE LA BASE DE DONNEES
' Nom                    : MBaseDeDonnees.bas
' Date de cr�ation : 26/03/1999
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

'--- d�clarations obligatoires ---
Option Explicit

'--- options g�n�rales ---
Option Base 1
DefVar A-Z
Declare Function GetCurrentProcessId Lib "kernel32" () As Long


Public mcInsertClipper As Object


Public mlID As Long
'--- constantes priv�es ---
Public Const TROUVE As String = "TROUVE"                         'r�ponses pour les recherches
Public Const ABANDON As String = "ABANDON"                   'r�ponses pour les recherches
Public Const NON_TROUVE As String = "NON TROUVE"       'r�ponses pour les recherches

'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' R�le      : Extrait un enregistrement de la table des d�tails des gammes d'anodisation
' Entr�es :                             Enregistrement -> Enregistrement de la table des d�tails des gammes d'anodisation
' Retours : TEnrDetailsGammesAnodisation -> Tableau contenant l'image d'un enregistrement de la table
'                                                                          des d�tails des gammes d'anodisation
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Sub ExtraitEnrDetailsGammesAnodisation(ByVal Enregistrement As ADODB.Recordset, _
                                                                                ByRef TEnrDetailsGammesAnodisation As EnrDetailsGammesAnodisation)
    
    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next

    With TEnrDetailsGammesAnodisation
            
        '--- extraction de l'enregistrement ---
        .NumLigne = C_Nullite_Champ(Enregistrement, "NumLigne", 0)
        .NumZone = C_Nullite_Champ(Enregistrement, "NumZone", 0)
        
        .TempsAuPosteTexte = C_Nullite_Champ(Enregistrement, "TempsAuPosteTexte", "00:00:00")
        .TempsAlerteTexte = C_Nullite_Champ(Enregistrement, "TempsAlerteTexte", "00:00:00")
        .TempsEgouttageTexte = C_Nullite_Champ(Enregistrement, "TempsEgouttageTexte", "00:00")
        
        .TempsAuPosteSecondes = C_Nullite_Champ(Enregistrement, "TempsAuPosteSecondes", 0)
        .TempsAlerteSecondes = C_Nullite_Champ(Enregistrement, "TempsAlerteSecondes", 0)
        .TempsEgouttageSecondes = C_Nullite_Champ(Enregistrement, "TempsEgouttageSecondes", 0)

        '********** UTILISER UNIQUEMENT EN PRODUCTION **********
        .NumPosteReel = 0
        .DecompteDuTempsAuPosteReelSecondes = ""                'chaine vide pour indiquer un d�compte
                                                                                                     'non commencer
        .FinDuTempsPosteReel = False
    
    End With
    
End Sub

'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' R�le      : Extrait un enregistrement de la table des d�tails des charges de production
' Entr�es :                          Enregistrement -> Enregistrement de la table des d�tails des charges de production
' Retours : TEnrDetailsChargesproduction -> Tableau contenant l'image d'un enregistrement de la table des d�tails
'                                                                       des charges de production
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Sub ExtraitEnrDetailsChargesProduction(ByVal Enregistrement As ADODB.Recordset, _
                                                                             ByRef TEnrDetailschargesProduction As EnrDetailsChargesProduction)
    
    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next
    
    '--- d�claration ---
    Dim a As Integer                    'pour les boucles FOR...NEXT
    
    With TEnrDetailschargesProduction
            
        '--- extraction de l'enregistrement ---
        .NumCommandeInterne = C_Nullite_Champ(Enregistrement, "NumCommandeInterne", "")
        .DateEntreeEnLigne = C_Nullite_Champ(Enregistrement, "DateEntreeEnLigne", Empty)
        .DateArriveeAuDechargement = C_Nullite_Champ(Enregistrement, "DateArriveeAuDechargement", Empty)
        .NumBarre = C_Nullite_Champ(Enregistrement, "NumBarre", 0)
        .NumLigne = C_Nullite_Champ(Enregistrement, "NumLigne", 0)
        .CodeClient = C_Nullite_Champ(Enregistrement, "CodeClient", "")
        .NbrPieces = C_Nullite_Champ(Enregistrement, "NbrPieces", 0)
        .Designation = C_Nullite_Champ(Enregistrement, "Designation", "")
        .Matiere = C_Nullite_Champ(Enregistrement, "Matiere", "")
        .NumLignesReferencesClient = C_Nullite_Champ(Enregistrement, "NumLignesReferencesClient", "")
        .NumGammeAnodisation = C_Nullite_Champ(Enregistrement, "NumGammeAnodisation", "")
        .RefGammeAnodisation = C_Nullite_Champ(Enregistrement, "RefGammeAnodisation", "")
        .NumFicheProduction = C_Nullite_Champ(Enregistrement, "NumFicheProduction", "")
        .ChargePrioritaire = C_Nullite_Champ(Enregistrement, "ChargePrioritaire", False)
        .AlarmesLigne = C_Nullite_Champ(Enregistrement, "AlarmesLigne", "")
        .ControleColmatage = C_Nullite_Champ(Enregistrement, "ControleColmatage", 0)
        .ControleEpaisseurAnodisation = C_Nullite_Champ(Enregistrement, "ControleEpaisseurAnodisation", 0)
        .ControleColoration = C_Nullite_Champ(Enregistrement, "ControleColoration", "")
        .ControleObservations = C_Nullite_Champ(Enregistrement, "ControleObservations", "")
    
    End With
    
End Sub

'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' R�le      : Extrait un enregistrement de la table des d�tails des fiches de production
' Entr�es :                       Enregistrement -> Enregistrement de la table des d�tails des fiches de production
' Retours : TEnrDetailsFichesproduction -> Tableau contenant l'image d'un enregistrement de la table des d�tails
'                                                                    des fiches de production
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Sub ExtraitEnrDetailsFichesProduction(ByVal Enregistrement As ADODB.Recordset, _
                                                                          ByRef TEnrDetailsFichesProduction As EnrDetailsFichesProduction)
    
    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next

    With TEnrDetailsFichesProduction
            
        '--- extraction de l'enregistrement ---
        .NumFicheProduction = C_Nullite_Champ(Enregistrement, "NumFicheProduction", "")
        .NumLigne = C_Nullite_Champ(Enregistrement, "NumLigne", 0)
        .NumPoste = C_Nullite_Champ(Enregistrement, "NumPoste", 0)
        .DateEntreePoste = C_Nullite_Champ(Enregistrement, "DateEntreePoste", Empty)
        .DateSortiePoste = C_Nullite_Champ(Enregistrement, "DateSortiePoste", Empty)
        .DateDebutEgouttage = C_Nullite_Champ(Enregistrement, "DateDebutEgouttage", Empty)
        .DateFinEgouttage = C_Nullite_Champ(Enregistrement, "DateFinEgouttage", Empty)
        .TemperatureEnEntree = C_Nullite_Champ(Enregistrement, "TemperatureEnEntree", 0)
        .TemperatureEnSortie = C_Nullite_Champ(Enregistrement, "TemperatureEnSortie", 0)
        .GrapheTemperature = C_Nullite_Champ(Enregistrement, "GrapheTemperature", "")
        .URedresseur = C_Nullite_Champ(Enregistrement, "URedresseur", 0)
        .IRedresseur = C_Nullite_Champ(Enregistrement, "IRedresseur", 0)
        .SensRedresseur = C_Nullite_Champ(Enregistrement, "SensRedresseur", 0)
        .GrapheRedresseur = C_Nullite_Champ(Enregistrement, "GrapheRedresseur", "")
        .AnalyseurEnEntree = C_Nullite_Champ(Enregistrement, "AnalyseurEnEntree", 0)
        .AnalyseurEnSortie = C_Nullite_Champ(Enregistrement, "AnalyseurEnSortie", 0)
        .GrapheAnalyseur = C_Nullite_Champ(Enregistrement, "GrapheAnalyseur", "")
        .AlarmesPoste = C_Nullite_Champ(Enregistrement, "AlarmesPoste", "")
    
    End With
    
End Sub

'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' R�le      : Extrait un enregistrement de la table des d�tails des gammes de production
' Entr�es :                           Enregistrement -> Enregistrement de la table des d�tails des gammes de production
' Retours : TEnrDetailsGammesProduction -> Tableau contenant l'image d'un enregistrement de la table
'                                                                         des d�tails des gammes de production
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Sub ExtraitEnrDetailsGammesProduction(ByVal Enregistrement As ADODB.Recordset, _
                                                                              ByRef TEnrDetailsGammesProduction As EnrDetailsGammesProduction)
    
    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next

    With TEnrDetailsGammesProduction
            
        '--- extraction de l'enregistrement ---
        .NumFicheProduction = C_Nullite_Champ(Enregistrement, "NumFicheProduction", "")
        .NumLigne = C_Nullite_Champ(Enregistrement, "NumLigne", 0)
        .NumZone = C_Nullite_Champ(Enregistrement, "NumZone", 0)
        .TempsAuPosteTexte = C_Nullite_Champ(Enregistrement, "TempsAuPosteTexte", "")
        .TempsEgouttageTexte = C_Nullite_Champ(Enregistrement, "TempsEgouttageTexte", "")
        .TempsAuPosteSecondes = C_Nullite_Champ(Enregistrement, "TempsAuPosteSecondes", 0)
        .TempsEgouttageSecondes = C_Nullite_Champ(Enregistrement, "TempsEgouttageSecondes", 0)
        .DecompteDuTempsAuPosteReelSecondes = C_Nullite_Champ(Enregistrement, "DecompteDuTempsAuPosteReelSecondes", "")
        .NumPosteReel = C_Nullite_Champ(Enregistrement, "NumPosteReel", 0)
    
    End With
    
End Sub

'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' R�le      : Extrait un enregistrement de la table des phases de production
' Entr�es :                         Enregistrement -> Enregistrement de la table des d�tails des phases de production
' Retours : TEnrDetailsPhasesProduction -> Tableau contenant l'image d'un enregistrement de la table des phases
'                                                                      de production
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Sub ExtraitEnrDetailsPhasesProduction(ByVal Enregistrement As ADODB.Recordset, _
                                                                            ByRef TEnrDetailsPhasesProduction As EnrDetailsPhasesProduction)
    
    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next

    With TEnrDetailsPhasesProduction
            
        '--- extraction de l'enregistrement ---
        .NumFicheProduction = C_Nullite_Champ(Enregistrement, "NumFicheProduction", "")
        .NumRedresseur = C_Nullite_Champ(Enregistrement, "NumRedresseur", 0)
        .ModeUouI = C_Nullite_Champ(Enregistrement, "ModeUouI", 0)
        .NumPhase = C_Nullite_Champ(Enregistrement, "NumPhase", 0)
        .TempsPhase = C_Nullite_Champ(Enregistrement, "TempsPhase", 0)
        .UPhase = C_Nullite_Champ(Enregistrement, "UPhase", 0)
        .IPhase = C_Nullite_Champ(Enregistrement, "IPhase", 0)
    
    End With
    
End Sub

'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' R�le      : Extrait un enregistrement de la table des gammes d'anodisation
' Entr�es :                  Enregistrement -> Enregistrement de la table des gammes d'anodisation
' Retours : TEnrGammesAnodisation -> Tableau contenant l'image d'un enregistrement de la table
'                                                               des gammes d'anodisation
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Sub ExtraitEnrGammesAnodisation(ByVal Enregistrement As ADODB.Recordset, _
                                                                     ByRef TEnrGammesAnodisation As EnrGammesAnodisation)
    
    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next

    '--- d�claration ---
    Dim a As Integer           'pour les boucles FOR...NEXT

    With TEnrGammesAnodisation
            
        '--- extraction de l'enregistrement ---
        .NumGamme = C_Nullite_Champ(Enregistrement, "NumGamme", FORMAT_NUM_GAMME_ANODISATION)
        .DateCreationGamme = C_Nullite_Champ(Enregistrement, "DateCreationGamme", Empty)
        .RefGamme = C_Nullite_Champ(Enregistrement, "RefGamme", "")
        .NomGamme = C_Nullite_Champ(Enregistrement, "NomGamme", "")
        .Designation = C_Nullite_Champ(Enregistrement, "Designation", "")
        
        For a = 1 To UBound(.TMatieresGamme())
            .TMatieresGamme(a) = C_Nullite_Champ(Enregistrement, "Matiere" & a, "")
        Next a
        
        .TempsAvantPostePrincipalTexte = C_Nullite_Champ(Enregistrement, "TempsAvantPostePrincipalTexte", "")
        .TempsPostePrincipalTexte = C_Nullite_Champ(Enregistrement, "TempsPostePrincipalTexte", "")
        .TempsApresPostePrincipalTexte = C_Nullite_Champ(Enregistrement, "TempsApresPostePrincipalTexte", "")
        .TempsTotalPostesTexte = C_Nullite_Champ(Enregistrement, "TempsTotalPostesTexte", "")
        .TempsTotalEgouttagesTexte = C_Nullite_Champ(Enregistrement, "TempsTotalEgouttagesTexte", "")
        .TempsTotalGammeTexte = C_Nullite_Champ(Enregistrement, "TempsTotalGammeTexte", "")
        .TempsAvantPostePrincipalSecondes = C_Nullite_Champ(Enregistrement, "TempsAvantPostePrincipalSecondes", 0)
        .TempsPostePrincipalSecondes = C_Nullite_Champ(Enregistrement, "TempsPostePrincipalSecondes", 0)
        .TempsApresPostePrincipalSecondes = C_Nullite_Champ(Enregistrement, "TempsApresPostePrincipalSecondes", 0)
        .TempsTotalPostesSecondes = C_Nullite_Champ(Enregistrement, "TempsTotalPostesSecondes", 0)
        .TempsTotalEgouttagesSecondes = C_Nullite_Champ(Enregistrement, "TempsTotalEgouttagesSecondes", 0)
        .TempsTotalGammeSecondes = C_Nullite_Champ(Enregistrement, "TempsTotalGammeSecondes", 0)
    
        .PassageAnodisation = C_Nullite_Champ(Enregistrement, "PassageAnodisation", False)
        .PassageSpectro = C_Nullite_Champ(Enregistrement, "PassageSpectro", False)
        .PassageOr = C_Nullite_Champ(Enregistrement, "PassageOr", False)
        .PassageNoir = C_Nullite_Champ(Enregistrement, "PassageNoir", False)
        
        .ModeUouI = C_Nullite_Champ(Enregistrement, "ModeUouI", 0)
        
        For a = LBound(.TDetailsPhases()) To UBound(.TDetailsPhases())
            With .TDetailsPhases(a)
                .TempsPhase = C_Nullite_Champ(Enregistrement, "TempsPhase" & a, 0)
                .UPhase = C_Nullite_Champ(Enregistrement, "UPhase" & a, 0)
                .IPhase = C_Nullite_Champ(Enregistrement, "IPhase" & a, 0)
            End With
        Next a
        
        Erase .TDetailsGammesAnodisation()
    
        '********** UTILISER UNIQUEMENT EN PRODUCTION **********
        .ChoixPosteAnodisation = CHOIX_POSTE_ANODISATION.C_AUTOMATIQUE
    
    End With
    
End Sub

'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' R�le      : Recherche les d�tails des gammes d'anodisation
' Entr�es :                                  NumGamme -> Num�ro de la gamme
' Retours : RechercheDetailsGammesAnodisation -> TROUVE           = Enregistrement(s) trouv� ou valid�
'                                                                           NON_TROUVE = Recherche non trouv�e ou abandonn�e
'                                                                                                       autres valeurs = N� du message d'erreur
'                                                   ATTENTION -> Les caract�ristiques de l'enregistrement se trouve dans la
'                                                                           m�moire temporaire
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function RechercheDetailsGammesAnodisation(ByVal NumGamme As String) As String
    
    '--- aiguillage en cas d'erreurs ---
    On Error GoTo GestionErreurs
    
    '--- constantes priv�es ---
    
    '--- d�claration ---
    Dim CptEnr As Long
    Dim Requete As String
    Dim ConnexionBDAnodisationSQL As New ADODB.Connection
    Dim Enregistrement As New ADODB.Recordset

    '--- contr�le ---
    If NumGamme = "" Then
        RechercheDetailsGammesAnodisation = NON_TROUVE
        Exit Function
    End If
    
    '--- red�claration ---
    ReDim TTempEnrDetailsGammesAnodisation(1 To 1) As EnrDetailsGammesAnodisation
    
    '--- affectation ---
    CptEnr = 1
    
    '--- ouverture de la connexion � la base de donn�es de l'anodisation en SQL SERVER 2000 ---
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
        Requete = "SELECT DetailsGammesAnodisation.* FROM DetailsGammesAnodisation WHERE (NumGamme = '" & NumGamme & "') ORDER BY NumLigne"
        .CursorLocation = adUseServer
        .MaxRecords = NBR_LIGNES_DETAILS_GAMMES_PRODUCTION
        .Open Requete, ConnexionBDAnodisationSQL, adOpenStatic, adLockOptimistic, adCmdText
        
        If .BOF = False And .EOF = False Then
        
            '--- pointer le premier enregistrement ---
            .MoveFirst
            
            '--- extraction du premier enregistrement ---
            ExtraitEnrDetailsGammesAnodisation Enregistrement, TTempEnrDetailsGammesAnodisation(CptEnr)
    
            '--- affectation ---
            RechercheDetailsGammesAnodisation = TROUVE
        
            '--- recherche des enregistrements suivants ---
            Do
         
                '--- passage � l'enregistrement suivant ---
                .MoveNext
                If .BOF = True Or .EOF = True Then Exit Do
                
                '--- incr�mentation ---
                Inc CptEnr
         
                '--- extraction ---
                ReDim Preserve TTempEnrDetailsGammesAnodisation(1 To CptEnr) As EnrDetailsGammesAnodisation
                ExtraitEnrDetailsGammesAnodisation Enregistrement, TTempEnrDetailsGammesAnodisation(CptEnr)
        
            Loop
                
        Else
            
            '--- affectation ---
            RechercheDetailsGammesAnodisation = NON_TROUVE
        
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
    RechercheDetailsGammesAnodisation = CStr(Err.Number)

    '--- fermeture de l'enregistrement / connexion ---
    On Error Resume Next
    Enregistrement.Close
    Set Enregistrement = Nothing
    ConnexionBDAnodisationSQL.Close
    Set ConnexionBDAnodisationSQL = Nothing

End Function

'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' R�le      : Recherche les d�tails des fiches de production
' Entr�es :                        NumFicheProduction -> Num�ro de la fiche de production
' Retours : RechercheDetailsFichesProduction -> TROUVE           = Enregistrement(s) trouv� ou valid�
'                                                                               NON_TROUVE = Recherche non trouv�e ou abandonn�e
'                                                                                                          autres valeurs = N� du message d'erreur
'                                                       ATTENTION -> Les caract�ristiques de l'enregistrement se trouve dans la
'                                                                               m�moire temporaire
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function RechercheDetailsFichesProduction(ByVal NumFicheProduction As String) As String
    
    '--- aiguillage en cas d'erreurs ---
    On Error GoTo GestionErreurs
    
    '--- d�claration ---
    Dim CptEnr As Long
    Dim Requete As String
    Dim ConnexionBDAnodisationSQL As New ADODB.Connection
    Dim Enregistrement As New ADODB.Recordset

    '--- contr�le ---
    If NumFicheProduction = "" Then
        RechercheDetailsFichesProduction = NON_TROUVE
        Exit Function
    End If
    
    '--- red�claration ---
    ReDim TTempEnrDetailsFichesProduction(1 To 1) As EnrDetailsFichesProduction
    
    '--- affectation ---
    CptEnr = 1
    
    '--- ouverture de la connexion � la base de donn�es de l'anodisation en SQL SERVER 2000 ---
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
        Requete = "SELECT DetailsFichesProduction.* FROM DetailsFichesProduction WHERE (NumFicheProduction = '" & NumFicheProduction & "') ORDER BY NumLigne"
        .CursorLocation = adUseServer
        .MaxRecords = NBR_LIGNES_DETAILS_CHARGES
        .Open Requete, ConnexionBDAnodisationSQL, adOpenStatic, adLockOptimistic, adCmdText
        
        If .BOF = False And .EOF = False Then
        
            '--- pointer le premier enregistrement ---
            .MoveFirst
            
            '--- extraction du premier enregistrement ---
            ExtraitEnrDetailsFichesProduction Enregistrement, TTempEnrDetailsFichesProduction(CptEnr)
    
            '--- affectation ---
            RechercheDetailsFichesProduction = TROUVE
        
            '--- recherche des enregistrements suivants ---
            Do
         
                '--- passage � l'enregistrement suivant ---
                .MoveNext
                If .BOF = True Or .EOF = True Then Exit Do
                
                '--- incr�mentation ---
                Inc CptEnr
         
                '--- extraction ---
                ReDim Preserve TTempEnrDetailsFichesProduction(1 To CptEnr) As EnrDetailsFichesProduction
                ExtraitEnrDetailsFichesProduction Enregistrement, TTempEnrDetailsFichesProduction(CptEnr)
        
            Loop
                
        Else
            
            '--- affectation ---
            RechercheDetailsFichesProduction = NON_TROUVE
        
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
    RechercheDetailsFichesProduction = CStr(Err.Number)

    '--- fermeture de l'enregistrement / connexion ---
    On Error Resume Next
    Enregistrement.Close
    Set Enregistrement = Nothing
    ConnexionBDAnodisationSQL.Close
    Set ConnexionBDAnodisationSQL = Nothing

End Function

'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' R�le      : Recherche les d�tails des charges de production
' Entr�es :                           NumFicheProduction -> Num�ro de la fiche de production
' Retours : RechercheDetailsChargesProduction -> TROUVE           = Enregistrement(s) trouv� ou valid�
'                                                                                  NON_TROUVE = Recherche non trouv�e ou abandonn�e
'                                                                                                              autres valeurs = N� du message d'erreur
'                                                          ATTENTION -> Les caract�ristiques de l'enregistrement se trouve dans la
'                                                                                  m�moire temporaire
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function RechercheDetailsChargesProduction(ByVal NumFicheProduction As String) As String
    
    '--- aiguillage en cas d'erreurs ---
    On Error GoTo GestionErreurs
    
    '--- d�claration ---
    Dim CptEnr As Long
    Dim Requete As String
    Dim ConnexionBDAnodisationSQL As New ADODB.Connection
    Dim Enregistrement As New ADODB.Recordset

    '--- contr�le ---
    If NumFicheProduction = "" Then
        RechercheDetailsChargesProduction = NON_TROUVE
        Exit Function
    End If
    
    '--- red�claration ---
    ReDim TTempEnrDetailsChargesProduction(1 To 1) As EnrDetailsChargesProduction
    
    '--- affectation ---
    CptEnr = 1
    
    '--- ouverture de la connexion � la base de donn�es de l'anodisation en SQL SERVER 2000 ---
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
        Requete = "SELECT DetailsChargesProduction.* FROM DetailsChargesProduction WHERE (NumFicheProduction = '" & NumFicheProduction & "') ORDER BY NumLigne"
        .CursorLocation = adUseServer
        .MaxRecords = NBR_LIGNES_DETAILS_CHARGES
        .Open Requete, ConnexionBDAnodisationSQL, adOpenStatic, adLockOptimistic, adCmdText
        
        If .BOF = False And .EOF = False Then
        
            '--- pointer le premier enregistrement ---
            .MoveFirst
            
            '--- extraction du premier enregistrement ---
            ExtraitEnrDetailsChargesProduction Enregistrement, TTempEnrDetailsChargesProduction(CptEnr)
    
            '--- affectation ---
            RechercheDetailsChargesProduction = TROUVE
        
            '--- recherche des enregistrements suivants ---
            Do
         
                '--- passage � l'enregistrement suivant ---
                .MoveNext
                If .BOF = True Or .EOF = True Then Exit Do
                
                '--- incr�mentation ---
                Inc CptEnr
         
                '--- extraction ---
                ReDim Preserve TTempEnrDetailsChargesProduction(1 To CptEnr) As EnrDetailsChargesProduction
                ExtraitEnrDetailsChargesProduction Enregistrement, TTempEnrDetailsChargesProduction(CptEnr)
        
            Loop
                
        Else
            
            '--- affectation ---
            RechercheDetailsChargesProduction = NON_TROUVE
        
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
    RechercheDetailsChargesProduction = CStr(Err.Number)

    '--- fermeture de l'enregistrement / connexion ---
    On Error Resume Next
    Enregistrement.Close
    Set Enregistrement = Nothing
    ConnexionBDAnodisationSQL.Close
    Set ConnexionBDAnodisationSQL = Nothing
    
End Function

'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' R�le      : Recherche les d�tails des gammes d'anodisation de production
' Entr�es :                                      NumFicheProduction -> Num�ro de la fiche de production
' Retours : RechercheDetailsGammesProduction -> TROUVE           = Enregistrement(s) trouv� ou valid�
'                                                                                             NON_TROUVE = Recherche non trouv�e ou abandonn�e
'                                                                                                                        autres valeurs = N� du message d'erreur
'                                                                    ATTENTION -> Les caract�ristiques de l'enregistrement se trouve
'                                                                                            dans la m�moire temporaire
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function RechercheDetailsGammesProduction(ByVal NumFicheProduction As String) As String
    
    '--- aiguillage en cas d'erreurs ---
    On Error GoTo GestionErreurs
    
    '--- d�claration ---
    Dim CptEnr As Long
    Dim Requete As String
    Dim ConnexionBDAnodisationSQL As New ADODB.Connection
    Dim Enregistrement As New ADODB.Recordset

    '--- contr�le ---
    If NumFicheProduction = "" Then
        RechercheDetailsGammesProduction = NON_TROUVE
        Exit Function
    End If
    
    '--- red�claration ---
    ReDim TTempEnrDetailsGammesProduction(1 To 1) As EnrDetailsGammesProduction
    
    '--- affectation ---
    CptEnr = 1
    
    '--- ouverture de la connexion � la base de donn�es de l'anodisation en SQL SERVER 2000 ---
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
        Requete = "SELECT DetailsGammesProduction.* FROM DetailsGammesProduction WHERE (NumFicheProduction = '" & NumFicheProduction & "') ORDER BY NumLigne"
        .CursorLocation = adUseServer
        .MaxRecords = NBR_LIGNES_DETAILS_GAMMES_PRODUCTION
        .Open Requete, ConnexionBDAnodisationSQL, adOpenStatic, adLockOptimistic, adCmdText
        
        If .BOF = False And .EOF = False Then
        
            '--- pointer le premier enregistrement ---
            .MoveFirst
            
            '--- extraction du premier enregistrement ---
            ExtraitEnrDetailsGammesProduction Enregistrement, TTempEnrDetailsGammesProduction(CptEnr)
    
            '--- affectation ---
            RechercheDetailsGammesProduction = TROUVE
        
            '--- recherche des enregistrements suivants ---
            Do
         
                '--- passage � l'enregistrement suivant ---
                .MoveNext
                If .BOF = True Or .EOF = True Then Exit Do
                
                '--- incr�mentation ---
                Inc CptEnr
         
                '--- extraction ---
                ReDim Preserve TTempEnrDetailsGammesProduction(1 To CptEnr) As EnrDetailsGammesProduction
                ExtraitEnrDetailsGammesProduction Enregistrement, TTempEnrDetailsGammesProduction(CptEnr)
        
            Loop
                
        Else
            
            '--- affectation ---
            RechercheDetailsGammesProduction = NON_TROUVE
        
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
    RechercheDetailsGammesProduction = CStr(Err.Number)

    '--- fermeture de l'enregistrement / connexion ---
    On Error Resume Next
    Enregistrement.Close
    Set Enregistrement = Nothing
    ConnexionBDAnodisationSQL.Close
    Set ConnexionBDAnodisationSQL = Nothing

End Function

'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' R�le      : Recherche les d�tails des tensions et intensit�s de production
' Entr�es :                 NumFicheProduction -> Num�ro de la fiche de production
' Retours : RechercheDetailsPhasesProduction -> TROUVE           = Enregistrement(s) trouv� ou valid�
'                                                                        NON_TROUVE = Recherche non trouv�e ou abandonn�e
'                                                                                                   autres valeurs = N� du message d'erreur
'                                                                        ATTENTION -> Les caract�ristiques de l'enregistrement se trouve
'                                                                                                dans la m�moire temporaire
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function RechercheDetailsPhasesProduction(ByVal NumFicheProduction As String) As String
    
    '--- aiguillage en cas d'erreurs ---
    On Error GoTo GestionErreurs
    
    '--- d�claration ---
    Dim CptEnr As Long
    Dim Requete As String
    Dim ConnexionBDAnodisationSQL As New ADODB.Connection
    Dim Enregistrement As New ADODB.Recordset

    '--- contr�le ---
    If NumFicheProduction = "" Then
        RechercheDetailsPhasesProduction = NON_TROUVE
        Exit Function
    End If
    
    '--- red�claration ---
    ReDim TTempEnrDetailsPhasesProduction(1 To 1) As EnrDetailsPhasesProduction
    
    '--- affectation ---
    CptEnr = 1
    
    '--- ouverture de la connexion � la base de donn�es de l'anodisation en SQL SERVER 2000 ---
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
        Requete = "SELECT DetailsPhasesProduction.* FROM DetailsPhasesProduction WHERE (NumFicheProduction = '" & NumFicheProduction & "') ORDER BY NumPhase"
        .CursorLocation = adUseServer
        .MaxRecords = 0
        .Open Requete, ConnexionBDAnodisationSQL, adOpenStatic, adLockOptimistic, adCmdText
        
        If .BOF = False And .EOF = False Then
        
            '--- pointer le premier enregistrement ---
            .MoveFirst
            
            '--- extraction du premier enregistrement ---
            ExtraitEnrDetailsPhasesProduction Enregistrement, TTempEnrDetailsPhasesProduction(CptEnr)
    
            '--- affectation ---
            RechercheDetailsPhasesProduction = TROUVE
        
            '--- recherche des enregistrements suivants ---
            Do
         
                '--- passage � l'enregistrement suivant ---
                .MoveNext
                If .BOF = True Or .EOF = True Then Exit Do
                
                '--- incr�mentation ---
                Inc CptEnr
         
                '--- extraction ---
                ReDim Preserve TTempEnrDetailsPhasesProduction(1 To CptEnr) As EnrDetailsPhasesProduction
                ExtraitEnrDetailsPhasesProduction Enregistrement, TTempEnrDetailsPhasesProduction(CptEnr)
        
            Loop
                
        Else
            
            '--- affectation ---
            RechercheDetailsPhasesProduction = NON_TROUVE
        
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
    RechercheDetailsPhasesProduction = CStr(Err.Number)

    '--- fermeture de l'enregistrement / connexion ---
    On Error Resume Next
    Enregistrement.Close
    Set Enregistrement = Nothing
    ConnexionBDAnodisationSQL.Close
    Set ConnexionBDAnodisationSQL = Nothing
    
End Function

'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' R�le      : Recherche des gammes d'anodisation compl�te
' Entr�es :                       NumGamme -> Num�ro de la gamme
' Retours : RechercheGammesAnodisation -> TROUVE           = Enregistrement(s) trouv� ou valid�
'                                                                NON_TROUVE = Recherche non trouv�e ou abandonn�e
'                                                                                           autres valeurs = N� du message d'erreur
'                                        ATTENTION -> Les caract�ristiques de l'enregistrement se trouve dans la
'                                                                m�moire temporaire
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function RechercheGammesAnodisation(ByVal NumGamme As String) As String
    
    '--- aiguillage en cas d'erreurs ---
    On Error GoTo GestionErreurs
    
    '--- d�claration ---
    Dim a As Integer
    Dim Requete As String
    Dim ConnexionBDAnodisationSQL As New ADODB.Connection
    Dim Enregistrement As New ADODB.Recordset

    '--- contr�le ---
    If NumGamme = "" Then
        RechercheGammesAnodisation = NON_TROUVE
        Exit Function
    End If
    
    '--- ouverture de la connexion � la base de donn�es de l'anodisation en SQL SERVER 2000 ---
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
        Requete = "SELECT GammesAnodisation.* FROM GammesAnodisation WHERE (NumGamme = '" & NumGamme & "')"
        .CursorLocation = adUseServer
        .MaxRecords = 1
        .Open Requete, ConnexionBDAnodisationSQL, adOpenStatic, adLockReadOnly, adCmdText
        
        If .BOF = False And .EOF = False Then
        
            '--- pointer le premier enregistrement ---
            .MoveFirst
        
            '--- analyse apr�s recherche ---
            If .BOF = False And .EOF = False Then
                ExtraitEnrGammesAnodisation Enregistrement, TTempEnrGammesAnodisation
                
                '--- recherche des d�tails ---
                If RechercheDetailsGammesAnodisation(NumGamme) = TROUVE Then
                    
                    With TTempEnrGammesAnodisation
                        For a = LBound(TTempEnrDetailsGammesAnodisation()) To UBound(TTempEnrDetailsGammesAnodisation())
                            .TDetailsGammesAnodisation(a) = TTempEnrDetailsGammesAnodisation(a)
                        Next a
                    End With
                
                    '--- affectation ---
                    RechercheGammesAnodisation = TROUVE
            
                Else
                
                    '--- affectation ---
                    RechercheGammesAnodisation = NON_TROUVE
            
                End If
            
            Else
                RechercheGammesAnodisation = NON_TROUVE
            End If
                
        Else
            
            '--- affectation ---
            RechercheGammesAnodisation = NON_TROUVE
        
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
    RechercheGammesAnodisation = CStr(Err.Number)
    
    '--- fermeture de l'enregistrement / connexion ---
    On Error Resume Next
    Enregistrement.Close
    Set Enregistrement = Nothing
    ConnexionBDAnodisationSQL.Close
    Set ConnexionBDAnodisationSQL = Nothing

End Function

'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' R�le      : Chargement des postes
' D�tails  :
' Entr�es :
' Retours :
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function ChargePostes() As String
    
    '--- aiguillage en cas d'erreurs ---
    On Error GoTo GestionErreurs
  
    '--- constante priv�es ---

    '--- d�claration ---
    Dim a As Integer, _
           NumPoste As Integer
    Dim Requete As String
    Dim ConnexionBDAnodisationSQL As New ADODB.Connection
    Dim Enregistrement As New ADODB.Recordset
    
    '--- affichage du type de t�che ---
    AfficheTypeTache "Chargement des postes"
    
    '--- ouverture de la connexion � la base de donn�es de l'anodisation en SQL SERVER 2000 ---
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
        Requete = "SELECT Postes.* FROM Postes ORDER BY NumPoste"
        .CursorLocation = adUseServer
        .CacheSize = 50
        .MaxRecords = 0
        .Open Requete, ConnexionBDAnodisationSQL, adOpenStatic, adLockOptimistic, adCmdText

        '--- extraction des valeurs ---
        Do While Not .EOF
            
            '--- affectation ---
            NumPoste = ![NumPoste]
            
            '--- affectation dans le tableau ---
            With TEtatsPostes(NumPoste).DefinitionPoste
                .NumPoste = Enregistrement![NumPoste]
                .NomPoste = Enregistrement![NomPoste]
                .LibellePoste = Enregistrement![LibellePoste]
                .AvecTemps = Enregistrement![AvecTemps]
                .RespectTempsObligatoire = Enregistrement![RespectTempsObligatoire]
                .AvecEgouttage = Enregistrement![AvecEgouttage]
                .PresenceCouvercles = Enregistrement![PresenceCouvercles]
                .PresenceRedresseur = Enregistrement![PresenceRedresseur]
                .PresenceAgitationBain = Enregistrement![PresenceAgitationBain]
                .XAxePosteSynoptique = Enregistrement![XAxePosteSynoptique]
                .XAxePosteLigne = Enregistrement![XAxePosteLigne]
                .XInferieurPosteSynoptique = Enregistrement![XInferieurPosteSynoptique]
                .YInferieurPosteSynoptique = Enregistrement![YInferieurPosteSynoptique]
                .XSuperieurPosteSynoptique = Enregistrement![XSuperieurPosteSynoptique]
                .YSuperieurPosteSynoptique = Enregistrement![YSuperieurPosteSynoptique]
                .XInferieurLibellePosteSynoptique = Enregistrement![XInferieurLibellePosteSynoptique]
                .YInferieurLibellePosteSynoptique = Enregistrement![YInferieurLibellePosteSynoptique]
                .XSuperieurLibellePosteSynoptique = Enregistrement![XSuperieurLibellePosteSynoptique]
                .YSuperieurLibellePosteSynoptique = Enregistrement![YSuperieurLibellePosteSynoptique]
            End With
            
            '--- enregistrement suivant ---
            .MoveNext
        
        Loop

    End With
    
    '--- fermeture de l'enregistrement / connexion ---
    Enregistrement.Close
    Set Enregistrement = Nothing
    ConnexionBDAnodisationSQL.Close
    Set ConnexionBDAnodisationSQL = Nothing

    '--- affichage du type de t�che ---
    AfficheTypeTache ("")
  
    Exit Function

GestionErreurs:
    
    '--- valeur de retour ---
    ChargePostes = CStr(Err.Number)

    '--- fermeture de l'enregistrement / connexion ---
    On Error Resume Next
    Enregistrement.Close
    Set Enregistrement = Nothing
    ConnexionBDAnodisationSQL.Close
    Set ConnexionBDAnodisationSQL = Nothing
    
End Function

'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' R�le      : Chargement des redresseurs
' D�tails  :
' Entr�es :
' Retours :
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function ChargeRedresseurs() As String
    
    '--- aiguillage en cas d'erreurs ---
    On Local Error GoTo GestionErreurs
  
    '--- constante priv�es ---

    '--- d�claration ---
    Dim a As Integer, _
           NumRedresseur As Integer
    Dim Requete As String
    Dim ConnexionBDAnodisationSQL As New ADODB.Connection
    Dim Enregistrement As New ADODB.Recordset
        
    '--- affichage du type de t�che ---
    AfficheTypeTache "Chargement des redresseurs"
    
    '--- ouverture de la connexion � la base de donn�es de l'anodisation en SQL SERVER 2000 ---
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
        Requete = "SELECT Redresseurs.* FROM Redresseurs ORDER BY NumRedresseur"
        .CursorLocation = adUseServer
        .CacheSize = 50
        .MaxRecords = 0
        .Open Requete, ConnexionBDAnodisationSQL, adOpenStatic, adLockOptimistic, adCmdText

        '--- extraction des valeurs ---
        Do While Not .EOF
            
            '--- affectation ---
            NumRedresseur = ![NumRedresseur]
            
            '--- affectation dans le tableau ---
            With TEtatsRedresseurs(NumRedresseur).DefinitionRedresseur
                .NumRedresseur = NumRedresseur
                .NomRedresseur = Enregistrement![NomRedresseur]
                .LibelleRedresseur = Enregistrement![LibelleRedresseur]
                .UMaxiRedresseur = Enregistrement![UMaxiRedresseur]
                .IMaxiRedresseur = Enregistrement![IMaxiRedresseur]
            End With
            
            '--- enregistrement suivant ---
            .MoveNext
        
        Loop

    End With
    
    '--- fermeture de l'enregistrement / connexion ---
    Enregistrement.Close
    Set Enregistrement = Nothing
    ConnexionBDAnodisationSQL.Close
    Set ConnexionBDAnodisationSQL = Nothing

    '--- affichage du type de t�che ---
    AfficheTypeTache ("")
  
    Exit Function

GestionErreurs:
    
    '--- valeur de retour ---
    ChargeRedresseurs = CStr(Err.Number)
    
    '--- fermeture de l'enregistrement / connexion ---
    On Error Resume Next
    Enregistrement.Close
    Set Enregistrement = Nothing
    ConnexionBDAnodisationSQL.Close
    Set ConnexionBDAnodisationSQL = Nothing

End Function

'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' R�le      : Chargement des barres
' D�tails  :
' Entr�es :
' Retours :
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function ChargeBarres() As String
    
    '--- aiguillage en cas d'erreurs ---
    On Error GoTo GestionErreurs
  
    '--- constante priv�es ---
    
    '--- d�claration ---
    Dim a As Integer, _
           NumBarre As Integer
    Dim Requete As String
    Dim ConnexionBDAnodisationSQL As New ADODB.Connection
    Dim Enregistrement As New ADODB.Recordset
    
    '--- d�claration ---
 
    '--- affichage du type de t�che ---
    AfficheTypeTache "Chargement des zones"
    
    '--- ouverture de la connexion � la base de donn�es de l'anodisation en SQL SERVER 2000 ---
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
        Requete = "SELECT barres.* FROM barres ORDER BY id"
        .CursorLocation = adUseClient
        .CacheSize = 50
        .MaxRecords = 0
        .Open Requete, ConnexionBDAnodisationSQL, adOpenStatic, adLockOptimistic, adCmdText

        '--- redimensionnement du tableau ---
        Erase TBarres()
        ReDim TBarres(1 To .RecordCount) As EnrBarres
        
        '--- affectation ---
        LIMITE_BASSE_BARRES = LBound(TBarres())
        LIMITE_HAUTE_BARRES = UBound(TBarres())
        
        '--- extraction des valeurs ---
        Do While Not .EOF
            
            '--- affectation ---
            NumBarre = ![ID]
            
            '--- affectation dans le tableau ---
            With TBarres(NumBarre)
                .NumBarre = Enregistrement![ID]
                .Libelle = Enregistrement![Libelle]
            End With
            
            '--- enregistrement suivant ---
            .MoveNext
        
        Loop

    End With
    
    '--- fermeture de l'enregistrement / connexion ---
    Enregistrement.Close
    Set Enregistrement = Nothing
    ConnexionBDAnodisationSQL.Close
    Set ConnexionBDAnodisationSQL = Nothing

    '--- affichage du type de t�che ---
    AfficheTypeTache ("")
  
    Exit Function

GestionErreurs:
    
    '--- valeur de retour ---
    ChargeBarres = CStr(Err.Number)
    
    '--- fermeture de l'enregistrement / connexion ---
    On Error Resume Next
    Enregistrement.Close
    Set Enregistrement = Nothing
    ConnexionBDAnodisationSQL.Close
    Set ConnexionBDAnodisationSQL = Nothing

End Function

'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' R�le      : Chargement des zones
' D�tails  :
' Entr�es :
' Retours :
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function ChargeZones() As String
    
    '--- aiguillage en cas d'erreurs ---
    On Error GoTo GestionErreurs
  
    '--- constante priv�es ---
    
    '--- d�claration ---
    Dim a As Integer, _
           NumZone As Integer
    Dim Requete As String
    Dim ConnexionBDAnodisationSQL As New ADODB.Connection
    Dim Enregistrement As New ADODB.Recordset
    
    '--- d�claration ---
 
    '--- affichage du type de t�che ---
    AfficheTypeTache "Chargement des zones"
    
    '--- ouverture de la connexion � la base de donn�es de l'anodisation en SQL SERVER 2000 ---
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
        Requete = "SELECT Zones.* FROM Zones ORDER BY NumZone"
        .CursorLocation = adUseClient
        .CacheSize = 50
        .MaxRecords = 0
        .Open Requete, ConnexionBDAnodisationSQL, adOpenStatic, adLockOptimistic, adCmdText

        '--- redimensionnement du tableau ---
        Erase TZones()
        ReDim TZones(1 To .RecordCount) As EnrZones
        
        '--- affectation ---
        LIMITE_BASSE_ZONES = LBound(TZones())
        LIMITE_HAUTE_ZONES = UBound(TZones())
        
        '--- extraction des valeurs ---
        Do While Not .EOF
            
            '--- affectation ---
            NumZone = ![NumZone]
            
            '--- affectation dans le tableau ---
            With TZones(NumZone)
                .NumZone = NumZone
                .Codezone = Enregistrement![Codezone]
                .LibelleZone = Enregistrement![LibelleZone]
                .NumPremierPoste = Enregistrement![NumPremierPoste]
                .NomPremierPoste = Enregistrement![NomPremierPoste]
                .NumDernierPoste = Enregistrement![NumDernierPoste]
                .NomDernierPoste = Enregistrement![NomDernierPoste]
                .NbrPostes = Enregistrement![NbrPostes]
            End With
            
            '--- enregistrement suivant ---
            .MoveNext
        
        Loop

    End With
    
    '--- fermeture de l'enregistrement / connexion ---
    Enregistrement.Close
    Set Enregistrement = Nothing
    ConnexionBDAnodisationSQL.Close
    Set ConnexionBDAnodisationSQL = Nothing

    '--- affichage du type de t�che ---
    AfficheTypeTache ("")
  
    Exit Function

GestionErreurs:
    
    '--- valeur de retour ---
    ChargeZones = CStr(Err.Number)
    
    '--- fermeture de l'enregistrement / connexion ---
    On Error Resume Next
    Enregistrement.Close
    Set Enregistrement = Nothing
    ConnexionBDAnodisationSQL.Close
    Set ConnexionBDAnodisationSQL = Nothing

End Function

'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' R�le      : Chargement des temps de mouvements
' D�tails  :
' Entr�es :
' Retours :
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function ChargeTempsMouvements() As String
    
    '--- aiguillage en cas d'erreurs ---
    On Error GoTo GestionErreurs
  
    '--- constante priv�es ---

    '--- d�claration ---
    Dim a As Integer, _
           NumCuveAPI As Integer, _
           NumPont As Integer, _
           NumPosteDepart As Integer, _
           NumPosteArrivee As Integer
    Dim Requete As String
    Dim ConnexionBDAnodisationSQL As New ADODB.Connection
    Dim Enregistrement As ADODB.Recordset
    
    '--- affichage du type de t�che ---
    AfficheTypeTache "Chargement des temps de mouvements"
        
    '--- ouverture de la connexion � la base de donn�es de l'anodisation en SQL SERVER 2000 ---
    With ConnexionBDAnodisationSQL
        .ConnectionString = PARAMETRES_CONNEXION_BD_ANODISATION_SQL
        .CursorLocation = adUseServer
        .Mode = adModeReadWrite
        .ConnectionTimeout = 2     'X secondes d'attente de connexion avant de lancer un message d'erreur
        .Open
    End With
        
    '*********************************************************************************************************************
    '*                                  TEMPS DE MOUVEMENTS DES PONTS SANS LA TRANSLATION
    '*********************************************************************************************************************
    Set Enregistrement = New ADODB.Recordset
    With Enregistrement

        '--- lancement de la requ�te ---
        Requete = "SELECT TempsMouvementsPontsSansTranslation.* FROM TempsMouvementsPontsSansTranslation ORDER BY NumPont"
        .CursorLocation = adUseServer
        .CacheSize = 50
        .MaxRecords = 0
        .Open Requete, ConnexionBDAnodisationSQL, adOpenStatic, adLockOptimistic, adCmdText

        '--- extraction des valeurs ---
        Do While Not .EOF
            
            '--- affectation ---
            NumPont = ![NumPont]
            
            '--- affectation dans le tableau ---
            If NumPont >= PONTS.P_1 And NumPont <= PONTS.P_2 Then
                With TEtatsPonts(NumPont).TTempsMouvements
                    .TempsAccrochesChargeVersHaut = Enregistrement![TempsAccrochesChargeVersHaut]
                    .TempsAccrochesChargeVersBas = Enregistrement![TempsAccrochesChargeVersBas]
                    .TempsDescenteHautVersBas = Enregistrement![TempsDescenteHautVersBas]
                    .TempsDescenteIntermediaireVersBas = Enregistrement![TempsDescenteIntermediaireVersBas]
                    .TempsMonteeBasVersIntermediaire = Enregistrement![TempsMonteeBasVersIntermediaire]
                    .TempsMonteeBasVersHaut = Enregistrement![TempsMonteeBasVersHaut]
                 End With
            End If
            
            '--- enregistrement suivant ---
            .MoveNext
        
        Loop

        '--- fermeture de l'enregistrement ---
        .Close

    End With
    
    '*********************************************************************************************************************
    '*                                   TEMPS DE MOUVEMENTS DE LA TRANSLATION DES PONTS
    '*********************************************************************************************************************
    Set Enregistrement = New ADODB.Recordset
    With Enregistrement
        
        '--- lancement de la requ�te ---
        Requete = "SELECT TempsMouvementsTranslationPonts.* FROM TempsMouvementsTranslationPonts ORDER BY NumPont, NumPosteDepart, NumPosteArrivee"
        .CursorLocation = adUseServer
        .CacheSize = 50
        .MaxRecords = 0
        .Open Requete, ConnexionBDAnodisationSQL, adOpenStatic, adLockOptimistic, adCmdText

        '--- extraction des valeurs ---
        Do While Not .EOF
            
            '--- affectation ---
            NumPont = ![NumPont]
            NumPosteDepart = ![NumPosteDepart]
            NumPosteArrivee = ![NumPosteArrivee]
            
            '--- affectation dans le tableau ---
            If NumPont >= PONTS.P_1 And NumPont <= PONTS.P_2 And _
               NumPosteDepart >= POSTES.P_CHGT_1 And NumPosteDepart <= DERNIER_POSTE And _
               NumPosteArrivee >= POSTES.P_CHGT_1 And NumPosteArrivee <= DERNIER_POSTE Then
                    With TEtatsPonts(NumPont).TTempsMouvements
                        .TTempsTranslation(NumPosteDepart, NumPosteArrivee) = Enregistrement![TempsTranslation]
                     End With
            End If
            
            '--- enregistrement suivant ---
            .MoveNext
        
        Loop

    End With
    
    '--- fermeture de l'enregistrement / connexion ---
    Enregistrement.Close
    Set Enregistrement = Nothing
    ConnexionBDAnodisationSQL.Close
    Set ConnexionBDAnodisationSQL = Nothing

    '--- affichage du type de t�che ---
    AfficheTypeTache ("")
  
    Exit Function

GestionErreurs:
    
    '--- valeur de retour ---
    ChargeTempsMouvements = CStr(Err.Number)
    
    '--- fermeture de l'enregistrement / connexion ---
    On Error Resume Next
    Enregistrement.Close
    Set Enregistrement = Nothing
    ConnexionBDAnodisationSQL.Close
    Set ConnexionBDAnodisationSQL = Nothing
    
End Function

'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' R�le      : Chargement des temps de mouvements
' Entr�es :
' Retours :
' D�tails  : l'enregistrement du type pr�misses est compos� comme suit (voir base de donn�es SQL) :
'
'                       ClePrimaire            : int                   'cl� primaire
'                       NumPont                : smallint          'n� du pont concern� d�fini comme r�gle au d�part
'                       NumPosteDepart   : smallint          'n� poste de d�but
'                       NumPosteArrivee  : smallint          'n� poste d'arriv�e
'                       PremisseCodee     : varchar           'pr�misse cod�e
'                       PremisseDecodee : varchar           'pr�misse d�cod�e
'
' ATTENTION : Les temps de mouvements n�cessaires aux calculs du temps de cycle de la pr�misse
'                       doivent �tre charg�s avant l'appel de cette routine
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function ChargePremisses() As String
    
    '--- aiguillage en cas d'erreurs ---
    On Error GoTo GestionErreurs
  
    '--- constante priv�es ---

    '--- d�claration ---
    Dim a As Integer, _
           NumPosteDepart As Integer, _
           NumPosteArrivee As Integer
    Dim TempsCycleSecondes As Long
    Dim Requete As String
    Dim ConnexionBDAnodisationSQL As New ADODB.Connection
    Dim Enregistrement As New ADODB.Recordset
    
    '--- affichage du type de t�che ---
    AfficheTypeTache "Chargement des pr�misses"
    
    '--- ouverture de la connexion � la base de donn�es de l'anodisation en SQL SERVER 2000 ---
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
        Requete = "SELECT Premisses.* FROM Premisses ORDER BY NumPosteDepart, NumPosteArrivee"
        .CursorLocation = adUseServer
        .CacheSize = 50
        .MaxRecords = 0
        .Open Requete, ConnexionBDAnodisationSQL, adOpenStatic, adLockOptimistic, adCmdText

        '--- extraction des valeurs ---
        Do While Not .EOF
            
            '--- affectation ---
            NumPosteDepart = ![NumPosteDepart]
            NumPosteArrivee = ![NumPosteArrivee]
            
            '--- affectation dans le tableau ---
            If NumPosteDepart >= POSTES.P_CHGT_1 And NumPosteDepart <= DERNIER_POSTE And _
                NumPosteArrivee >= POSTES.P_CHGT_1 And NumPosteArrivee <= DERNIER_POSTE Then
                    With TPremisses(NumPosteDepart, NumPosteArrivee)
                        
                        .NumPont = Enregistrement![NumPont]
                        
                        .NumPontIA = .NumPont   'par d�faut NumPontIA = NumPont
                                                                  'le moteur d'inf�rence change en temps r�el cette valeur en fonction
                                                                  'des cas se pr�sentant dans la ligne
                        
                        .PremisseCodee = Enregistrement![PremisseCodee]
                        .PremisseDecodee = Enregistrement![PremisseDecodee]
                        
                        '--- calcul du temps de cycle en secondes (temps th�orique par apprentissage des temps de mouvements) ---
                        If CalculTempsCyclePremisse(NumPosteDepart, NumPosteArrivee, TempsCycleSecondes) = OK Then
                            .TempsCycleSecondes = TempsCycleSecondes
                        Else
                            .TempsCycleSecondes = 0
                        End If
                     
                     End With
            End If
            
            '--- enregistrement suivant ---
            .MoveNext
        
        Loop

    End With
    
    '--- fermeture de l'enregistrement / connexion ---
    Enregistrement.Close
    Set Enregistrement = Nothing
    ConnexionBDAnodisationSQL.Close
    Set ConnexionBDAnodisationSQL = Nothing

    '--- affichage du type de t�che ---
    AfficheTypeTache ("")
  
    Exit Function

GestionErreurs:
    
    '--- valeur de retour ---
    ChargePremisses = CStr(Err.Number)
    
    '--- fermeture de l'enregistrement / connexion ---
    On Error Resume Next
    Enregistrement.Close
    Set Enregistrement = Nothing
    ConnexionBDAnodisationSQL.Close
    Set ConnexionBDAnodisationSQL = Nothing
    
End Function

'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' R�le      : R�g�n�ration compl�te des pr�misses
' D�tails  :
' Entr�es :
' Retours :
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function RegenerationCompletePremisses() As String
    
    '--- aiguillage en cas d'erreurs ---
    On Error GoTo GestionErreurs
  
    '--- constante priv�es ---

    '--- d�claration ---
    Dim a As Integer, _
           NumPosteDepart As Integer, _
           NumPosteArrivee As Integer, _
           PremisseCodee As String, _
           PremisseDecodee As String
    Dim Requete As String
    Dim ConnexionBDAnodisationSQL As New ADODB.Connection
    Dim Enregistrement As New ADODB.Recordset
    
    '--- affichage du type de t�che ---
    AfficheTypeTache "R�g�n�ration compl�te des pr�misses"
    
    '--- ouverture de la connexion � la base de donn�es de l'anodisation en SQL SERVER 2000 ---
    With ConnexionBDAnodisationSQL
        .ConnectionString = PARAMETRES_CONNEXION_BD_ANODISATION_SQL
        '.CursorLocation = adUseServer
        .Mode = adModeReadWrite
        .ConnectionTimeout = 2     'X secondes d'attente de connexion avant de lancer un message d'erreur
        .Open
    End With

    '--- recherche ---
    With Enregistrement

        '--- lancement de la requ�te ---
        Requete = "SELECT Premisses.* FROM Premisses ORDER BY NumPosteDepart, NumPosteArrivee"
        .CursorLocation = adUseServer
        .CacheSize = 50
        .MaxRecords = 0
        .Open Requete, ConnexionBDAnodisationSQL, adOpenStatic, adLockOptimistic, adCmdText

        '--- extraction des valeurs ---
        Do While Not .EOF
            
            '--- affectation ---
            NumPosteDepart = ![NumPosteDepart]
            NumPosteArrivee = ![NumPosteArrivee]
            
            '--- modification de l'enregistrement (ATTENTION si la pr�misse existe d�j�) ---
            If NumPosteDepart >= POSTES.P_CHGT_1 And NumPosteDepart <= DERNIER_POSTE And _
                NumPosteArrivee >= POSTES.P_CHGT_1 And NumPosteArrivee <= DERNIER_POSTE Then
                            
                    '--- affectation ---
                    PremisseDecodee = Enregistrement![PremisseDecodee]
                    
                    If PremisseDecodee <> "" Then
                        PremisseDecodee = CalculAutomatiquePremisseDecodee(NumPosteDepart, NumPosteArrivee)
                        PremisseCodee = PremisseDecodeeVersCodee(PremisseDecodee)
                        Enregistrement![PremisseDecodee] = PremisseDecodee
                        Enregistrement![PremisseCodee] = PremisseCodee
                        '.UpdateBatch adAffectCurrent
                        .Update
                        
                        'Debug.Print TEtatsPostes(NumPosteDepart).DefinitionPoste.NomPoste, TEtatsPostes(NumPosteArrivee).DefinitionPoste.NomPoste
                    End If

            End If
            
            '--- enregistrement suivant ---
            .MoveNext
        
        Loop

    End With
    
    '--- fermeture de l'enregistrement / connexion ---
    Enregistrement.Close
    Set Enregistrement = Nothing
    ConnexionBDAnodisationSQL.Close
    Set ConnexionBDAnodisationSQL = Nothing
    
    '--- affichage du type de t�che ---
    AfficheTypeTache ("")
  
    Exit Function

GestionErreurs:
    
    '--- valeur de retour ---
    RegenerationCompletePremisses = CStr(Err.Number)
    
    '--- fermeture de l'enregistrement / connexion ---
    On Error Resume Next
    Enregistrement.Close
    Set Enregistrement = Nothing
    ConnexionBDAnodisationSQL.Close
    Set ConnexionBDAnodisationSQL = Nothing

End Function

'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' R�le      : Chargement des cuves
' D�tails  :
' Entr�es :
' Retours :
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function ChargeCuves() As String
    
    '--- aiguillage en cas d'erreurs ---
    On Error GoTo GestionErreurs

    ' C00
    CORRESPONDANCES_IDX_AUTOMATE(1) = 1
    ' DEC
    CORRESPONDANCES_IDX_AUTOMATE(2) = 2
    ' C7
    CORRESPONDANCES_IDX_AUTOMATE(3) = 7
    ' C13
    CORRESPONDANCES_IDX_AUTOMATE(4) = 8
    ' C14
    CORRESPONDANCES_IDX_AUTOMATE(5) = 9
    ' C15
    CORRESPONDANCES_IDX_AUTOMATE(6) = 10
    ' C22
    CORRESPONDANCES_IDX_AUTOMATE(7) = 13
    ' C27
    CORRESPONDANCES_IDX_AUTOMATE(8) = 14
    ' C28
    CORRESPONDANCES_IDX_AUTOMATE(9) = 15
    ' C31
    CORRESPONDANCES_IDX_AUTOMATE(10) = 16
    ' C32
    CORRESPONDANCES_IDX_AUTOMATE(11) = 17
    
    
    CORRESPONDANCES_IDX_CUVES_API(1) = 1
    CORRESPONDANCES_IDX_CUVES_API(2) = 2
    CORRESPONDANCES_IDX_CUVES_API(3) = -1
    CORRESPONDANCES_IDX_CUVES_API(4) = -1
    CORRESPONDANCES_IDX_CUVES_API(5) = -1
    CORRESPONDANCES_IDX_CUVES_API(6) = -1
    CORRESPONDANCES_IDX_CUVES_API(7) = 3
    CORRESPONDANCES_IDX_CUVES_API(8) = 4
    CORRESPONDANCES_IDX_CUVES_API(9) = 5
    CORRESPONDANCES_IDX_CUVES_API(10) = 6
    CORRESPONDANCES_IDX_CUVES_API(11) = -1
    CORRESPONDANCES_IDX_CUVES_API(12) = -1
    CORRESPONDANCES_IDX_CUVES_API(13) = 7
    CORRESPONDANCES_IDX_CUVES_API(14) = 8
    CORRESPONDANCES_IDX_CUVES_API(15) = 9
    CORRESPONDANCES_IDX_CUVES_API(16) = 10
    CORRESPONDANCES_IDX_CUVES_API(17) = 11
    CORRESPONDANCES_IDX_CUVES_API(18) = -1
    
    
    
    
    
    
    
    
    
    
    'Dim i As Integer
    'i = getCuveId(17)
    
    
   
    'Array(1, 2, 7, 13, 14, 15, 22, 27, 28, 31, 32)
    '--- constante priv�es ---

    '--- d�claration ---
    Dim a As Integer, _
           NumCuve As Integer, _
           CptCuvesAPI As Integer
    Dim Requete As String
    Dim ConnexionBDAnodisationSQL As New ADODB.Connection
    Dim Enregistrement As New ADODB.Recordset
    
    '--- affichage du type de t�che ---
    AfficheTypeTache "Chargement des cuves"

    '--- affectation ---
    CptCuvesAPI = 1
    
    

    
    '--- ouverture de la connexion � la base de donn�es de l'anodisation en SQL SERVER 2000 ---
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
        Requete = "SELECT Cuves.* FROM Cuves ORDER BY NumCuve"
        .CursorLocation = adUseServer
        .CacheSize = 50
        .MaxRecords = 0
        .Open Requete, ConnexionBDAnodisationSQL, adOpenStatic, adLockOptimistic, adCmdText
        
        '--- extraction des valeurs ---
        Do While Not .EOF
            
            '--- affectation ---
            NumCuve = ![NumCuve]
            
            '--- affectation dans le tableau ---
            With TCaracteristiquesCuves(NumCuve).DefinitionCuve
                
                .NumCuve = Enregistrement![NumCuve]
                .NomCuve = Enregistrement![NomCuve]
                .LibelleCuve = Enregistrement![LibelleCuve]
                .GestionAPI = Enregistrement![GestionAPI]
                .PresencePompe = Enregistrement![PresencePompe]
                .NbrChauffages = Enregistrement![NbrChauffages]
                .PresenceRefroidissementBain = Enregistrement![PresenceRefroidissementBain]
                .PresenceNiveauBas = Enregistrement![PresenceNiveauBas]
                .PresenceNiveauHaut = Enregistrement![PresenceNiveauHaut]
                .PresenceEVEau = Enregistrement![PresenceEVEau]
                '.PresenceAnalyseurAnodisation = Enregistrement![PresenceAnalyseurAnodisation]
            
            End With
            
            '--- construction du tableau des cuves g�r�es par l'automate ---
            If TCaracteristiquesCuves(NumCuve).DefinitionCuve.GestionAPI = True Then
                TEtatsCuves(CptCuvesAPI).DefinitionCuve = TCaracteristiquesCuves(NumCuve).DefinitionCuve
                TEtatsCuves(CptCuvesAPI).IndexAutomate = CORRESPONDANCES_IDX_AUTOMATE(CptCuvesAPI)
                Inc CptCuvesAPI
            End If
            
            '--- enregistrement suivant ---
            .MoveNext
        
        Loop

    End With
    
    '--- fermeture de l'enregistrement / connexion ---
    Enregistrement.Close
    Set Enregistrement = Nothing
    ConnexionBDAnodisationSQL.Close
    Set ConnexionBDAnodisationSQL = Nothing

    '--- affichage du type de t�che ---
    AfficheTypeTache ("")
  
    Exit Function

GestionErreurs:
    
    '--- valeur de retour ---
    ChargeCuves = CStr(Err.Number)

    '--- fermeture de l'enregistrement / connexion ---
    On Error Resume Next
    Enregistrement.Close
    Set Enregistrement = Nothing
    ConnexionBDAnodisationSQL.Close
    Set ConnexionBDAnodisationSQL = Nothing

End Function

'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' R�le      : Chargement des bains
' Retours :
' D�tails  :
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function ChargeBains() As String
    
    '--- aiguillage en cas d'erreurs ---
    'On Error GoTo GestionErreurs
  
    '--- constante priv�es ---

    '--- d�claration ---
'    Dim a As Integer, _
'           NbrBains As Integer
'    Dim Requete As String
'    Dim Enregistrement As New ADODB.Recordset
'
'    '--- affichage du type de t�che ---
'    AfficheTypeTache "Chargement des bains"
'
'    '--- recherche ---
'    With Enregistrement
'
'        '--- lancement de la requ�te ---
'        Requete = "SELECT Matieres.* FROM Matieres ORDER BY NumMatiere"
'        .CursorLocation = adUseServer
'        .CacheSize = 50
'        .MaxRecords = 0
'        .Open Requete, ConnexionBDAnodisationSQL, adOpenStatic, adLockOptimistic, adCmdText
'
'
'        '--- extraction des valeurs ---
'        Do While Not .EOF
'
'            '--- affectation ---
'            NumMatiere = ![NumMatiere]
'
'            '--- affectation dans le tableau ---
'            With TMatieres(NumMatiere)
'                .NumMatiere = NumMatiere
'                .LibelleMatiere = Enregistrement![LibelleMatiere]
'            End With
'
'            '--- enregistrement suivant ---
'            .MoveNext
'
'        Loop
'
'        '--- fermeture de l'enregistrement ---
'        .Close
'
'    End With
'
'    '--- affichage du type de t�che ---
'    AfficheTypeTache ("")
'
'    Exit Function
'
'GestionErreurs:
'
'    '--- fermeture de l'enregistrement ---
'    On Error Resume Next
'    Enregistrement.Close
'    Set Enregistrement = Nothing
'
'    '--- valeur de retour ---
'    ChargeMatieres = CStr(Err.Number)

End Function

' SZP 2021
Public Function getIDBARRE() As Integer
    
    '--- aiguillage en cas d'erreurs ---
    On Error GoTo GestionErreurs
  
    '--- constante priv�es ---

    '--- d�claration ---
    Dim a As Integer, _
           NumAction As Integer
    Dim Requete As String
    Dim ConnexionBD As New ADODB.Connection
    Dim Enregistrement As New ADODB.Recordset
    Dim cmd As New ADODB.Command
    
    
    '--- ouverture de la connexion � la base de donn�es de l'anodisation en SQL SERVER 2000 ---
    With ConnexionBD
        .ConnectionString = PARAMETRES_CONNEXION_BD_ANODISATION_SQL
        .CursorLocation = adUseServer
        .Mode = adModeRead
        .ConnectionTimeout = 2     'X secondes d'attente de connexion avant de lancer un message d'erreur
        .Open
    End With
   
    ' Set up a command object for the stored procedure.
    cmd.ActiveConnection = ConnexionBD
    cmd.CommandText = "usp_GetWebCallSequence"
    cmd.CommandType = adCmdStoredProc
    ' Execute command to run stored procedure
    Set Enregistrement = cmd.Execute
    
   

       
    
    
    
    ' Enregistrement.v

    Dim res As Integer
    res = Enregistrement.Fields(0)
    '--- fermeture de l'enregistrement / connexion ---
   
    Enregistrement.Close
     ConnexionBD.Close
   
    
    Set Enregistrement = Nothing
    Set ConnexionBD = Nothing
    Set cmd.ActiveConnection = Nothing
    'cmd = Nothing
    getIDBARRE = res
    
   
    Exit Function

GestionErreurs:
    
    '--- valeur de retour ---
    MsgBox (Err.Description)
    getIDBARRE = 0
    
    '--- fermeture de l'enregistrement / connexion ---
    On Error Resume Next
    Enregistrement.Close
    Set Enregistrement = Nothing
    ConnexionBD.Close
    Set ConnexionBD = Nothing
    

End Function



Sub insertionClipperPointage(ByVal NumCharge As Integer)

    
    If (mlID = 0) Then
        modMultiThreading.Initialize
        modMultiThreading.EnablePrivateMarshaling True
        Set mcInsertClipper = CreatePrivateObjectByNameInNewThread("CInsertionClipper", , mlID)
        'Call Log("Cr�ation de mcInsertClipper avec pid: " & mlID)
    End If

    'Call Log("DEBUT insertionClipperPointage -> AsynchDispMethodCall")
    AsynchDispMethodCall mlID, "insertionClipper", VbMethod, OccFSynoptique, "CopyComplete", NumCharge
    'Call Log("AsynchDispMethodCall has been called  in other thread")
    
 
   
End Sub






Public Function TEST_CLIPPER() As String
    
    '--- aiguillage en cas d'erreurs ---
    On Error GoTo GestionErreurs
  
  
    '--- affichage du type de t�che ---
    AfficheTypeTache "TEST CLIPPER"
    '--- constante priv�es ---

    '--- d�claration ---
    Dim a As Integer, _
           NumAction As Integer
    Dim Requete As String
    Dim ConnexionBDClipper As New ADODB.Connection
    Dim Enregistrement As New ADODB.Recordset
    
    '--- affichage du type de t�che ---
    AfficheTypeTache "Test clipper"
    ' MsgBox ("GA_DES2=TEST TOTO")
    ' Call Log("GA_DES2 TEST")
    '--- ouverture de la connexion � la base de donn�es de l'anodisation en SQL SERVER 2000 ---
    With ConnexionBDClipper
        .ConnectionString = PARAMETRES_CONNEXION_BD_CLIPPER_HF
        .CursorLocation = adUseServer
        .Mode = adModeReadWrite
        .ConnectionTimeout = 2     'X secondes d'attente de connexion avant de lancer un message d'erreur
        .Open
    End With
   
    '--- recherche ---
    With Enregistrement

        '--- lancement de la requ�te ---
        Requete = "SELECT CP.compl�ment  FROM GAMME G,COMPLES CP where GACLEUNIK = '367292' and CP.Cl�unik=G.GACLEUNIK AND CP.COPAR='GACPL01'"
        .CursorLocation = adUseServer
        .CacheSize = 50
        .MaxRecords = 0
        .Open Requete, ConnexionBDClipper, adOpenStatic, adLockOptimistic, adCmdText

        If Not .EOF Then
           '.MoveFirst
            TEST_CLIPPER = "GA_DES2:" & ![GA_DES2]
        Else
          TEST_CLIPPER = "GA_DES2 non trouv�"
       
        End If
       
     End With
     Enregistrement.Close

   



    '--- fermeture de l'enregistrement / connexion ---
    Enregistrement.Close
    Set Enregistrement = Nothing
    ConnexionBDClipper.Close
    Set ConnexionBDClipper = Nothing
    
   
    Exit Function

GestionErreurs:
    
    '--- valeur de retour ---
    TEST_CLIPPER = CStr(Err.Number)
    
    '--- fermeture de l'enregistrement / connexion ---
    On Error Resume Next
    Enregistrement.Close
    Set Enregistrement = Nothing
    ConnexionBDClipper.Close
    Set ConnexionBDClipper = Nothing

End Function
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' R�le      : Chargement des param�tres
' Retours :
' D�tails  :
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function ChargeParametres() As String
    
    '--- aiguillage en cas d'erreurs ---
    On Error GoTo GestionErreurs
  
    '--- constante priv�es ---


    '--- d�claration ---
    Dim a As Integer, _
           NumAction As Integer
    Dim Requete As String
    Dim ConnexionBDAnodisationSQL As New ADODB.Connection
    Dim Enregistrement As New ADODB.Recordset
    
    '--- affichage du type de t�che ---
    AfficheTypeTache "Chargement des actions"
    
    '--- ouverture de la connexion � la base de donn�es de l'anodisation en SQL SERVER 2000 ---
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
        Requete = "SELECT valeur  FROM Parametres where libell� ='DISTANCE_SECURITE'"
        .CursorLocation = adUseServer
        .CacheSize = 50
        .MaxRecords = 0
        .Open Requete, ConnexionBDAnodisationSQL, adOpenStatic, adLockOptimistic, adCmdText

        If Not .EOF Then
                '.MoveFirst
            DISTANCE_SECURITE = ![Valeur]
        Else
          ChargeParametres = "Param�tre 'DISTANCE_SECURITE' non pr�sent  !!"
          Exit Function
        End If
       
     End With
     Enregistrement.Close

    With Enregistrement
        Requete = "SELECT valeur  FROM Parametres where libell� ='DEBUG_MODE'"
        .CursorLocation = adUseServer
        .CacheSize = 50
        .MaxRecords = 0
        .Open Requete, ConnexionBDAnodisationSQL, adOpenStatic, adLockOptimistic, adCmdText

        If Not .EOF Then
                '.MoveFirst
            DEBUG_MODE = ![Valeur]
        Else
          ChargeParametres = "Param�tre 'DEBUG_MODE' non pr�sent  !!"
          Exit Function
        End If
        
        
        

    End With





    '--- fermeture de l'enregistrement / connexion ---
    Enregistrement.Close
    Set Enregistrement = Nothing
    ConnexionBDAnodisationSQL.Close
    Set ConnexionBDAnodisationSQL = Nothing
    
   
    Exit Function

GestionErreurs:
    
    '--- valeur de retour ---
    ChargeParametres = CStr(Err.Number)
    
    '--- fermeture de l'enregistrement / connexion ---
    On Error Resume Next
    Enregistrement.Close
    Set Enregistrement = Nothing
    ConnexionBDAnodisationSQL.Close
    Set ConnexionBDAnodisationSQL = Nothing

End Function

'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' R�le      : Chargement des actions
' Retours :
' D�tails  :
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function ChargeActions() As String
    
    '--- aiguillage en cas d'erreurs ---
    On Error GoTo GestionErreurs
  
    '--- constante priv�es ---

    '--- d�claration ---
    Dim a As Integer, _
           NumAction As Integer
    Dim Requete As String
    Dim ConnexionBDAnodisationSQL As New ADODB.Connection
    Dim Enregistrement As New ADODB.Recordset
    
    '--- affichage du type de t�che ---
    AfficheTypeTache "Chargement des actions"
    
    '--- ouverture de la connexion � la base de donn�es de l'anodisation en SQL SERVER 2000 ---
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
        Requete = "SELECT Actions.* FROM Actions ORDER BY NumAction"
        .CursorLocation = adUseServer
        .CacheSize = 50
        .MaxRecords = 0
        .Open Requete, ConnexionBDAnodisationSQL, adOpenStatic, adLockOptimistic, adCmdText

        '--- extraction des valeurs ---
        Do While Not .EOF
            
            '--- affectation ---
            NumAction = ![NumAction]
            
            '--- affectation dans le tableau ---
            If NumAction >= NUM_ACTION_NOP And NumAction <= NUM_ACTION_FCY Then
                TActions(NumAction).NumAction = NumAction
                TActions(NumAction).CodeAction = ![CodeAction]
                TActions(NumAction).LibelleAction = ![LibelleAction]
                TActions(NumAction).ParametreOuiNon = ![ParametreOuiNon]
                TActions(NumAction).LibelleParametre = ![LibelleParametre]
            End If
            
            '--- enregistrement suivant ---
            .MoveNext
        
        Loop

    End With

    '--- fermeture de l'enregistrement / connexion ---
    Enregistrement.Close
    Set Enregistrement = Nothing
    ConnexionBDAnodisationSQL.Close
    Set ConnexionBDAnodisationSQL = Nothing
    
    '--- affichage du type de t�che ---
    AfficheTypeTache ("")
  
    Exit Function

GestionErreurs:
    
    '--- valeur de retour ---
    ChargeActions = CStr(Err.Number)
    
    '--- fermeture de l'enregistrement / connexion ---
    On Error Resume Next
    Enregistrement.Close
    Set Enregistrement = Nothing
    ConnexionBDAnodisationSQL.Close
    Set ConnexionBDAnodisationSQL = Nothing

End Function

'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' R�le      : Chargement des d�fauts
' Retours :
' D�tails  :
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function ChargeDefauts() As String
    
    '--- aiguillage en cas d'erreurs ---
    On Error GoTo GestionErreurs
  
    '--- constante priv�es ---

    '--- d�claration ---
    Dim a As Integer, _
           NumDefaut As Integer
    Dim Requete As String
    Dim ConnexionBDAnodisationSQL As New ADODB.Connection
    Dim Enregistrement As New ADODB.Recordset
    
    '--- affichage du type de t�che ---
    AfficheTypeTache "Chargement des d�fauts"
    
    '--- ouverture de la connexion � la base de donn�es de l'anodisation en SQL SERVER 2000 ---
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
        Requete = "SELECT ListeDefauts.* FROM ListeDefauts ORDER BY NumDefaut"
        .CursorLocation = adUseServer
        .CacheSize = 50
        .MaxRecords = 0
        .Open Requete, ConnexionBDAnodisationSQL, adOpenStatic, adLockOptimistic, adCmdText

        '--- extraction des valeurs ---
        Do While Not .EOF
            
            '--- affectation ---
            NumDefaut = ![NumDefaut]
            
            '--- affectation dans le tableau ---
            TDefauts(NumDefaut).SignalerOuiNon = ![SignalerOuiNon]
            TDefauts(NumDefaut).GyrophareOuiNon = ![GyrophareOuiNon]
            TDefauts(NumDefaut).KlaxonOuiNon = ![KlaxonOuiNon]
            TDefauts(NumDefaut).MessageVocalOuiNon = ![MessageVocalOuiNon]
            TDefauts(NumDefaut).AfficheurOuiNon = ![AfficheurOuiNon]
            TDefauts(NumDefaut).InformationsAPI = ![InformationsAPI]
            TDefauts(NumDefaut).LibelleDefaut = ![LibelleDefaut]
            TDefauts(NumDefaut).LibelleDefautAfficheur = ![LibelleDefautAfficheur]
            TDefauts(NumDefaut).TNumIntervenants(1) = ![NumIntervenant1]
            TDefauts(NumDefaut).TNumIntervenants(2) = ![NumIntervenant2]
            TDefauts(NumDefaut).TNumIntervenants(3) = ![NumIntervenant3]
            TDefauts(NumDefaut).TNumIntervenants(4) = ![NumIntervenant4]
            TDefauts(NumDefaut).TNumIntervenants(5) = ![NumIntervenant5]
            
            '--- enregistrement suivant ---
            .MoveNext
        
        Loop

    End With
    
    '--- fermeture de l'enregistrement / connexion ---
    Enregistrement.Close
    Set Enregistrement = Nothing
    ConnexionBDAnodisationSQL.Close
    Set ConnexionBDAnodisationSQL = Nothing
    
    '--- affichage du type de t�che ---
    AfficheTypeTache ("")
  
    Exit Function

'--- gestion des erreurs ---
GestionErreurs:
    
    '--- valeur de retour ---
    ChargeDefauts = CStr(Err.Number)

    '--- fermeture de l'enregistrement / connexion ---
    On Error Resume Next
    Enregistrement.Close
    Set Enregistrement = Nothing
    ConnexionBDAnodisationSQL.Close
    Set ConnexionBDAnodisationSQL = Nothing

End Function


'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' R�le      : Extrait un enregistrement de la table des fiches d'atelier
' Entr�es :                Enregistrement -> Enregistrement de la table des fiches d'atelier
' Retours : TEnrCommandesInterne -> Tableau contenant l'image d'un enregistrement de la table des fiches d'atelier
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Sub ExtraitEnrCommandesInterne(ByVal Enregistrement As ADODB.Recordset, _
                                                                   ByRef TEnrCommandesInterne As EnrCommandesInterne)
    
    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next

    With TEnrCommandesInterne
            
        '--- extraction de l'enregistrement ---
        .NumCommandeInterne = C_Nullite_Champ(Enregistrement, "CdeInterne", 0)
        .CodeClient = C_Nullite_Champ(Enregistrement, "CodeClient", "")
        .Designation = C_Nullite_Champ(Enregistrement, "Designation", "")
        .NbrPieces = 0                      'affectation par d�faut car non pr�vu dans SAGE
        .Matiere = ""                         'affectation par d�faut car non pr�vu dans SAGE

    End With
    
End Sub

'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' R�le      : Copie une gamme d'anodisation
' Entr�es :         NumGammeACopier -> Num�ro de la gamme � copier
'                       NouveauNumGamme -> Nouveau num�ro de gamme
' Retours : CopieGammeAnodisation -> "" = pas d'incident sinon num�ro de l'erreur
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function CopieGammeAnodisation(ByVal NumGammeACopier As String, _
                                                                    ByVal NouveauNumGamme As String) As String

    '--- aiguillage en cas d'erreurs ---
    On Error GoTo GestionErreurs
    
    '--- d�claration ---
    Dim Requete As String
    Dim ConnexionBDAnodisationSQL As New ADODB.Connection
    Dim Enregistrement As New ADODB.Recordset
    Dim Enregistrement2 As New ADODB.Recordset
    
    '--- contr�le ---
    If NumGammeACopier = "" Or NouveauNumGamme = "" Then
        CopieGammeAnodisation = NON_TROUVE
        Exit Function
    End If
    
    '--- ouverture de la connexion � la base de donn�es de l'anodisation en SQL SERVER 2000 ---
    With ConnexionBDAnodisationSQL
        .ConnectionString = PARAMETRES_CONNEXION_BD_ANODISATION_SQL
        .CursorLocation = adUseServer
        .Mode = adModeReadWrite
        .ConnectionTimeout = 2     'X secondes d'attente de connexion avant de lancer un message d'erreur
        .Open
    End With
    
    With Enregistrement

        '--- affectation de la requ�te ---
        Requete = "INSERT INTO GammesAnodisation " & _
                         "(NumGamme, RefGamme, DateCreationGamme, NomGamme,Designation, Matiere1, Matiere2, Matiere3, Matiere4, Matiere5," & _
                         "Matiere6, Matiere7, Matiere8, Matiere9, Matiere10,TempsAvantPostePrincipalTexte, TempsPostePrincipalTexte," & _
                         "TempsApresPostePrincipalTexte, TempsTotalPostesTexte,TempsTotalEgouttagesTexte, TempsTotalGammeTexte," & _
                         "TempsAvantPostePrincipalSecondes, TempsPostePrincipalSecondes, TempsApresPostePrincipalSecondes, TempsTotalPostesSecondes, TempsTotalEgouttagesSecondes," & _
                         "TempsTotalGammeSecondes, ModeUouI, TempsPhase1, UPhase1, IPhase1, TempsPhase2, UPhase2, IPhase2, TempsPhase3, UPhase3, IPhase3, TempsPhase4, UPhase4, IPhase4) " & _
                         "SELECT '" & NouveauNumGamme & "'" & _
                         ", RefGamme, DateCreationGamme, NomGamme, Designation, Matiere1, Matiere2, Matiere3, Matiere4, Matiere5,Matiere6, Matiere7, Matiere8, Matiere9, Matiere10," & _
                         "TempsAvantPostePrincipalTexte, TempsPostePrincipalTexte,TempsApresPostePrincipalTexte, TempsTotalPostesTexte,TempsTotalEgouttagesTexte, TempsTotalGammeTexte," & _
                         "TempsAvantPostePrincipalSecondes,TempsPostePrincipalSecondes,TempsApresPostePrincipalSecondes,TempsTotalPostesSecondes,TempsTotalEgouttagesSecondes,TempsTotalGammeSecondes, " & _
                         "ModeUouI, TempsPhase1, UPhase1, IPhase1, TempsPhase2, UPhase2, IPhase2, TempsPhase3, UPhase3, IPhase3, TempsPhase4, UPhase4, IPhase4 " & _
                         "FROM GammesAnodisation " & _
                        "WHERE (NumGamme = '" & NumGammeACopier & "')"
        
        '--- lancement de la requ�te ---
        .CursorLocation = adUseServer
        .Open Requete, ConnexionBDAnodisationSQL, adOpenStatic, adLockOptimistic

    End With
    
    With Enregistrement2
        
        '--- affectation de la requ�te ---
        Requete = "INSERT INTO DetailsGammesAnodisation " & _
                         "(NumGamme, NumLigne, NumZone, TempsAuPosteTexte, TempsAlerteTexte, TempsEgouttageTexte, TempsAuPosteSecondes, TempsEgouttageSecondes) " & _
                         "SELECT '" & NouveauNumGamme & "'" & _
                         ", NumLigne, NumZone, TempsAuPosteTexte, TempsAlerteTexte, TempsEgouttageTexte, TempsAuPosteSecondes , TempsEgouttageSecondes " & _
                         "FROM DetailsGammesAnodisation " & _
                         "WHERE (NumGamme = '" & NumGammeACopier & "')"
        
        '--- lancement de la requ�te ---
        .CursorLocation = adUseServer
        .Open Requete, ConnexionBDAnodisationSQL, adOpenStatic, adLockOptimistic

    End With
    
    '--- fermeture de l'enregistrement / connexion ---
    Set Enregistrement = Nothing
    Set Enregistrement2 = Nothing
    ConnexionBDAnodisationSQL.Close
    Set ConnexionBDAnodisationSQL = Nothing
    
    Exit Function

'--- gestion des erreurs ---
GestionErreurs:
    
    '--- valeur de retour ---
    CopieGammeAnodisation = CStr(Err.Number)
    
    '--- fermeture de l'enregistrement / connexion ---
    On Error Resume Next
    Enregistrement.Close
    Set Enregistrement = Nothing
    ConnexionBDAnodisationSQL.Close
    Set ConnexionBDAnodisationSQL = Nothing

End Function

'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' R�le      : V�rifie l'existence d'une gamme d'anodisation
' Entr�es :                               NumGamme -> Num�ro de la gamme recherch�e
' Retours : ExistenceGammesAnodisation ->           TROUVE = Enregistrement trouv� ou valid�
'                                                                        NON_TROUVE = Recherche non trouv�e ou abandonn�e
'                                                                                                   autres valeurs = N� du message d'erreur
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function ExistenceGammesAnodisation(ByVal NumGamme As String) As String
    
    '--- aiguillage en cas d'erreurs ---
    On Error GoTo GestionErreurs
    
    '--- d�claration ---
    Dim Requete As String
    Dim ConnexionBDAnodisationSQL As New ADODB.Connection
    Dim Enregistrement As New ADODB.Recordset

    '--- contr�le ---
    If NumGamme = "" Then
        ExistenceGammesAnodisation = NON_TROUVE
        Exit Function
    End If
    
    '--- ouverture de la connexion � la base de donn�es de l'anodisation en SQL SERVER 2000 ---
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
        Requete = "SELECT NumGamme FROM GammesAnodisation WHERE (NumGamme = '" & NumGamme & "')"
        .CursorLocation = adUseServer
        .MaxRecords = 1
        .Open Requete, ConnexionBDAnodisationSQL, adOpenStatic, adLockReadOnly, adCmdText
        
        '--- affectation ---
        If .BOF = True Or .EOF = True Then
            ExistenceGammesAnodisation = NON_TROUVE
        Else
            ExistenceGammesAnodisation = TROUVE
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
    ExistenceGammesAnodisation = CStr(Err.Number)
    
    '--- fermeture de l'enregistrement / connexion ---
    On Error Resume Next
    Enregistrement.Close
    Set Enregistrement = Nothing
    ConnexionBDAnodisationSQL.Close
    Set ConnexionBDAnodisationSQL = Nothing
    
End Function


'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' R�le      : Chargement des mati�res
' Retours :
' D�tails  :
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function ChargeMatieres() As String
    
    '--- aiguillage en cas d'erreurs ---
    On Error GoTo GestionErreurs
  
    '--- constante priv�es ---

    '--- d�claration ---
    Dim a As Integer, _
           OrdrePourAffichage As Integer
    Dim Requete As String
    Dim ConnexionBDAnodisationSQL As New ADODB.Connection
    Dim Enregistrement As New ADODB.Recordset
    
    '--- affichage du type de t�che ---
    AfficheTypeTache "Chargement des mati�res"
    
    '--- ouverture de la connexion � la base de donn�es de l'anodisation en SQL SERVER 2000 ---
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
        Requete = "SELECT Matieres.* FROM Matieres ORDER BY OrdrePourAffichage"
        .CursorLocation = adUseServer
        .CacheSize = 50
        .MaxRecords = 0
        .Open Requete, ConnexionBDAnodisationSQL, adOpenStatic, adLockOptimistic, adCmdText

        '--- extraction des valeurs ---
        Do While Not .EOF
            
            '--- affectation ---
            OrdrePourAffichage = ![OrdrePourAffichage]
            
            '--- affectation dans le tableau ---
            With TMatieres(OrdrePourAffichage)
                .Matiere = Enregistrement![Matiere]
                .TypeMatiere = Enregistrement![TypeMatiere]
                .CompositionMatiere = Enregistrement![CompositionMatiere]
                .OrdrePourAffichage = OrdrePourAffichage
            End With
            
            '--- enregistrement suivant ---
            .MoveNext
        
        Loop

    End With

    '--- fermeture de l'enregistrement / connexion ---
    Enregistrement.Close
    Set Enregistrement = Nothing
    ConnexionBDAnodisationSQL.Close
    Set ConnexionBDAnodisationSQL = Nothing
    
    '--- affichage du type de t�che ---
    AfficheTypeTache ("")
  
    Exit Function

GestionErreurs:
    
    '--- valeur de retour ---
    ChargeMatieres = CStr(Err.Number)
    
    '--- fermeture de l'enregistrement / connexion ---
    On Error Resume Next
    Enregistrement.Close
    Set Enregistrement = Nothing
    ConnexionBDAnodisationSQL.Close
    Set ConnexionBDAnodisationSQL = Nothing
    
End Function


Public Sub EnregistrementProductionLocal(ByVal NumCharge As Integer)

    '--- aiguillage en cas d'erreurs ---
    On Error GoTo GestionErreurs

    Dim showLogs As Boolean
    showLogs = False
    
    '--- d�claration ---
    Dim a As Integer                                                'pour les boucles FOR...NEXT
    Dim b As Integer                                                'pour les boucles FOR...NEXT
    Dim NumRedresseur As Integer                        'num�ro d'un redresseur
    
    Dim MsgTracabilite As String
    
    Dim NumFicheProduction As String                   'num�ro de fiche de production
    
    Dim ConnexionBDAnodisationSQL As ADODB.Connection
    Dim Enregistrement As ADODB.Recordset
    
    Dim FicheVideEtatsCharges As etatsCharges
    
    '--- affectation ---
    'EnregistrementProduction = ""
    
    Call Log("ProchainNumFicheProduction  DEBUT", showLogs)
    '--- recherche du prochain num�ro de fiche de production ---
    NumFicheProduction = ProchainNumFicheProduction()
    Call Log("ProchainNumFicheProduction  FIN", showLogs)
                    
    If NumFicheProduction <> "" Then
    
        '--- ouverture de la connexion ---
        Set ConnexionBDAnodisationSQL = New ADODB.Connection
        With ConnexionBDAnodisationSQL
            .ConnectionString = PARAMETRES_CONNEXION_BD_ANODISATION_SQL
            .CursorLocation = adUseServer
            .Mode = adModeReadWrite
            .ConnectionTimeout = 2     'X secondes d'attente de connexion avant de lancer un message d'erreur
            .Open
        End With
        
        Call Log("DETAILS DES CHARGES DE PRODUCTION  DEBUT", showLogs)
        '--- extraction et enregistrement ---
        With TEtatsCharges(NumCharge)
    
            '****************************************************************************************************************
            '*                                                DETAILS DES CHARGES DE PRODUCTION
            '****************************************************************************************************************
            
            '--- ouverture de la table ---
            Set Enregistrement = New ADODB.Recordset
            With Enregistrement
                .CursorLocation = adUseServer
                .CursorType = adOpenStatic 'adOpenKeyset
                .LockType = adLockBatchOptimistic    'adLockOptimistic
                .Open TABLE_DETAILS_CHARGES_PRODUCTION, ConnexionBDAnodisationSQL, , adCmdTable
            End With
    
            '--- enregistrement des d�tails des charges ---
            For a = LBound(.TDetailsCharges()) To UBound(.TDetailsCharges())
                With .TDetailsCharges(a)
                    
                    If .NumCommandeInterne > 0 Then
                    
                        '--- enregistrement de la fiche ---
                        Enregistrement.AddNew
                        Enregistrement("NumCommandeInterne") = .NumCommandeInterne
                        Enregistrement("NbrReparations") = .TypeReparation
                        Enregistrement("DateEntreeEnLigne") = TEtatsCharges(NumCharge).DateEntreeEnLigne
                        Enregistrement("DateArriveeAuDechargement") = TEtatsCharges(NumCharge).DateArriveeAuDechargement
                        Enregistrement("NumBarre") = TEtatsCharges(NumCharge).NumBarre
                        Call Log("Enregistrement(NumBarre) = TEtatsCharges(NumCharge).NumBarre=" & TEtatsCharges(NumCharge).NumBarre, showLogs)
                        
                        Enregistrement("NumLigne") = a
                        Enregistrement("CodeClient") = .CodeClient
                        Enregistrement("NbrPieces") = .NbrPieces
                        Enregistrement("Designation") = .Designation
                        Enregistrement("NumLignesReferencesClient") = .NumLignesReferencesClient
                        Enregistrement("Matiere") = .Matiere
                        Enregistrement("NumGammeAnodisation") = TEtatsCharges(NumCharge).TGammesAnodisation.NumGamme
                        Enregistrement("RefGammeAnodisation") = TEtatsCharges(NumCharge).TGammesAnodisation.RefGamme
                        Enregistrement("TempsAnodisationTexte") = CTemps(TEtatsCharges(NumCharge).TempsTotalGammeRedresseur)
                        Enregistrement("NumFicheProduction") = NumFicheProduction
                        If TEtatsCharges(NumCharge).ChargePrioritaire = True Then
                            Enregistrement("ChargePrioritaire") = 1
                        Else
                            Enregistrement("ChargePrioritaire") = 0
                        End If
                        'Call Log("barre2 =" & NumCharge, showLogs)
                        Enregistrement("AlarmesLigne") = TEtatsCharges(NumCharge).AlarmesLigne
                        'Enregistrement.Update
                    
                    Else
                        
                        '--- sortie directe si plus de n� de fiche d�tails de charge ---
                        Exit For
            
                    End If
                
                End With
            Next a
            
            Enregistrement.UpdateBatch
            
            Enregistrement.Close
            Call Log("DETAILS DES CHARGES DE PRODUCTION  FIN", showLogs)
            Call Log("DETAILS DE LA GAMME D'ANODISATION DE PRODUCTION DEBUT", showLogs)
        
            '****************************************************************************************************************
            '*                                      DETAILS DE LA GAMME D'ANODISATION DE PRODUCTION
            '****************************************************************************************************************
            
            '--- ouverture de la table ---
            Set Enregistrement = New ADODB.Recordset
            With Enregistrement
                .CursorLocation = adUseServer
                .CursorType = adOpenStatic 'adOpenKeyset
                .LockType = adLockBatchOptimistic    'adLockOptimistic
                .Open TABLE_DETAILS_GAMMES_ANODISATION_PRODUCTION, ConnexionBDAnodisationSQL, , adCmdTable
            End With
            
            '--- enregistrement des d�tails de la gamme d'anodisation ---
            For a = LBound(.TGammesAnodisation.TDetailsGammesAnodisation()) To UBound(.TGammesAnodisation.TDetailsGammesAnodisation())
                
                With .TGammesAnodisation.TDetailsGammesAnodisation(a)
                    
                    If .NumZone <> 0 Then
                     
                        '--- enregistrement de la fiche ---
                        Enregistrement.AddNew
                        Enregistrement("NumFicheProduction") = NumFicheProduction
                        Enregistrement("NumLigne") = a
                        Enregistrement("NumZone") = .NumZone
                        Enregistrement("TempsAuPosteTexte") = .TempsAuPosteTexte
                        Enregistrement("TempsEgouttageTexte") = .TempsEgouttageTexte
                        Enregistrement("TempsAuPosteSecondes") = .TempsAuPosteSecondes
                        Enregistrement("TempsEgouttageSecondes") = .TempsEgouttageSecondes
                        Enregistrement("DecompteDuTempsAuPosteReelSecondes") = .DecompteDuTempsAuPosteReelSecondes
                        Enregistrement("NumPosteReel") = .NumPosteReel
                        
                        '--- affectation du num�ro de redresseur ---
                        Select Case .NumPosteReel
                            Case POSTES.P_C13: NumRedresseur = REDRESSEURS.R_C13
                            Case POSTES.P_C14: NumRedresseur = REDRESSEURS.R_C14
                            Case POSTES.P_C15: NumRedresseur = REDRESSEURS.R_C15
                            Case POSTES.P_C16: NumRedresseur = REDRESSEURS.R_C16
                            Case Else
                        End Select
                        
                        '--- enregistrement ---
                        'Enregistrement.Update
                    
                    Else
                        
                        '--- sortie directe si plus de n� de fiche d�tails de charge ---
                        Exit For
            
                    End If
                
                End With
            Next a
            Enregistrement.UpdateBatch
            Enregistrement.Close
            Call Log("DETAILS DE LA GAMME D'ANODISATION DE PRODUCTION FIN", showLogs)
            
            Call Log("TRACABILITE DES REDRESSEURS DEBUT", showLogs)
            '****************************************************************************************************************
            '*                                                  TRACABILITE DES REDRESSEURS
            '****************************************************************************************************************
            If NumRedresseur > 0 Then                 'enregistrement de la production uniquement si passage dans un
                                                                          'des redresseurs
                Bidon = SauveTra�abiliteRedresseurs(NumCharge:=NumCharge, _
                                                                               NumFicheProduction:=NumFicheProduction, _
                                                                               DateEntreeEnLigne:=TEtatsCharges(NumCharge).DateEntreeEnLigne, _
                                                                               NumRedresseur:=NumRedresseur)
            
            End If

            Call Log("TRACABILITE DES REDRESSEURS FIN", showLogs)
            Call Log("DETAILS DES PHASES DE PRODUCTION DEBUT", showLogs)
            '****************************************************************************************************************
            '*                                       DETAILS DES PHASES DE PRODUCTION
            '****************************************************************************************************************
            
            '--- ouverture de la table ---
            Set Enregistrement = New ADODB.Recordset
            With Enregistrement
                .CursorLocation = adUseServer
                .CursorType = adOpenStatic 'adOpenKeyset
                .LockType = adLockBatchOptimistic    'adLockOptimistic
                .Open TABLE_DETAILS_PHASES_PRODUCTION, ConnexionBDAnodisationSQL, , adCmdTable
            End With
            
            '--- enregistrement des d�tails de la gamme d'anodisation ---
            For a = LBound(.TDetailsPhasesProduction()) To UBound(.TDetailsPhasesProduction())
                
                With .TDetailsPhasesProduction(a)
                    
                    '--- enregistrement de la fiche ---
                    Enregistrement.AddNew
                    Enregistrement("NumFicheProduction") = NumFicheProduction
                    Enregistrement("NumRedresseur") = NumRedresseur
                    Enregistrement("ModeUouI") = TEtatsCharges(NumCharge).ModeUouI
                    Enregistrement("NumPhase") = a
                    Enregistrement("TempsPhase") = .TempsPhase
                    Enregistrement("UPhase") = .UPhase
                    Enregistrement("IPhase") = .IPhase
                    'Enregistrement.Update
                    
                End With
            
            Next a
            Enregistrement.UpdateBatch
            Enregistrement.Close
            
            
            '****************************************************************************************************************
            '*                                                 DETAILS DES FICHES DE PRODUCTION
            '****************************************************************************************************************
        
            Call Log("DETAILS DES PHASES DE PRODUCTION FIN", showLogs)
            Call Log("DETAILS DES FICHES DE PRODUCTION DEBUT", showLogs)
            '--- ouverture de la table ---
            Set Enregistrement = New ADODB.Recordset
            With Enregistrement
                .CursorLocation = adUseServer
                .CursorType = adOpenStatic 'adOpenKeyset
                .LockType = adLockBatchOptimistic    'adLockOptimistic
                .Open TABLE_DETAILS_FICHES_PRODUCTION, ConnexionBDAnodisationSQL, , adCmdTable
            End With
        
            '--- enregistrement des d�tails des fiches de production ---
            For a = LBound(.TDetailsFichesProduction()) To UBound(.TDetailsFichesProduction())
                
                With .TDetailsFichesProduction(a)
                    
                    If .NumPoste <> 0 Then
                   
                        '--- enregistrement de la fiche ---
                        Enregistrement.AddNew
                        Enregistrement("NumFicheProduction") = NumFicheProduction
                        Enregistrement("NumLigne") = a
                        Enregistrement("NumPoste") = .NumPoste
                        Enregistrement("DateEntreePoste") = .DateEntreePoste
                        Enregistrement("DateSortiePoste") = .DateSortiePoste
                        Enregistrement("DateDebutEgouttage") = .DateDebutEgouttage
                        Enregistrement("DateFinEgouttage") = .DateFinEgouttage
                        Enregistrement("TemperatureEnEntree") = .TemperatureEnEntree
                        Enregistrement("TemperatureEnSortie") = .TemperatureEnSortie
                        Enregistrement("GrapheTemperature") = .GrapheTemperature
                        Enregistrement("URedresseur") = .URedresseur
                        Enregistrement("IRedresseur") = .IRedresseur
                        Enregistrement("GrapheRedresseur") = .GrapheRedresseur
                        Enregistrement("AnalyseurEnEntree") = .AnalyseurEnEntree
                        Enregistrement("AnalyseurEnSortie") = .AnalyseurEnSortie
                        Enregistrement("GrapheAnalyseur") = .GrapheAnalyseur
                        Enregistrement("AlarmesPoste") = .AlarmesPoste
                        'Enregistrement.Update
                    
                    Else
                   
                        '--- sortie directe si plus de n� de fiche d�tails de charge ---
                        Exit For
           
                    End If
           
                End With
            Next a
            Enregistrement.UpdateBatch
            Enregistrement.Close
        
            Call Log("DETAILS DES FICHES DE PRODUCTION FIN", showLogs)
        End With
    Else
        
        Call Log("Pas de fiche Production trouv�e !!")
    End If
    
    '----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    
    '****************************************************************************************************************
    '*                                        VIDAGE DE LA CHARGE DANS LE TABLEAU
    '****************************************************************************************************************
    TEtatsCharges(NumCharge) = FicheVideEtatsCharges
    
    '----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    Call Log("FIN ENREGISTREMENT DE PRODUCTION FIN", showLogs)
    '--- fermeture des enregistrements / connexions ---
    Select Case Enregistrement.State
        Case adStateClosed
        Case Else: Enregistrement.Close
    End Select
    ConnexionBDAnodisationSQL.Close
    
    '--- effacement des objets ---
    Set Enregistrement = Nothing
    Set ConnexionBDAnodisationSQL = Nothing
    
    Exit Sub

GestionErreurs:
    
    '--- valeur de retour ---
    'EnregistrementProductionLocal = CStr(Err.Number)
    
    AfficheRenseignements ROUGE_0, "Erreur d'enregitrement en base: " & CStr(Err.Number) & vbCrLf
    Call Log("Erreur d'enregitrement en base: " & CStr(Err.Description))
    '--- fermeture de l'enregistrement / connexion ---
    On Error Resume Next
    Enregistrement.Close
    Set Enregistrement = Nothing
    ConnexionBDAnodisationSQL.Close
    Set ConnexionBDAnodisationSQL = Nothing

End Sub



'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' R�le      : Recherche le prochain num�ro d'une fiche de production
' Entr�es :
' Retours :
' D�tails  :
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function ProchainNumFicheProduction() As String

    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next
    
    '--- constantes priv�es ---
    
    '--- d�claration ---
    Dim Requete As String
    Dim ConnexionBDAnodisationSQL As New ADODB.Connection
    Dim Enregistrement As New ADODB.Recordset
        
    '--- affectation ---
    ProchainNumFicheProduction = ""
    
    '--- ouverture de la connexion ---
    With ConnexionBDAnodisationSQL
        .ConnectionString = PARAMETRES_CONNEXION_BD_ANODISATION_SQL
        .CursorLocation = adUseServer
        .Mode = adModeReadWrite
        .ConnectionTimeout = 2     'X secondes d'attente de connexion avant de lancer un message d'erreur
        .Open
    End With
    
    '--- recherche du dernier num�ro ---
    With Enregistrement
        
        '--- ouverture / pointer le premier enregistrement ---
        .CursorLocation = adUseServer
        .MaxRecords = 1
        Requete = "SELECT MAX(NumFicheProduction) AS Expr1 FROM DetailsChargesProduction"
        .Open Requete, ConnexionBDAnodisationSQL, adOpenStatic, adLockOptimistic, adCmdText
        .MoveFirst
        
        '--- affectation ---
        ProchainNumFicheProduction = Right("00000000" & CStr(CLng(Trim(Enregistrement("Expr1"))) + 1), 8)
     
    End With
     
    '--- fermeture de l'enregistrement / connexion ---
    Enregistrement.Close
    Set Enregistrement = Nothing
    ConnexionBDAnodisationSQL.Close
    Set ConnexionBDAnodisationSQL = Nothing
    
End Function

Public Function testRecord()
    On Error GoTo GestionErreurs

    Dim ConnexionBDAnodisationSQL As ADODB.Connection
    Dim Enregistrement As ADODB.Recordset
    
   Dim FicheVideEtatsCharges As etatsCharges
    
    
    
    
        '--- ouverture de la connexion ---
    Set ConnexionBDAnodisationSQL = New ADODB.Connection
    With ConnexionBDAnodisationSQL
        .ConnectionString = PARAMETRES_CONNEXION_BD_ANODISATION_SQL
        .CursorLocation = adUseServer
        .Mode = adModeReadWrite
        .ConnectionTimeout = 2     'X secondes d'attente de connexion avant de lancer un message d'erreur
        .Open
    End With


    
    '--- ouverture de la table ---
    Set Enregistrement = New ADODB.Recordset
    With Enregistrement
        .CursorLocation = adUseServer
        .CursorType = adOpenStatic 'adOpenKeyset
        .LockType = adLockBatchOptimistic    'adLockOptimistic
        .Open TABLE_DETAILS_CHARGES_PRODUCTION, ConnexionBDAnodisationSQL, , adCmdTable
    End With
    
    MsgBox ("ICI")
    
    Enregistrement.AddNew
    Enregistrement("ClePrimaire") = 100906
    Enregistrement("NumCommandeInterne") = "260797"
    Enregistrement("NbrReparations") = "5"
    Enregistrement("DateEntreeEnLigne") = Now
    Enregistrement("DateArriveeAuDechargement") = Now
    Enregistrement("NumBarre") = 15
    Enregistrement("NumLigne") = 1
    Enregistrement("CodeClient") = "ZDEV"
    Enregistrement("NbrPieces") = 15
    Enregistrement("Designation") = "croquete fido"
    Enregistrement("NumLignesReferencesClient") = "15654"
    Enregistrement("Matiere") = "CACA"
    Enregistrement("NumGammeAnodisation") = "000512"
    Enregistrement("RefGammeAnodisation") = "TOTOGAMME"
    Enregistrement("TempsAnodisationTexte") = "15:22"
    Enregistrement("NumFicheProduction") = 212154
    Enregistrement("ChargePrioritaire") = 1
    Enregistrement("AlarmesLigne") = "554"
    Enregistrement("ControleObservations") = "TOTOGAMME"
    
    Enregistrement.AddNew
    Enregistrement("ClePrimaire") = 100907
    Enregistrement("NumCommandeInterne") = "260798"
    Enregistrement("NbrReparations") = 1
    Enregistrement("DateEntreeEnLigne") = Now
    Enregistrement("DateArriveeAuDechargement") = Now
    Enregistrement("NumBarre") = 16
    Enregistrement("NumLigne") = 1
    Enregistrement("CodeClient") = "ZDEV"
    Enregistrement("NbrPieces") = 15
    Enregistrement("Designation") = "croquete cheba"
    Enregistrement("NumLignesReferencesClient") = "15854"
    Enregistrement("Matiere") = "CROTTE"
    Enregistrement("NumGammeAnodisation") = "000512"
    Enregistrement("RefGammeAnodisation") = "TOTO GAMME"
    Enregistrement("TempsAnodisationTexte") = "25:42"
    Enregistrement("NumFicheProduction") = 212155
    Enregistrement("ChargePrioritaire") = 0
    Enregistrement("AlarmesLigne") = "54"
    Enregistrement("ControleObservations") = "TOTOGAMME"
        
            
    Enregistrement.UpdateBatch
    
    Enregistrement.Close
    
    Select Case Enregistrement.State
        Case adStateClosed
        Case Else: Enregistrement.Close
    End Select
    ConnexionBDAnodisationSQL.Close
    
    '--- effacement des objets ---
    Set Enregistrement = Nothing
    Set ConnexionBDAnodisationSQL = Nothing
    
    Exit Function
GestionErreurs:
    
    '--- valeur de retour ---
   
    
    MsgBox ("Erreur d'enregitrement en base: " & CStr(Err.Description))
    
    '--- fermeture de l'enregistrement / connexion ---
    On Error Resume Next
    Enregistrement.Close
    Set Enregistrement = Nothing
    ConnexionBDAnodisationSQL.Close
    Set ConnexionBDAnodisationSQL = Nothing
    
    
End Function





'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' R�le      : Enregistrement complet de la production d'une charge
' Entr�es :                        NumCharge -> Num�ro de la charge concern�e
' Retours : EnregistrementProduction -> "" = pas d'incident sinon num�ro de l'erreur
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function EnregistrementProduction(ByVal NumCharge As Integer) As String
   
    'Call Log("DEBUT EnregistrementProduction -> AsynchDispMethodCall")
    AsynchDispMethodCall mlID, "EnregistrementProductionAutreThread", VbMethod, OccFSynoptique, "CopyComplete", NumCharge
    'Call Log("AsynchDispMethodCall for EnregistrementProduction  has been called  in other thread")
   
   
End Function

'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' R�le      : Enregistrement complet des temps de mouvements (couvercles et ponts)
' Entr�es :
' Retours : TempsDeMouvements -> "" = pas d'incident sinon num�ro de l'erreur
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function EnregistrementTempsMouvements() As String
    
    '--- aiguillage en cas d'erreurs ---
    On Error GoTo GestionErreurs

    '--- constantes priv�es ---
    
    '--- d�claration ---
    Dim a As Integer, _
            NumPont As Integer, _
            NumCuveAPI As Integer, _
            NumPosteDepart As Integer, _
            NumPosteArrivee As Integer
    Dim Requete As String
    Dim ConnexionBDAnodisationSQL As New ADODB.Connection
    Dim Enregistrement As ADODB.Recordset
    
    '--- affectation ---
    EnregistrementTempsMouvements = ""
    
    '--- ouverture de la connexion ---
    With ConnexionBDAnodisationSQL
        .ConnectionString = PARAMETRES_CONNEXION_BD_ANODISATION_SQL
        .CursorLocation = adUseServer
        .Mode = adModeReadWrite
        .ConnectionTimeout = 2     'X secondes d'attente de connexion avant de lancer un message d'erreur
        .Open
    End With
        
    '*********************************************************************************************************************
    '*                                 TEMPS DE MOUVEMENTS DES PONTS SANS LA TRANSLATION
    '*********************************************************************************************************************
    Set Enregistrement = New ADODB.Recordset
    With Enregistrement

        '--- lancement de la requ�te ---
        Requete = "SELECT TempsMouvementsPontsSansTranslation.* FROM TempsMouvementsPontsSansTranslation ORDER BY NumPont"
        .CursorLocation = adUseServer
        .CacheSize = 50
        .MaxRecords = 0
        .Open Requete, ConnexionBDAnodisationSQL, adOpenStatic, adLockOptimistic, adCmdText

        '--- extraction des valeurs ---
        Do While Not .EOF
            
            '--- affectation ---
            NumPont = ![NumPont]
            
            '--- affectation dans le tableau ---
            If NumPont >= PONTS.P_1 And NumPont <= PONTS.P_2 Then
                With TEtatsPonts(NumPont).TTempsMouvements
                    Enregistrement![TempsAccrochesChargeVersHaut] = .TempsAccrochesChargeVersHaut
                    Enregistrement![TempsAccrochesChargeVersBas] = .TempsAccrochesChargeVersBas
                    Enregistrement![TempsDescenteHautVersBas] = .TempsDescenteHautVersBas
                    Enregistrement![TempsDescenteIntermediaireVersBas] = .TempsDescenteIntermediaireVersBas
                    Enregistrement![TempsMonteeBasVersIntermediaire] = .TempsMonteeBasVersIntermediaire
                    Enregistrement![TempsMonteeBasVersHaut] = .TempsMonteeBasVersHaut
                 End With
            End If
            
            '--- m�morisation ---
            .Update
            
            '--- enregistrement suivant ---
            .MoveNext
        
        Loop

        '--- fermeture de l'enregistrement ---
        .Close

    End With
    
    '*********************************************************************************************************************
    '*                                   TEMPS DE MOUVEMENTS DE LA TRANSLATION DES PONTS
    '*********************************************************************************************************************
    Set Enregistrement = New ADODB.Recordset
    With Enregistrement

        '--- lancement de la requ�te ---
        Requete = "SELECT TempsMouvementsTranslationPonts.* FROM TempsMouvementsTranslationPonts ORDER BY NumPont, NumPosteDepart, NumPosteArrivee"
        .CursorLocation = adUseServer
        .CacheSize = 50
        .MaxRecords = 0
        .Open Requete, ConnexionBDAnodisationSQL, adOpenStatic, adLockOptimistic, adCmdText

        '--- extraction des valeurs ---
        Do While Not .EOF
            
            '--- affectation ---
            NumPont = ![NumPont]
            NumPosteDepart = ![NumPosteDepart]
            NumPosteArrivee = ![NumPosteArrivee]
            
            '--- affectation dans le tableau ---
            If NumPont >= PONTS.P_1 And NumPont <= PONTS.P_2 And _
               NumPosteDepart >= POSTES.P_CHGT_1 And NumPosteDepart <= DERNIER_POSTE And _
               NumPosteArrivee >= POSTES.P_CHGT_1 And NumPosteArrivee <= DERNIER_POSTE Then
                    With TEtatsPonts(NumPont).TTempsMouvements
                        
                        If .TTempsTranslation(NumPosteDepart, NumPosteArrivee) <> 0 Then
                        
                            '--- changement des valeurs ---
                            Enregistrement![TempsTranslation] = .TTempsTranslation(NumPosteDepart, NumPosteArrivee)
                    
                            '--- m�morisation ---
                            Enregistrement.Update
                    
                        End If
                    
                    End With
            End If
            
            '--- enregistrement suivant ---
            .MoveNext
        
        Loop

        '--- fermeture de l'enregistrement ---
        .Close

    End With
    
    '--- fermeture des enregistrements / connexions ---
    Select Case Enregistrement.State
        Case adStateClosed
        Case Else: Enregistrement.Close
    End Select
    ConnexionBDAnodisationSQL.Close
    
    '--- effacement des objets ---
    Set Enregistrement = Nothing
    Set ConnexionBDAnodisationSQL = Nothing
    
    Exit Function

GestionErreurs:
    
    '--- valeur de retour ---
    EnregistrementTempsMouvements = CStr(Err.Number)

    '--- fermeture de l'enregistrement / connexion ---
    On Error Resume Next
    Enregistrement.Close
    Set Enregistrement = Nothing
    ConnexionBDAnodisationSQL.Close
    Set ConnexionBDAnodisationSQL = Nothing

End Function

'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' R�le      : Enregistrement d'un d�faut dans la tra�abilit� des alarmes
' Entr�es :                                                         NumDefaut -> Num�ro du d�faut que l'on souhaite enregistr�
' Retours : EnregistrementDefautDansTra�abiliteAlarmes -> "" = pas d'incident sinon num�ro de l'erreur
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function EnregistrementDefautDansTra�abiliteAlarmes(ByVal NumDefaut As Integer, EtatDefaut As Boolean) As String
    
    '--- aiguillage en cas d'erreurs ---
    On Error GoTo GestionErreurs

    '--- constantes priv�es ---
    
    '--- d�claration ---
    Dim ConnexionBDAnodisationSQL As New ADODB.Connection
    Dim Enregistrement As New ADODB.Recordset
    Dim Requete As String
    Dim ComplementDefaut As String                     'contient le texte du compl�ment ajout� au libell� du d�faut
    Dim LibelleCompleteDefaut As String               'repr�sente un libell� coml�t� d'un d�faut (pour les num�ros de d�faut des variateurs, etc ...)
    
    '--- affectation ---
    EnregistrementDefautDansTra�abiliteAlarmes = ""
        
    '--- ouverture de la connexion ---
    With ConnexionBDAnodisationSQL
        .ConnectionString = PARAMETRES_CONNEXION_BD_ANODISATION_SQL
        .CursorLocation = adUseServer
        .Mode = adModeReadWrite
        .ConnectionTimeout = 2     'X secondes d'attente de connexion avant de lancer un message d'erreur
        .Open
    End With
    
    If EtatDefaut = True Then
    
        With Enregistrement
              
            '--- ouverture de la table ---
            .CursorLocation = adUseServer
            .CursorType = adOpenKeyset
            .LockType = adLockOptimistic
            .Open TABLE_TRACABILITE_ALARMES, ConnexionBDAnodisationSQL, , adCmdTable
        
            '--- enregistrement du d�faut ---
            .AddNew
            !NumDefaut = NumDefaut
            !DateDetectionDefaut = Now
            
            '--- compl�ment du d�faut ---
            LibelleCompleteDefaut = CompleteLibelleDefaut(NumDefaut, ComplementDefaut)
            !ComplementDefaut = ComplementDefaut
            
            .Update
            .Close
        
        End With
    
    Else
        
        With Enregistrement
              
              '--- lancement d'une requ�te ---
            .CursorLocation = adUseServer
            .CursorType = adOpenKeyset
            .LockType = adLockOptimistic
            Requete = "UPDATE " & TABLE_TRACABILITE_ALARMES & _
                              " SET DateCorrectionDefaut='" & CStr(Now) & "'" & _
                              "WHERE NumDefaut=" & NumDefaut & " AND ISDATE(DateCorrectionDefaut)=0"
            .Open Requete, ConnexionBDAnodisationSQL, , adCmdText
        
        End With
   
   End If
    
    '--- effacement des objets ---
    Set Enregistrement = Nothing
    Set ConnexionBDAnodisationSQL = Nothing
    
    Exit Function

GestionErreurs:
    
    '--- valeur de retour ---
    EnregistrementDefautDansTra�abiliteAlarmes = CStr(Err.Number)
    
    '--- fermeture de l'enregistrement / connexion ---
    On Error Resume Next
    Enregistrement.Close
    Set Enregistrement = Nothing
     ConnexionBDAnodisationSQL.Close
    Set ConnexionBDAnodisationSQL = Nothing

End Function

'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' R�le      : Construction de l'impression de la tra�abilit� d'une charge
' Entr�es :                                  NumFicheProduction -> Num�ro de la fiche de production
' Retours : ConstructionImpressionTracabiliteCharge -> "" = pas d'incident sinon num�ro de l'erreur
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function ConstructionImpressionTracabiliteCharge(ByVal NumFicheProduction As String) As String
    
    '--- aiguillage en cas d'erreurs ---
    On Error GoTo GestionErreurs

    '--- constantes priv�es ---
    
    '--- d�claration ---
    Dim a As Integer
    Dim TempsEnSecondes  As Long
    Dim DateDuJour As String, _
           Texte As String
    Dim ConnexionBDAnodisationSQL As ADODB.Connection
    Dim Enregistrement As ADODB.Recordset
    
    '--- affectation ---
    ConstructionImpressionTracabiliteCharge = ""
    
    If NumFicheProduction <> "" Then
    
        '--- ouverture de la connexion ---
        Set ConnexionBDAnodisationSQL = New ADODB.Connection
        With ConnexionBDAnodisationSQL
            .ConnectionString = PARAMETRES_CONNEXION_BD_ANODISATION_SQL
            .CursorLocation = adUseServer
            .Mode = adModeReadWrite
            .ConnectionTimeout = 2     'X secondes d'attente de connexion avant de lancer un message d'erreur
            .Open
        End With
            
        '--- ouverture de la table ---
        Set Enregistrement = New ADODB.Recordset
        With Enregistrement
            .CursorLocation = adUseServer
            .CursorType = adOpenKeyset
            .LockType = adLockOptimistic
            .Open TABLE_IMP_TRACABILITE_CHARGE_1, ConnexionBDAnodisationSQL, , adCmdTable
            .MoveFirst
        End With

        '--- enregistrement du n� de la fiche de production ---
        Enregistrement("NumFicheProduction") = NumFicheProduction
        
        '--- enregistrement de la date du jour ---
        DateDuJour = Format(Now, "dd/mm/yyyy")
        Enregistrement("DateDuJour") = DateDuJour
        
        '--- extraction des donn�es de la production ---
        If RechercheDetailsChargesProduction(NumFicheProduction) = TROUVE Then
            
            '--- enregistrement des valeurs ---
            Enregistrement("DateEntreeEnLigne") = Format(TTempEnrDetailsChargesProduction(1).DateEntreeEnLigne, "dd/mm/yyyy � hh:nn:ss")
            Enregistrement("ChargePrioritaire") = IIf(TTempEnrDetailsChargesProduction(1).ChargePrioritaire = True, "OUI", "NON")
            
        Else
        
            '--- affectation ---
            Enregistrement("DateEntreeEnLigne") = ""
            Enregistrement("ChargePrioritaire") = ""
        
        End If
        
        '--- mise � jour ---
        Enregistrement.Update
        
        '--- fermeture des enregistrements ---
        Enregistrement.Close
        Set Enregistrement = Nothing
        
        '********************************************************************************************************************
        '                                                               CONTRUCTION DES DETAILS
        '********************************************************************************************************************
        
        '--- effacement de la table ---
        Set Enregistrement = New ADODB.Recordset
        With Enregistrement
            .CursorLocation = adUseServer
            .CursorType = adOpenKeyset
            .LockType = adLockOptimistic
            .Open "DELETE FROM " & TABLE_IMP_DETAILS_TRACABILITE_CHARGE_1, ConnexionBDAnodisationSQL, , adCmdText
        End With
        
        '--- ouverture de la table ---
        Set Enregistrement = New ADODB.Recordset
        With Enregistrement
            .CursorLocation = adUseServer
            .CursorType = adOpenKeyset
            .LockType = adLockOptimistic
            .Open TABLE_IMP_DETAILS_TRACABILITE_CHARGE_1, ConnexionBDAnodisationSQL, , adCmdTable
        End With
        
        '--- construction de la table des d�tails ---
        If RechercheDetailsFichesProduction(NumFicheProduction) = TROUVE Then
            For a = LBound(TTempEnrDetailsFichesProduction()) To UBound(TTempEnrDetailsFichesProduction())
                With TTempEnrDetailsFichesProduction(a)
                    
                    '--- nouvel enregistrement ---
                    Enregistrement.AddNew
                    
                    '--- construction de la fiche ---
                    Enregistrement("NumFicheProduction") = NumFicheProduction
                    Enregistrement("NumLigne") = .NumLigne
                    
                    '--- nom et libell� du poste ---
                    Enregistrement("NomPoste") = TEtatsPostes(.NumPoste).DefinitionPoste.NomPoste
                    Enregistrement("LibellePoste") = TEtatsPostes(.NumPoste).DefinitionPoste.LibellePoste
                
                    '--- temps r�el au poste ---
                    Texte = "Entr�e le " & Format(.DateEntreePoste, FORMAT_DATE_HEURE_1) & vbCr
                    If .DateSortiePoste = Empty Then
                        Texte = Texte & "-" & vbCr & "-"
                    Else
                        TempsEnSecondes = DateDiff("s", .DateEntreePoste, .DateSortiePoste)
                        Texte = Texte & _
                                     "Sortie le " & Format(.DateSortiePoste, FORMAT_DATE_HEURE_1) & vbCr & _
                                     "Temps = " & CTemps2(TempsEnSecondes)
                    End If
                    Enregistrement("TempsReelPoste") = Texte

                    '--- temps r�el d'�gouttage ---
                    If .DateDebutEgouttage = Empty Then
                        Texte = "-" & vbCr
                    Else
                        Texte = "D�but le " & Format(.DateDebutEgouttage, FORMAT_DATE_HEURE_1) & vbCr
                    End If
                    If .DateFinEgouttage = Empty Then
                        Texte = Texte & "-" & vbCr & "-"
                    Else
                        TempsEnSecondes = DateDiff("s", .DateDebutEgouttage, .DateFinEgouttage)
                        Texte = Texte & _
                                     "Fin le " & Format(.DateFinEgouttage, FORMAT_DATE_HEURE_1) & vbCr & _
                                     "Temps = " & CTemps2(TempsEnSecondes)
                    End If
                    Enregistrement("TempsReelEgouttage") = Texte

                    '--- temp�ratures ---
                    If .TemperatureEnEntree = 0 Then
                        Texte = "-" & vbCr & "-"
                    Else
                        Texte = "En entrant : " & Format(.TemperatureEnEntree, FORMAT_TEMPERATURE_1_DECIMALE_UNITE)
                        If .TemperatureEnSortie = 0 Then
                            Texte = Texte & vbCr & "-"
                        Else
                            Texte = Texte & vbCr & _
                                         "En sortant : " & Format(.TemperatureEnSortie, FORMAT_TEMPERATURE_1_DECIMALE_UNITE)
                        End If
                    End If
                    Enregistrement("Temperatures") = Texte

                    '--- redresseur ---
                    If .URedresseur = 0 Then
                        Texte = "-" & vbCr & "-"
                    Else
                        Select Case .NumPoste
                            Case POSTES.P_C13, POSTES.P_C14, POSTES.P_C15, POSTES.P_C16
                                'If .SensRedresseur = SENS_REDRESSEUR.R_EN_CATHODIQUE_OU_POLARISATION Then
                                '    Texte = TEXTE_POLARISATION
                                'Else
                                 '   Texte = TEXTE_AMORCAGE
                                'End If
                            Case Else
                        End Select
                        Texte = Texte & vbCr & "U = " & Format(.URedresseur, FORMAT_TENSION_1_DECIMALE_UNITE)
                        If .IRedresseur = 0 Then
                            Texte = Texte & vbCr & "-"
                        Else
                            Texte = Texte & vbCr & _
                                         "I = " & Format(.IRedresseur, FORMAT_INTENSITE_ENTIER_UNITE)
                        End If
                    End If
                    Enregistrement("Redresseurs") = Texte

                    '--- analyseur ---
                    If .AnalyseurEnEntree = 0 Then
                        Texte = "-" & vbCr & "-"
                    Else
                        Texte = "En entrant : " & Format(.AnalyseurEnEntree, FORMAT_ANALYSEUR_UNITE)
                        If .AnalyseurEnSortie = 0 Then
                            Texte = Texte & vbCr & "-"
                        Else
                            Texte = Texte & vbCr & _
                                         "En sortant : " & Format(.AnalyseurEnSortie, FORMAT_ANALYSEUR_UNITE)
                        End If
                    End If
                    Enregistrement("Analyseur") = Texte

                    '--- alarmes de poste ---
                    '.AlarmesPoste = "101-102-103-104-105-106"   'pour essai
                    If .AlarmesPoste = "" Then
                        Texte = "Pas d'alarmes"
                    Else
                        Texte = DecodeAlarmesPoste(.AlarmesPoste)
                    End If
                    Enregistrement("AlarmesPoste") = Texte
                
                    '--- mise � jour ---
                    Enregistrement.Update
                
                End With
            Next a
        End If
        
        '--- fermeture des enregistrements / connexions ---
        Select Case Enregistrement.State
            Case adStateClosed
            Case Else: Enregistrement.Close
        End Select
        ConnexionBDAnodisationSQL.Close
        
        '--- effacement des objets ---
        Set Enregistrement = Nothing
        Set ConnexionBDAnodisationSQL = Nothing
        
    End If
        
    Exit Function
    
GestionErreurs:
    
    '--- valeur de retour ---
    ConstructionImpressionTracabiliteCharge = CStr(Err.Number)
    
    '--- fermeture de l'enregistrement / connexion ---
    On Error Resume Next
    Enregistrement.Close
    Set Enregistrement = Nothing
    ConnexionBDAnodisationSQL.Close
    Set ConnexionBDAnodisationSQL = Nothing

End Function

'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' R�le      : Construction de l'impression des alarmes de la ligne
' Entr�es :                            NumFicheProduction -> Num�ro de la fiche de production
' Retours : ConstructionImpressionAlarmesLigne -> "" = pas d'incident sinon num�ro de l'erreur
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function ConstructionImpressionAlarmesLigne(ByVal NumFicheProduction As String) As String
    
    '--- aiguillage en cas d'erreurs ---
    On Error GoTo GestionErreurs

    '--- constantes priv�es ---
    
    '--- d�claration ---
    Dim a As Integer
    Dim TempsEnSecondes  As Long
    Dim DateDuJour As String, _
           AlarmesLigne As String, _
           Texte As String
    Dim ConnexionBDAnodisationSQL As ADODB.Connection
    Dim Enregistrement As ADODB.Recordset
    Dim TAlarmesLigne As Variant
    
    '--- affectation ---
    ConstructionImpressionAlarmesLigne = ""
    
    If NumFicheProduction <> "" Then
    
        '--- ouverture de la connexion ---
        Set ConnexionBDAnodisationSQL = New ADODB.Connection
        With ConnexionBDAnodisationSQL
            .ConnectionString = PARAMETRES_CONNEXION_BD_ANODISATION_SQL
            .CursorLocation = adUseServer
            .Mode = adModeReadWrite
            .ConnectionTimeout = 2     'X secondes d'attente de connexion avant de lancer un message d'erreur
            .Open
        End With
            
        '--- ouverture de la table ---
        Set Enregistrement = New ADODB.Recordset
        With Enregistrement
            .CursorLocation = adUseServer
            .CursorType = adOpenKeyset
            .LockType = adLockOptimistic
            .Open TABLE_IMP_ALARMES_LIGNE_1, ConnexionBDAnodisationSQL, , adCmdTable
            .MoveFirst
        End With

        '--- enregistrement du n� de la fiche de production ---
        Enregistrement("NumFicheProduction") = NumFicheProduction
        
        '--- enregistrement de la date du jour ---
        DateDuJour = Format(Now, "dd/mm/yyyy")
        Enregistrement("DateDuJour") = DateDuJour
        
        '--- extraction des donn�es de la production ---
        If RechercheDetailsChargesProduction(NumFicheProduction) = TROUVE Then
            
            '--- enregistrement des valeurs ---
            Enregistrement("DateEntreeEnLigne") = Format(TTempEnrDetailsChargesProduction(1).DateEntreeEnLigne, "dd/mm/yyyy � hh:nn:ss")
            Enregistrement("ChargePrioritaire") = IIf(TTempEnrDetailsChargesProduction(1).ChargePrioritaire = True, "OUI", "NON")
            
            '--- extraction des alarmes ---
            AlarmesLigne = TTempEnrDetailsChargesProduction(1).AlarmesLigne
            
        Else
        
            '--- affectation ---
            Enregistrement("DateEntreeEnLigne") = ""
            Enregistrement("ChargePrioritaire") = ""
        
        End If
        
        '--- mise � jour ---
        Enregistrement.Update
        
        '--- fermeture des enregistrements ---
        Enregistrement.Close
        Set Enregistrement = Nothing
        
        '********************************************************************************************************************
        '                                                               CONTRUCTION DES DETAILS
        '********************************************************************************************************************
        
        '--- effacement de la table ---
        Set Enregistrement = New ADODB.Recordset
        With Enregistrement
            .CursorLocation = adUseServer
            .CursorType = adOpenKeyset
            .LockType = adLockOptimistic
            .Open "DELETE FROM " & TABLE_IMP_DETAILS_ALARMES_LIGNE_1, ConnexionBDAnodisationSQL, , adCmdText
        End With

        '--- ouverture de la table ---
        Set Enregistrement = New ADODB.Recordset
        With Enregistrement
            .CursorLocation = adUseServer
            .CursorType = adOpenKeyset
            .LockType = adLockOptimistic
            .Open TABLE_IMP_DETAILS_ALARMES_LIGNE_1, ConnexionBDAnodisationSQL, , adCmdTable
        End With

        '--- construction de la table des d�tails ---
        'AlarmesLigne = "1-2-3-4-5-6-7-8-9-10-11-12-13-14-15-16-17-18-19-20"     'pour essai
        If AlarmesLigne = "" Then
                    
            '--- nouvel enregistrement ---
            Enregistrement.AddNew
            
            '--- n� du d�faut ---
            Enregistrement("NumFicheProduction") = NumFicheProduction
            Enregistrement("NumDefaut") = "-"
            Enregistrement("LibelleDefaut") = "AUCUN INCIDENT DURANT CE TRAITEMENT"

            '--- mise � jour ---
            Enregistrement.Update
       
       Else
        
            '--- construction du tableau contenant les num�ros d'alarmes ---
            TAlarmesLigne = Split(AlarmesLigne, SEPARATEUR_NUM_DEFAUTS)
                            
            '--- construction de la chaine des libell�s ---
            For a = LBound(TAlarmesLigne) To UBound(TAlarmesLigne)
                If IsNumeric(TAlarmesLigne(a)) = True Then
            
                    '--- nouvel enregistrement ---
                    Enregistrement.AddNew
            
                    '--- n� du d�faut ---
                    Enregistrement("NumFicheProduction") = NumFicheProduction
                    Enregistrement("NumDefaut") = TAlarmesLigne(a)
                    Enregistrement("LibelleDefaut") = TDefauts(TAlarmesLigne(a)).LibelleDefaut

                    '--- mise � jour ---
                    Enregistrement.Update
                    
                End If
            Next a
       
       End If

        '--- fermeture des enregistrements / connexions ---
        Select Case Enregistrement.State
            Case adStateClosed
            Case Else: Enregistrement.Close
        End Select
        ConnexionBDAnodisationSQL.Close
        
        '--- effacement des objets ---
        Set Enregistrement = Nothing
        Set ConnexionBDAnodisationSQL = Nothing
        
    End If
        
    Exit Function
    
GestionErreurs:
    
    '--- valeur de retour ---
    ConstructionImpressionAlarmesLigne = CStr(Err.Number)
    
    '--- fermeture de l'enregistrement / connexion ---
    On Error Resume Next
    Enregistrement.Close
    Set Enregistrement = Nothing
    ConnexionBDAnodisationSQL.Close
    Set ConnexionBDAnodisationSQL = Nothing

End Function

'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' R�le      : Construction de l'impression des gammes d'anodisation de production
' Entr�es :                                                NumFicheProduction -> Num�ro de la fiche de production
' Retours : ConstructionImpressionGammesProduction -> "" = pas d'incident sinon num�ro de l'erreur
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function ConstructionImpressionGammesProduction(ByVal NumFicheProduction As String) As String
    
    '--- aiguillage en cas d'erreurs ---
    On Error GoTo GestionErreurs

    '--- constantes priv�es ---
    
    '--- d�claration ---
    Dim a As Integer
    Dim TempsEnSecondes  As Long
    Dim DateDuJour As String, _
           Texte As String
    Dim ConnexionBDAnodisationSQL As ADODB.Connection
    Dim Enregistrement As ADODB.Recordset
    
    '--- affectation ---
    ConstructionImpressionGammesProduction = ""
    
    If NumFicheProduction <> "" Then
    
        '--- ouverture de la connexion ---
        Set ConnexionBDAnodisationSQL = New ADODB.Connection
        With ConnexionBDAnodisationSQL
            .ConnectionString = PARAMETRES_CONNEXION_BD_ANODISATION_SQL
            .CursorLocation = adUseServer
            .Mode = adModeReadWrite
            .ConnectionTimeout = 2     'X secondes d'attente de connexion avant de lancer un message d'erreur
            .Open
        End With
            
        '--- ouverture de la table ---
        Set Enregistrement = New ADODB.Recordset
        With Enregistrement
            .CursorLocation = adUseServer
            .CursorType = adOpenKeyset
            .LockType = adLockOptimistic
            .Open TABLE_IMP_GAMMES_ANODISATION_PRODUCTION_1, ConnexionBDAnodisationSQL, , adCmdTable
            .MoveFirst
        End With

        '--- enregistrement du n� de la fiche de production ---
        Enregistrement("NumFicheProduction") = NumFicheProduction
        
        '--- enregistrement de la date du jour ---
        DateDuJour = Format(Now, "dd/mm/yyyy")
        Enregistrement("DateDuJour") = DateDuJour
        
        '--- extraction des donn�es de la production ---
        If RechercheDetailsChargesProduction(NumFicheProduction) = TROUVE Then
            
            '--- enregistrement des valeurs ---
            Enregistrement("DateEntreeEnLigne") = Format(TTempEnrDetailsChargesProduction(1).DateEntreeEnLigne, "dd/mm/yyyy � hh:nn:ss")
            Enregistrement("ChargePrioritaire") = IIf(TTempEnrDetailsChargesProduction(1).ChargePrioritaire = True, "OUI", "NON")
            Enregistrement("NumGammeAnodisation") = TTempEnrDetailsChargesProduction(1).NumGammeAnodisation
            
        Else
        
            '--- affectation ---
            Enregistrement("DateEntreeEnLigne") = ""
            Enregistrement("ChargePrioritaire") = ""
            Enregistrement("NumGammeAnodisation") = ""
        
        End If
        
        '--- mise � jour ---
        Enregistrement.Update
        
        '--- fermeture des enregistrements ---
        Enregistrement.Close
        Set Enregistrement = Nothing
        
        '********************************************************************************************************************
        '                                                               CONTRUCTION DES DETAILS
        '********************************************************************************************************************
        
        '--- effacement de la table ---
        Set Enregistrement = New ADODB.Recordset
        With Enregistrement
            .CursorLocation = adUseServer
            .CursorType = adOpenKeyset
            .LockType = adLockOptimistic
            .Open "DELETE FROM " & TABLE_IMP_DETAILS_GAMMES_ANODISATION_PRODUCTION_1, ConnexionBDAnodisationSQL, , adCmdText
        End With

        '--- ouverture de la table ---
        Set Enregistrement = New ADODB.Recordset
        With Enregistrement
            .CursorLocation = adUseServer
            .CursorType = adOpenKeyset
            .LockType = adLockOptimistic
            .Open TABLE_IMP_DETAILS_GAMMES_ANODISATION_PRODUCTION_1, ConnexionBDAnodisationSQL, , adCmdTable
        End With

        '--- construction de la table des d�tails ---
        If RechercheDetailsGammesProduction(NumFicheProduction) = TROUVE Then
            For a = LBound(TTempEnrDetailsGammesProduction()) To UBound(TTempEnrDetailsGammesProduction())
                With TTempEnrDetailsGammesProduction(a)

                    '--- nouvel enregistrement ---
                    Enregistrement.AddNew

                    '--- construction de la fiche ---
                    Enregistrement("NumFicheProduction") = NumFicheProduction
                    Enregistrement("NumLigne") = .NumLigne
                    
                    If .NumZone > 0 Then

                        '--- code et libell� de la zone ---
                        Enregistrement("CodeZone") = TZones(.NumZone).Codezone
                        Enregistrement("LibelleZone") = TZones(.NumZone).LibelleZone
                    
                        '--- nom du poste r�el ---
                        If .NumPosteReel >= POSTES.P_CHGT_1 And .NumPosteReel <= DERNIER_POSTE Then
                            Enregistrement("NomPosteReel") = TEtatsPostes(.NumPosteReel).DefinitionPoste.NomPoste
                        End If
                        
                        '--- temps au poste et �gouttage ---
                        Enregistrement("TempsAuPosteTexte") = .TempsAuPosteTexte
                        Enregistrement("TempsEgouttageTexte") = .TempsEgouttageTexte
                        
                        '--- d�compte du temps r�el en HH:MM:SS ---
                        If .DecompteDuTempsAuPosteReelSecondes = "" Then
                            Enregistrement("DecompteDuTempsAuPosteReelTexte") = "-"
                        Else
                            If IsNumeric(.DecompteDuTempsAuPosteReelSecondes) = True Then
                                Enregistrement("DecompteDuTempsAuPosteReelTexte") = CTemps(CLng(.DecompteDuTempsAuPosteReelSecondes))
                            Else
                                Enregistrement("DecompteDuTempsAuPosteReelTexte") = "-"
                            End If
                        End If
                    
                    End If

                    '--- mise � jour ---
                    Enregistrement.Update

                End With
            Next a
        End If
        
        '--- fermeture des enregistrements / connexions ---
        Select Case Enregistrement.State
            Case adStateClosed
            Case Else: Enregistrement.Close
        End Select
        ConnexionBDAnodisationSQL.Close
        
        '--- effacement des objets ---
        Set Enregistrement = Nothing
        Set ConnexionBDAnodisationSQL = Nothing
        
    End If
        
    Exit Function
    
GestionErreurs:
    
    '--- valeur de retour ---
    ConstructionImpressionGammesProduction = CStr(Err.Number)

    '--- fermeture de l'enregistrement / connexion ---
    On Error Resume Next
    Enregistrement.Close
    Set Enregistrement = Nothing
    ConnexionBDAnodisationSQL.Close
    Set ConnexionBDAnodisationSQL = Nothing

End Function

'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' R�le      : Construction de l'impression des d�tails d'une charge
' Entr�es :                                  NumFicheProduction -> Num�ro de la fiche de production
'                                              NumCommandeInterne -> Num�ro de commande interne
' Retours : ConstructionImpressionTracabiliteCharge -> "" = pas d'incident sinon num�ro de l'erreur
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function ConstructionImpressionDetailsCharge(ByVal NumFicheProduction As String, _
                                                                                        ByVal NumCommandeInterne As Long) As String
    
    '--- aiguillage en cas d'erreurs ---
    On Error GoTo GestionErreurs

    '--- constantes priv�es ---
    
    '--- d�claration ---
    Dim a As Integer, _
           b As Integer, _
           NbrLignesReferencesClient As Integer
    Dim TempsEnSecondes  As Long
    Dim DateDuJour As String, _
           Texte As String, _
           LesReferencesClient As String
    Dim ConnexionBDAnodisationSQL As ADODB.Connection
    Dim Enregistrement As ADODB.Recordset, _
           Enregistrement2 As ADODB.Recordset
    Dim TReferencesClient As Variant
    
    '--- affectation ---
    ConstructionImpressionDetailsCharge = ""
    
    If NumFicheProduction <> "" Then
    
        '--- ouverture de la connexion ---
        Set ConnexionBDAnodisationSQL = New ADODB.Connection
        With ConnexionBDAnodisationSQL
            .ConnectionString = PARAMETRES_CONNEXION_BD_ANODISATION_SQL
            .CursorLocation = adUseServer
            .Mode = adModeReadWrite
            .ConnectionTimeout = 2     'X secondes d'attente de connexion avant de lancer un message d'erreur
            .Open
        End With
            
        '--- ouverture de la table ---
        Set Enregistrement = New ADODB.Recordset
        With Enregistrement
            .CursorLocation = adUseServer
            .CursorType = adOpenKeyset
            .LockType = adLockOptimistic
            .Open TABLE_IMP_DETAILS_CHARGE_1, ConnexionBDAnodisationSQL, , adCmdTable
            .MoveFirst
        End With

        '--- enregistrement du n� de la fiche de production ---
        Enregistrement("NumFicheProduction") = NumFicheProduction
        
        '--- enregistrement de la date du jour ---
        DateDuJour = Format(Now, "dd/mm/yyyy")
        Enregistrement("DateDuJour") = DateDuJour
        
        '--- extraction des donn�es de la production ---
        If RechercheDetailsChargesProduction(NumFicheProduction) = TROUVE Then
            
            '--- enregistrement des valeurs ---
            Enregistrement("DateEntreeEnLigne") = Format(TTempEnrDetailsChargesProduction(1).DateEntreeEnLigne, "dd/mm/yyyy � hh:nn:ss")
            Enregistrement("ChargePrioritaire") = IIf(TTempEnrDetailsChargesProduction(1).ChargePrioritaire = True, "OUI", "NON")
            
        Else
        
            '--- affectation ---
            Enregistrement("DateEntreeEnLigne") = ""
            Enregistrement("ChargePrioritaire") = ""
        
        End If
        
        '--- mise � jour ---
        Enregistrement.Update
        
        '--- fermeture des enregistrements ---
        Enregistrement.Close
        Set Enregistrement = Nothing
        
        '********************************************************************************************************************
        '                                CONTRUCTION DES DETAILS ET DES REFERENCES DU CLIENT
        '********************************************************************************************************************
        
        '--- effacement de la table (d�tails de la charge) ---
        Set Enregistrement = New ADODB.Recordset
        With Enregistrement
            .CursorLocation = adUseServer
            .CursorType = adOpenKeyset
            .LockType = adLockOptimistic
            .Open "DELETE FROM " & TABLE_IMP_DETAILS_DETAILS_CHARGE_1, ConnexionBDAnodisationSQL, , adCmdText
        End With
        
        '--- ouverture de la table (d�tails de la charge) ---
        Set Enregistrement = New ADODB.Recordset
        With Enregistrement
            .CursorLocation = adUseServer
            .CursorType = adOpenKeyset
            .LockType = adLockOptimistic
            .Open TABLE_IMP_DETAILS_DETAILS_CHARGE_1, ConnexionBDAnodisationSQL, , adCmdTable
        End With
        
        '--- effacement de la table (d�tails des r�f�rences clients) ---
        Set Enregistrement2 = New ADODB.Recordset
        With Enregistrement2
            .CursorLocation = adUseServer
            .CursorType = adOpenKeyset
            .LockType = adLockOptimistic
            .Open "DELETE FROM " & TABLE_IMP_DETAILS_REFERENCES_CLIENTS_1, ConnexionBDAnodisationSQL, , adCmdText
        End With
        
        '--- ouverture de la table (d�tails des r�f�rences clients) ---
        Set Enregistrement2 = New ADODB.Recordset
        With Enregistrement2
            .CursorLocation = adUseServer
            .CursorType = adOpenKeyset
            .LockType = adLockOptimistic
            .Open TABLE_IMP_DETAILS_REFERENCES_CLIENTS_1, ConnexionBDAnodisationSQL, , adCmdTable
        End With
        
        '--- construction de la table des d�tails ---
        If RechercheDetailsChargesProduction(NumFicheProduction) = TROUVE Then
            For a = LBound(TTempEnrDetailsChargesProduction()) To UBound(TTempEnrDetailsChargesProduction())
                With TTempEnrDetailsChargesProduction(a)
                    
                    '--- prendre toutes les fiches d'atelier si NumCommandeInterne = "" sinon
                    '    ne prendre que le num�ro de commande interne recherch� ---
                    If NumCommandeInterne = 0 Or .NumCommandeInterne = NumCommandeInterne Then
                    
                        '--- nouvel enregistrement ---
                        Enregistrement.AddNew
                        
                        '--- construction de la fiche ---
                        Enregistrement("NumFicheProduction") = NumFicheProduction
                        Enregistrement("NumCommandeInterne") = .NumCommandeInterne
                        Enregistrement("NumLigne") = .NumLigne
                        Enregistrement("CodeClient") = .CodeClient
                        Enregistrement("NbrPieces") = .NbrPieces
                        Enregistrement("Designation") = .Designation
                        
                        'Enregistrement("Matiere") = .Matiere
                    
                        '--- mise � jour ---
                        Enregistrement.Update
                    
                        '--- enregistrement des r�f�rences clients ---
                        If .NumLignesReferencesClient = "" Then
                            
                            '--- nouvel enregistrement / construction de la fiche / mise � jour ---
                            Enregistrement2.AddNew
                            Enregistrement2("NumCommandeInterne") = .NumCommandeInterne
                            Enregistrement2("ReferencesClient") = "Totalit� de la commande interne (" & .NbrPieces & "/" & .NbrPieces & ")"
                            Enregistrement2.Update
                        
                        Else
                        
                            '--- recherche des r�f�rences du client ---
                            'LesReferencesClient = ExtraitReferencesClient(.NumCommandeInterne, _
                                                                                                         .NumLignesReferencesClient, _
                                                                                                          NbrLignesReferencesClient)
                            'TReferencesClient = Split(LesReferencesClient, vbCr)

                            '--- enregistrement des r�f�rences du client ---
                            'For b = LBound(TReferencesClient) To UBound(TReferencesClient)
                            '    Enregistrement2.AddNew
                            '    Enregistrement2("NumCommandeInterne") = .NumCommandeInterne
                            '    Enregistrement2("ReferencesClient") = TReferencesClient(b)
                            '    Enregistrement2.Update
                            'Next b
                    
                        End If
                    
                    End If
                
                End With
            Next a
        End If
        
        '--- fermeture des enregistrements / connexions ---
        Select Case Enregistrement.State
            Case adStateClosed
            Case Else: Enregistrement.Close
        End Select
        Select Case Enregistrement2.State
            Case adStateClosed
            Case Else: Enregistrement2.Close
        End Select
        ConnexionBDAnodisationSQL.Close
        
        '--- effacement des objets ---
        Set Enregistrement = Nothing
        Set Enregistrement2 = Nothing
        Set ConnexionBDAnodisationSQL = Nothing
        
    End If
        
    Exit Function
    
GestionErreurs:
    
    '--- valeur de retour ---
    ConstructionImpressionDetailsCharge = CStr(Err.Number)
    
    '--- fermeture de l'enregistrement / connexion ---
    On Error Resume Next
    Enregistrement.Close
    Set Enregistrement = Nothing
    ConnexionBDAnodisationSQL.Close
    Set ConnexionBDAnodisationSQL = Nothing

End Function



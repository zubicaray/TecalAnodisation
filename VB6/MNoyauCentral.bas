Attribute VB_Name = "MNoyauCentral"
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' R�le                    : MODULE CONTENANT LES ROUTINES DU NOYAU CENTRAL
' Nom                    : MNoyauCentral.bas
' Date de cr�ation : 31/07/2000
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

'--- d�clarations obligatoires ---
Option Explicit

'--- options g�n�rales ---
Option Base 1
DefVar A-Z

'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' R�le        : Noyau central du programme multi-t�ches
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Sub NoyauCentral()
    
    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next
    
    '--- d�claration ---
    Static PassageGestionCommandesOperateur As Boolean
    Dim TempsNoyauCentral As Long
    Dim InstantX As Date
    
    '--- r�f�renciel de temps ---
    InstantX = Now
    Maintenant = Format(InstantX, "yyyymmddhhnnss")
    DateMaintenant = Format(InstantX, "dd/mm/yyyy")
    HeureMaintenant = Format(InstantX, "hh:nn:ss")
    AnneesMaintenant = Format(InstantX, "yyyy")
    MoisMaintenant = Format(InstantX, "mm")
    JoursMaintenant = Format(InstantX, "dd")
    HeuresMaintenant = Format(InstantX, "hh")
    MinutesMaintenant = Format(InstantX, "nn")
    SecondesMaintenant = Format(InstantX, "ss")

    

    '--- indicateur ---
    With OccFPrincipale.LTempsNoyauCentral
        .Caption = "-"
        .BackColor = COULEURS.VERT_3
        .ForeColor = COULEURS.NOIR
        .Refresh
    End With
    
    '--- entretien des graphes de production ---
    If JoursMaintenant = 1 And MoisMaintenant = 1 And HeuresMaintenant = 1 And MinutesMaintenant = 1 Then
        EntretienGraphesProduction = True
    End If
    
    '--- analyse du programmateur cyclique ---
    AnalyseProgrammateurCyclique
        
    '--- analyse de toutes les cuves ---
    AnalyseCuves
        
    '--- analyse des temps de mouvements ---
    'AnalyseTempsMouvements                             'volontairement en commentaire car l'apprentissage est termin�
    
    '--- analyse des charges en ligne pour les PONTS et POSTES ---
    AnalyseChargesEnLignePonts
    AnalyseChargesEnLignePostes
    
    '--- analyse de la fin de cycle de l'�tuve ---
    AnalyseFinDeCycleEtuve
    
    '--- commandes de l'op�rateur / gestion des gammes par IA ---
    If TEtatsPonts(PONTS.P_1).ControleParOperateur = True And _
        TEtatsPonts(PONTS.P_2).ControleParOperateur = True Then
        
        '--- RAZ du tableau des commandes op�rateurs au moment du passage en commandes op�rateur ---
        If PassageGestionCommandesOperateur = False Then
            Erase TCommandesOperateur
            PassageGestionCommandesOperateur = True
        End If
        
        '--- gestion des commandes de l'op�rateur ---
        'uniquement lorsque les 2 ponts sont sous le contr�le de l'op�rateur
        GestionCommandesOperateur
        Call Log("gestion des commandes de l'op�rateur ")
    Else
       
        '--- RAZ de la variable de passage en gestion des commandes op�rateur ---
        PassageGestionCommandesOperateur = False
        
        '--- appel du moteur d'inf�rence ---
        MoteurInference
    
    End If
  
    '--- tra�abilit� des redresseurs ---
    EffectueTra�abiliteRedresseurs
  
    '--- signalisation des d�fauts sur le gyrophare et le klaxon ---
    SignalisationDefautsGyrophareKlaxonVersAPI
    
    '--- affichage du temps de d�roulement du noyau central ---
    TempsNoyauCentral = DateDiff("s", InstantX, Now)
    With OccFPrincipale.LTempsNoyauCentral
        Select Case TempsNoyauCentral
            Case 0
                If .Caption <> "<1" Then
                    .Caption = "<1"
                    .BackColor = COULEURS.VERT_3
                    .ForeColor = COULEURS.NOIR
                    .Refresh
                End If
            Case 1 To 9
                If .Caption <> ">" & CStr(TempsNoyauCentral) Then
                    .Caption = ">" & CStr(TempsNoyauCentral)
                    Select Case TempsNoyauCentral
                        Case 1 To 2: .BackColor = COULEURS.VERT_3
                        Case 3 To 4: .BackColor = COULEURS.ORANGE_3
                        Case 5 To 9: .BackColor = COULEURS.ROUGE_3
                        Case Else
                    End Select
                    .ForeColor = COULEURS.NOIR
                    .Refresh
                End If
            Case 10 To 99
                If .Caption <> CStr(TempsNoyauCentral) Then
                    .Caption = CStr(TempsNoyauCentral)
                    .BackColor = COULEURS.ROUGE_3
                    .ForeColor = COULEURS.JAUNE_3
                    .Refresh
                End If
            Case Else
                If .Caption <> "-" Then
                    .Caption = "-"
                    .BackColor = COULEURS.ROUGE_3
                    .ForeColor = COULEURS.JAUNE_3
                    .Refresh
                End If
        End Select
        
    End With
        
    '--- affectation ---
    PremierPassageNoyauCentral = True
    
End Sub

'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' R�le        : Analyse du programmateur cyclique au moment du changement de date
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Sub AnalyseProgrammateurCyclique()
    
    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next

    '--- d�claration ---
    Dim a As Integer, b As Integer, c As Integer
    Dim TypeDeJournee As Integer
    Dim DateATraiter As Variant

    '--- analyse en fonction du PC ---
    'If TypePC <> TYPES_PC. Then Exit Sub

    '--- ne pas lancer l'analyse si la fen�tre du programmateur cyclique est ouverte _
         pour �viter l'�crasement au passage de minuit ---
    If FProgCycliqueChargee = True Then Exit Sub

    If MemDateProgCyclique <> DateMaintenant Then

        '--- d�calage du programmateur cyclique ---
        For a = 2 To NBR_JOURS_PROG_CYCLIQUE
            For b = CUVES_REGULATION.C_C00 To DERNIERE_CUV_REGULATION
                With TProgCyclique(Pred(a), b)

                    '--- transfert ---
                    .TypeDeJournee = TProgCyclique(a, b).TypeDeJournee
                    For c = 1 To NBR_TOPS_POSSIBLES
                        .TTopsDebutPompe(c) = TProgCyclique(a, b).TTopsDebutPompe(c)
                        .TTopsFinPompe(c) = TProgCyclique(a, b).TTopsFinPompe(c)
                        .TCyclesPompe(c) = TProgCyclique(a, b).TCyclesPompe(c)
                        .TTopsDebutChauffage(c) = TProgCyclique(a, b).TTopsDebutChauffage(c)
                        .TTopsFinChauffage(c) = TProgCyclique(a, b).TTopsFinChauffage(c)
                        .TModesChauffage(c) = TProgCyclique(a, b).TModesChauffage(c)
                    Next c

                    '--- vidage du dernier jour ---
                    If a = NBR_JOURS_PROG_CYCLIQUE Then
                        TProgCyclique(a, b).TypeDeJournee = JOURNEES_TYPES.J_ARRET
                        For c = 1 To NBR_TOPS_POSSIBLES
                            TProgCyclique(a, b).TTopsDebutPompe(c) = ""
                            TProgCyclique(a, b).TTopsFinPompe(c) = ""
                            TProgCyclique(a, b).TCyclesPompe(c) = CYCLES_POMPES.CP_ARRET
                            TProgCyclique(a, b).TTopsDebutChauffage(c) = ""
                            TProgCyclique(a, b).TTopsFinChauffage(c) = ""
                            TProgCyclique(a, b).TModesChauffage(c) = MODES_PRODUCTION.M_ARRET
                        Next c
                    End If

                End With
            Next b
        Next a

        '--- contr�le de la validit� des dates ---
        For a = 1 To NBR_JOURS_PROG_CYCLIQUE
            For b = CUVES_REGULATION.C_C00 To DERNIERE_CUV_REGULATION
                With TProgCyclique(a, b)

                    '--- affectation ---
                    DateATraiter = DateAdd("d", Pred(a), DateMaintenant)

                    '--- type de journ�e ---
                    Select Case Weekday(DateATraiter)
                        Case vbMonday: TypeDeJournee = JOURNEES_TYPES.J_TRAVAIL
                        Case vbTuesday: TypeDeJournee = JOURNEES_TYPES.J_TRAVAIL
                        Case vbWednesday: TypeDeJournee = JOURNEES_TYPES.J_TRAVAIL
                        Case vbThursday: TypeDeJournee = JOURNEES_TYPES.J_TRAVAIL
                        Case vbFriday: TypeDeJournee = JOURNEES_TYPES.J_TRAVAIL
                        Case vbSaturday: TypeDeJournee = JOURNEES_TYPES.J_VEILLE
                        Case vbSunday: TypeDeJournee = JOURNEES_TYPES.J_REPRISE
                        Case Else: TypeDeJournee = JOURNEES_TYPES.J_ARRET
                    End Select

                    '--- affectation ---
                    DateATraiter = Format(DateATraiter, "yyyymmdd")

                    '--- contr�le avec la date ---
                    If DateATraiter = Left(.TTopsDebutChauffage(1), 8) Then
                        Exit For
                    Else

                        '--- type de journ�e ---
                        .TypeDeJournee = TypeDeJournee

                        '--- transfert des nouvelles valeurs ---
                        For c = 1 To NBR_TOPS_POSSIBLES

                            '--- pompe ---
                            .TTopsDebutPompe(c) = TJourneesTypes(b, TypeDeJournee).TTopsDebutPompe(c)
                            If Left(.TTopsDebutPompe(c), 1) = "X" Then
                                .TTopsDebutPompe(c) = DateATraiter + Mid(.TTopsDebutPompe(c), 9)
                            End If
                            .TTopsFinPompe(c) = TJourneesTypes(b, TypeDeJournee).TTopsFinPompe(c)
                            If Left(.TTopsFinPompe(c), 1) = "X" Then
                                .TTopsFinPompe(c) = DateATraiter + Mid(.TTopsFinPompe(c), 9)
                            End If
                            .TCyclesPompe(c) = TJourneesTypes(b, TypeDeJournee).TCyclesPompe(c)

                            '--- chauffage ---
                            .TTopsDebutChauffage(c) = TJourneesTypes(b, TypeDeJournee).TTopsDebutChauffage(c)
                            If Left(.TTopsDebutChauffage(c), 1) = "X" Then
                                .TTopsDebutChauffage(c) = DateATraiter + Mid(.TTopsDebutChauffage(c), 9)
                            End If
                            .TTopsFinChauffage(c) = TJourneesTypes(b, TypeDeJournee).TTopsFinChauffage(c)
                            If Left(.TTopsFinChauffage(c), 1) = "X" Then
                                .TTopsFinChauffage(c) = DateATraiter + Mid(.TTopsFinChauffage(c), 9)
                            End If
                            .TModesChauffage(c) = TJourneesTypes(b, TypeDeJournee).TModesChauffage(c)

                        Next c

                    End If

                End With
            Next b
        Next a

        '--- sauvegarde du programmateur cyclique ---
        SauveProgCyclique

        '--- affectation ---
        MemDateProgCyclique = DateMaintenant

        '--- sauvegarde de la configuration ---
        SauveConfiguration

    End If

End Sub

'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' R�le        : Analyse de toutes les cuves
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Sub AnalyseCuves()
    
    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next

    '--- d�claration ---
    Dim a As Integer, _
           b As Integer
    
    '--- analyse en fonction du PC ---
    'If TypePC <> TYPES_PC. Then Exit Sub

    '--- marche automatique des pompes et des chauffages ---
    For a = CUVES_REGULATION.C_C00 To DERNIERE_CUV_REGULATION
        With TProgCyclique(1, a)

            '--- analyse pour la pompe (si cuve avec pompe) ---
            If TEtatsCuves(a).DefinitionCuve.PresencePompe = True Then
            
                For b = 1 To NBR_TOPS_POSSIBLES
                    If Maintenant >= Val(.TTopsDebutPompe(b)) And Maintenant <= Val(.TTopsFinPompe(b)) Then
                                
                        '--- affectation et transfert vers l'automate ---
                        TEtatsCuves(a).CyclePompe = .TCyclesPompe(b)
                        AutomatiquePompe a

                        Exit For

                    End If
                Next b

            End If

            '--- analyse pour le chauffage ---
            For b = 1 To NBR_TOPS_POSSIBLES
                If Maintenant >= Val(.TTopsDebutChauffage(b)) And Maintenant <= Val(.TTopsFinChauffage(b)) Then
                    
                    Select Case .TModesChauffage(b)
                        Case MODES_PRODUCTION.M_ARRET To MODES_PRODUCTION.M_PRODUCTION
                            TEtatsCuves(a).ModeProduction = .TModesChauffage(b)
                        Case Else
                    End Select
                    
                    '--- transfert vers l'automate ---
                    AutomatiqueChauffage a

                    Exit For

                End If
            Next b

        End With

    Next a
    
End Sub

'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' R�le        : Initialise automatiquement divers �l�ments de la ligne sur une marche g�n�rale
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Sub InitialisationSurMarcheGenerale()

    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next

    '--- constantes priv�es ---
    
    '--- d�claration ---
    Dim a As Integer                        'pour les boucles FOR...NEXT

End Sub

'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' R�le        : Analyse de tous les temps de mouvements (apprentissage pour renseigner le moteur d'inf�rence)
'                   les temps serviront aux pr�misses pour connaitre le temps exacte d'un transfert entre un poste
'                   de d�part et un poste d'arriv�e
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Sub AnalyseTempsMouvements()
    
    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next

    '--- constantes priv�es ---
    
    '--- d�claration ---
    Static AntiRebondMemorisationMouvements As Boolean
        
    Dim a As Integer, _
           b As Integer, _
           NumPoste As Integer
    Dim TempsMouvementSecondes As Single                                          'temps d'un mouvement en secondes
    Static TPostesDepart(PONTS.P_1 To PONTS.P_2)  As Long, _
              TPostesArrivee(PONTS.P_1 To PONTS.P_2)  As Long
    
    'Static TAnalyseOuverturesCouvercles(CUVES_API.C_A1 To CUVES_API.C_C15) As VarAnalyseMouvements
    'Static TAnalyseFermeturesCouvercles(CUVES_API.C_A1 To CUVES_API.C_C15) As VarAnalyseMouvements
    
    Static TAnalyseAccrochesChargeVersHaut(PONTS.P_1 To PONTS.P_2)  As VarAnalyseMouvements
    Static TAnalyseAccrochesChargeVersBas(PONTS.P_1 To PONTS.P_2)  As VarAnalyseMouvements

    Static TAnalyseDescenteHautVersBas(PONTS.P_1 To PONTS.P_2)  As VarAnalyseMouvements
    Static TAnalyseDescenteIntermediaireVersBas(PONTS.P_1 To PONTS.P_2)  As VarAnalyseMouvements
        
    Static TAnalyseMonteeBasVersIntermediaire(PONTS.P_1 To PONTS.P_2)  As VarAnalyseMouvements
    Static TAnalyseMonteeBasVersHaut(PONTS.P_1 To PONTS.P_2)  As VarAnalyseMouvements
        
    Static TAnalyseTranslation(PONTS.P_1 To PONTS.P_2)  As VarAnalyseMouvements
    
    '--- analyse en fonction du PC ---
    'If TypePC <> TYPES_PC. Then Exit Sub
    
    '----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    
    '********************************************************************************************
    '*               MEMORISATION DE L'ENSEMBLE DES TEMPS DE MOUVEMENTS
    '********************************************************************************************
    If Hour(Now) = 23 Then                      'm�morisation � 23H00 avec enclenchement de l'anti-rebond
        If AntiRebondMemorisationMouvements = False Then
            Bidon = EnregistrementTempsMouvements
            AntiRebondMemorisationMouvements = True
        End If
    Else
        AntiRebondMemorisationMouvements = False
    End If

    '----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    
    '--- l'analyse se fait uniquement si les ponts sont en automatique et aucun contr�le par op�rateur ---
    If TEtatsPonts(PONTS.P_1).ModePont = MODES_PONTS.M_AUTOMATIQUE And _
       TEtatsPonts(PONTS.P_2).ModePont = MODES_PONTS.M_AUTOMATIQUE Then
            If TEtatsPonts(PONTS.P_1).ControleParOperateur = True Or TEtatsPonts(PONTS.P_2).ControleParOperateur = True Then
                Exit Sub
            End If
    Else
        Exit Sub
    End If
    
    '----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    
    '*********************************************************************************************************************
    '*                                                         GESTION POUR LES COUVERCLES
    '*********************************************************************************************************************

'    For a = LBound(TEtatsCuves()) To UBound(TEtatsCuves())
'
'        With TEtatsCuves(a)
'
'            '--- recherche du poste pour les couvercles ---
'            NumPoste = CorrespondanceCuvesAPIPostes(a)
'
'            '--- couvercles ---
'            If NumPoste > 0 Then
'
'                With TEtatsPostes(NumPoste)
'
'                    If .DefinitionPoste.PresenceCouvercles = True Then
'
'                        '********************************************************************************************
'                        '*                                             OUVERTURE DES COUVERCLES
'                        '********************************************************************************************
'
'                        '--- initialisation de la s�quence ---
'                        If .EtatsCouvercles = ETATS_COUVERCLES.E_COUVERCLES_FERMES Then
'                            'TAnalyseOuverturesCouvercles(a).EtatMouvement = ETATS_MOUVEMENTS.E_PAS_DE_MOUVEMENT
'                        ElseIf .EtatsCouvercles = ETATS_COUVERCLES.E_DISJONCTION_EV_COUVERCLES Then
'                            'TAnalyseOuverturesCouvercles(a).DateDebutMouvement = Empty
'                            'TAnalyseOuverturesCouvercles(a).DateFinMouvement = Empty
'                        End If
'
'                        '--- analyse des mouvements ---
'                        Select Case TAnalyseOuverturesCouvercles(a).EtatMouvement
'
'                            Case ETATS_MOUVEMENTS.E_PAS_DE_MOUVEMENT
'                                '--- pas de mouvement ---
'                                If .EtatsCouvercles = ETATS_COUVERCLES.E_COUVERCLES_FERMES Then
'                                    TAnalyseOuverturesCouvercles(a).DateDebutMouvement = Now
'                                End If
'                                If .EtatsCouvercles = ETATS_COUVERCLES.E_COUVERCLES_EN_OUVERTURE Then
'                                    TAnalyseOuverturesCouvercles(a).EtatMouvement = ETATS_MOUVEMENTS.E_MOUVEMENT_EN_COURS
'                                End If
'
'                            Case ETATS_MOUVEMENTS.E_MOUVEMENT_EN_COURS
'                                '--- mouvement en cours ---
'                                If .EtatsCouvercles = ETATS_COUVERCLES.E_COUVERCLES_OUVERTS Then
'                                    TAnalyseOuverturesCouvercles(a).DateFinMouvement = Now
'                                    TAnalyseOuverturesCouvercles(a).EtatMouvement = ETATS_MOUVEMENTS.E_FIN_DU_MOUVEMENT
'                                End If
'
'                            Case ETATS_MOUVEMENTS.E_FIN_DU_MOUVEMENT
'                                '--- fin du mouvement (calcul du temps du mouvement) ---
'                                With TAnalyseOuverturesCouvercles(a)
'                                    If .DateDebutMouvement <> Empty And .DateFinMouvement <> Empty Then
'                                        TempsMouvementSecondes = DateDiff("s", .DateDebutMouvement, .DateFinMouvement)
'                                        If TempsMouvementSecondes > 0 Then
'                                            TEtatsCuves(a).TTempsMouvements.TempsOuvertureCouvercles = TempsMouvementSecondes
'                                            .DateDebutMouvement = Empty
'                                            .DateFinMouvement = Empty
'                                        End If
'                                    End If
'                                End With
'
'                            Case Else
'
'                        End Select
'
'                        '********************************************************************************************
'                        '*                                          FERMETURE DES COUVERCLES
'                        '********************************************************************************************
'
'                        '--- initialisation de la s�quence ---
'                        If .EtatsCouvercles = ETATS_COUVERCLES.E_COUVERCLES_OUVERTS Then
'                            TAnalyseFermeturesCouvercles(a).EtatMouvement = ETATS_MOUVEMENTS.E_PAS_DE_MOUVEMENT
'                        ElseIf .EtatsCouvercles = ETATS_COUVERCLES.E_DISJONCTION_EV_COUVERCLES Then
'                            TAnalyseFermeturesCouvercles(a).DateDebutMouvement = Empty
'                            TAnalyseFermeturesCouvercles(a).DateFinMouvement = Empty
'                        End If
'
'                        '--- analyse des mouvements ---
'                        Select Case TAnalyseFermeturesCouvercles(a).EtatMouvement
'
'                            Case ETATS_MOUVEMENTS.E_PAS_DE_MOUVEMENT
'                                '--- pas de mouvement ---
'                                If .EtatsCouvercles = ETATS_COUVERCLES.E_COUVERCLES_OUVERTS Then
'                                    TAnalyseFermeturesCouvercles(a).DateDebutMouvement = Now
'                                End If
'                                If .EtatsCouvercles = ETATS_COUVERCLES.E_COUVERCLES_EN_FERMETURE Then
'                                    TAnalyseFermeturesCouvercles(a).EtatMouvement = ETATS_MOUVEMENTS.E_MOUVEMENT_EN_COURS
'                                End If
'
'                            Case ETATS_MOUVEMENTS.E_MOUVEMENT_EN_COURS
'                                '--- mouvement en cours ---
'                                If .EtatsCouvercles = ETATS_COUVERCLES.E_COUVERCLES_FERMES Then
'                                    TAnalyseFermeturesCouvercles(a).DateFinMouvement = Now
'                                    TAnalyseFermeturesCouvercles(a).EtatMouvement = ETATS_MOUVEMENTS.E_FIN_DU_MOUVEMENT
'                                End If
'
'                            Case ETATS_MOUVEMENTS.E_FIN_DU_MOUVEMENT
'                                '--- fin du mouvement (calcul du temps du mouvement) ---
'                                With TAnalyseFermeturesCouvercles(a)
'                                    If .DateDebutMouvement <> Empty And .DateFinMouvement <> Empty Then
'                                        TempsMouvementSecondes = DateDiff("s", .DateDebutMouvement, .DateFinMouvement)
'                                        If TempsMouvementSecondes > 0 Then
'                                            TEtatsCuves(a).TTempsMouvements.TempsFermetureCouvercles = TempsMouvementSecondes
'                                            .DateDebutMouvement = Empty
'                                            .DateFinMouvement = Empty
'                                        End If
'                                    End If
'                                End With
'
'                            Case Else
'
'                        End Select
'
'                    End If
'
'                End With
'
'            End If
'
'        End With
'
'    Next a
    
    '*********************************************************************************************************************
    '*                                                             GESTION POUR LES PONTS
    '*********************************************************************************************************************
    
    For a = LBound(TEtatsPonts()) To UBound(TEtatsPonts())
    
        With TEtatsPonts(a)
    
            '********************************************************************************************
            '                                                       MONTEE DES ACCROCHES
            '********************************************************************************************
            
            '--- initialisation de la s�quence ---
            If .EtatsAccrochesCharge = ETATS_ACCROCHES.E_ACCROCHES_EN_BAS Then
                TAnalyseAccrochesChargeVersHaut(a).EtatMouvement = ETATS_MOUVEMENTS.E_PAS_DE_MOUVEMENT
            End If
            
            '--- analyse des mouvements ---
            Select Case TAnalyseAccrochesChargeVersHaut(a).EtatMouvement
            
                Case ETATS_MOUVEMENTS.E_PAS_DE_MOUVEMENT
                    '--- pas de mouvement ---
                    If .EtatsAccrochesCharge = ETATS_ACCROCHES.E_ACCROCHES_EN_BAS Then
                        TAnalyseAccrochesChargeVersHaut(a).DateDebutMouvement = Now
                    End If
                    If .EtatsAccrochesCharge = ETATS_ACCROCHES.E_ACCROCHES_VERS_HAUT Then
                        TAnalyseAccrochesChargeVersHaut(a).EtatMouvement = ETATS_MOUVEMENTS.E_MOUVEMENT_EN_COURS
                    End If
        
                Case ETATS_MOUVEMENTS.E_MOUVEMENT_EN_COURS
                    '--- mouvement en cours ---
                    If .EtatsAccrochesCharge = ETATS_ACCROCHES.E_ACCROCHES_EN_HAUT Then
                        TAnalyseAccrochesChargeVersHaut(a).DateFinMouvement = Now
                        TAnalyseAccrochesChargeVersHaut(a).EtatMouvement = ETATS_MOUVEMENTS.E_FIN_DU_MOUVEMENT
                    End If
                
                Case ETATS_MOUVEMENTS.E_FIN_DU_MOUVEMENT
                    '--- fin du mouvement (calcul du temps du mouvement) ---
                    With TAnalyseAccrochesChargeVersHaut(a)
                        If .DateDebutMouvement <> Empty And .DateFinMouvement <> Empty Then
                            TempsMouvementSecondes = DateDiff("s", .DateDebutMouvement, .DateFinMouvement)
                            If TempsMouvementSecondes > 0 Then
                                TEtatsPonts(a).TTempsMouvements.TempsAccrochesChargeVersHaut = TempsMouvementSecondes
                                .DateDebutMouvement = Empty
                                .DateFinMouvement = Empty
                            End If
                        End If
                    End With
                
                Case Else
            
            End Select
            
            '********************************************************************************************
            '*                                           DESCENTE DES ACCROCHES
            '********************************************************************************************
            
            '--- initialisation de la s�quence ---
            If .EtatsAccrochesCharge = ETATS_ACCROCHES.E_ACCROCHES_EN_HAUT Then
                TAnalyseAccrochesChargeVersBas(a).EtatMouvement = ETATS_MOUVEMENTS.E_PAS_DE_MOUVEMENT
            End If
            
            '--- analyse des mouvements ---
            Select Case TAnalyseAccrochesChargeVersBas(a).EtatMouvement
            
                Case ETATS_MOUVEMENTS.E_PAS_DE_MOUVEMENT
                    '--- pas de mouvement ---
                    If .EtatsAccrochesCharge = ETATS_ACCROCHES.E_ACCROCHES_EN_HAUT Then
                        TAnalyseAccrochesChargeVersBas(a).DateDebutMouvement = Now
                    End If
                    If .EtatsAccrochesCharge = ETATS_ACCROCHES.E_ACCROCHES_VERS_BAS Then
                        TAnalyseAccrochesChargeVersBas(a).EtatMouvement = ETATS_MOUVEMENTS.E_MOUVEMENT_EN_COURS
                    End If
        
                Case ETATS_MOUVEMENTS.E_MOUVEMENT_EN_COURS
                    '--- mouvement en cours ---
                    If .EtatsAccrochesCharge = ETATS_ACCROCHES.E_ACCROCHES_VERS_BAS Then
                        TAnalyseAccrochesChargeVersBas(a).DateFinMouvement = Now
                        TAnalyseAccrochesChargeVersBas(a).EtatMouvement = ETATS_MOUVEMENTS.E_FIN_DU_MOUVEMENT
                    End If
                
                Case ETATS_MOUVEMENTS.E_FIN_DU_MOUVEMENT
                    '--- fin du mouvement (calcul du temps du mouvement) ---
                    With TAnalyseAccrochesChargeVersBas(a)
                        If .DateDebutMouvement <> Empty And .DateFinMouvement <> Empty Then
                            TempsMouvementSecondes = DateDiff("s", .DateDebutMouvement, .DateFinMouvement)
                            If TempsMouvementSecondes > 0 Then
                                TEtatsPonts(a).TTempsMouvements.TempsAccrochesChargeVersHaut = TempsMouvementSecondes
                                .DateDebutMouvement = Empty
                                .DateFinMouvement = Empty
                            End If
                        End If
                    End With
            
                Case Else
            
            End Select
    
            '********************************************************************************************
            '*                           DESCENTE DU NIVEAU HAUT VERS LE NIVEAU BAS
            '********************************************************************************************
            
            '--- initialisation de la s�quence ---
            If .NiveauActuel = NIVEAUX_PONTS.N_HAUT And .NiveauDestination = NIVEAUX_PONTS.N_HAUT Then
                TAnalyseDescenteHautVersBas(a).EtatMouvement = ETATS_MOUVEMENTS.E_PAS_DE_MOUVEMENT
            End If
            If .UnDefautAuMoinsSignale = True Then
                TAnalyseDescenteHautVersBas(a).DateDebutMouvement = Empty
                TAnalyseDescenteHautVersBas(a).DateFinMouvement = Empty
            End If
            
            '--- analyse des mouvements ---
            Select Case TAnalyseDescenteHautVersBas(a).EtatMouvement
            
                Case ETATS_MOUVEMENTS.E_PAS_DE_MOUVEMENT
                    '--- pas de mouvement ---
                    If .NiveauActuel = NIVEAUX_PONTS.N_HAUT And .NiveauDestination = NIVEAUX_PONTS.N_BAS Then
                        TAnalyseDescenteHautVersBas(a).DateDebutMouvement = Now
                        TAnalyseDescenteHautVersBas(a).EtatMouvement = ETATS_MOUVEMENTS.E_MOUVEMENT_EN_COURS
                    End If
        
                Case ETATS_MOUVEMENTS.E_MOUVEMENT_EN_COURS
                    '--- mouvement en cours ---
                    If .NiveauActuel = NIVEAUX_PONTS.N_BAS Then
                        TAnalyseDescenteHautVersBas(a).DateFinMouvement = Now
                        TAnalyseDescenteHautVersBas(a).EtatMouvement = ETATS_MOUVEMENTS.E_FIN_DU_MOUVEMENT
                    End If
                
                Case ETATS_MOUVEMENTS.E_FIN_DU_MOUVEMENT
                    '--- fin du mouvement (calcul du temps du mouvement) ---
                    With TAnalyseDescenteHautVersBas(a)
                        If .DateDebutMouvement <> Empty And .DateFinMouvement <> Empty Then
                            TempsMouvementSecondes = DateDiff("s", .DateDebutMouvement, .DateFinMouvement)
                            If TempsMouvementSecondes > 0 Then
                                TEtatsPonts(a).TTempsMouvements.TempsDescenteHautVersBas = TempsMouvementSecondes
                                .DateDebutMouvement = Empty
                                .DateFinMouvement = Empty
                            End If
                        End If
                    End With
                
                Case Else
            
            End Select
            
            '********************************************************************************************
            '*                DESCENTE DU NIVEAU INTERMEDIAIRE VERS LE NIVEAU BAS
            '********************************************************************************************
            
            '--- initialisation de la s�quence ---
            If .NiveauActuel = NIVEAUX_PONTS.N_INTERMEDIAIRE And .NiveauDestination = NIVEAUX_PONTS.N_INTERMEDIAIRE Then
                TAnalyseDescenteIntermediaireVersBas(a).EtatMouvement = ETATS_MOUVEMENTS.E_PAS_DE_MOUVEMENT
            End If
            If .UnDefautAuMoinsSignale = True Then
                TAnalyseDescenteIntermediaireVersBas(a).DateDebutMouvement = Empty
                TAnalyseDescenteIntermediaireVersBas(a).DateFinMouvement = Empty
            End If
            
            '--- analyse des mouvements ---
            Select Case TAnalyseDescenteIntermediaireVersBas(a).EtatMouvement
            
                Case ETATS_MOUVEMENTS.E_PAS_DE_MOUVEMENT
                    '--- pas de mouvement ---
                    If .NiveauActuel = NIVEAUX_PONTS.N_INTERMEDIAIRE And .NiveauDestination = NIVEAUX_PONTS.N_BAS Then
                        TAnalyseDescenteIntermediaireVersBas(a).DateDebutMouvement = Now
                        TAnalyseDescenteIntermediaireVersBas(a).EtatMouvement = ETATS_MOUVEMENTS.E_MOUVEMENT_EN_COURS
                    End If
        
                Case ETATS_MOUVEMENTS.E_MOUVEMENT_EN_COURS
                    '--- mouvement en cours ---
                    If .NiveauActuel = NIVEAUX_PONTS.N_BAS Then
                        TAnalyseDescenteIntermediaireVersBas(a).DateFinMouvement = Now
                        TAnalyseDescenteIntermediaireVersBas(a).EtatMouvement = ETATS_MOUVEMENTS.E_FIN_DU_MOUVEMENT
                    End If
                
                Case ETATS_MOUVEMENTS.E_FIN_DU_MOUVEMENT
                    '--- fin du mouvement (calcul du temps du mouvement) ---
                    With TAnalyseDescenteIntermediaireVersBas(a)
                        If .DateDebutMouvement <> Empty And .DateFinMouvement <> Empty Then
                            TempsMouvementSecondes = DateDiff("s", .DateDebutMouvement, .DateFinMouvement)
                            If TempsMouvementSecondes > 0 Then
                                TEtatsPonts(a).TTempsMouvements.TempsDescenteIntermediaireVersBas = TempsMouvementSecondes
                                .DateDebutMouvement = Empty
                                .DateFinMouvement = Empty
                            End If
                        End If
                    End With
                
                Case Else
            
            End Select
            
            '********************************************************************************************
            '*                             MONTEE DU NIVEAU BAS VERS INTERMEDIAIRE
            '********************************************************************************************
            
            '--- initialisation de la s�quence ---
            If .NiveauActuel = NIVEAUX_PONTS.N_BAS And .NiveauDestination = NIVEAUX_PONTS.N_BAS Then
                TAnalyseMonteeBasVersIntermediaire(a).EtatMouvement = ETATS_MOUVEMENTS.E_PAS_DE_MOUVEMENT
            End If
            If .UnDefautAuMoinsSignale = True Then
                TAnalyseMonteeBasVersIntermediaire(a).DateDebutMouvement = Empty
                TAnalyseMonteeBasVersIntermediaire(a).DateFinMouvement = Empty
            End If
            
            '--- analyse des mouvements ---
            Select Case TAnalyseMonteeBasVersIntermediaire(a).EtatMouvement
            
                Case ETATS_MOUVEMENTS.E_PAS_DE_MOUVEMENT
                    '--- pas de mouvement ---
                    If .NiveauActuel = NIVEAUX_PONTS.N_BAS And .NiveauDestination = NIVEAUX_PONTS.N_INTERMEDIAIRE Then
                        TAnalyseMonteeBasVersIntermediaire(a).DateDebutMouvement = Now
                        TAnalyseMonteeBasVersIntermediaire(a).EtatMouvement = ETATS_MOUVEMENTS.E_MOUVEMENT_EN_COURS
                    End If
        
                Case ETATS_MOUVEMENTS.E_MOUVEMENT_EN_COURS
                    '--- mouvement en cours ---
                    If .NiveauActuel = NIVEAUX_PONTS.N_INTERMEDIAIRE Then
                        TAnalyseMonteeBasVersIntermediaire(a).DateFinMouvement = Now
                        TAnalyseMonteeBasVersIntermediaire(a).EtatMouvement = ETATS_MOUVEMENTS.E_FIN_DU_MOUVEMENT
                    End If
                
                Case ETATS_MOUVEMENTS.E_FIN_DU_MOUVEMENT
                    '--- fin du mouvement (calcul du temps du mouvement) ---
                    With TAnalyseMonteeBasVersIntermediaire(a)
                        If .DateDebutMouvement <> Empty And .DateFinMouvement <> Empty Then
                            TempsMouvementSecondes = DateDiff("s", .DateDebutMouvement, .DateFinMouvement)
                            If TempsMouvementSecondes > 0 Then
                                TEtatsPonts(a).TTempsMouvements.TempsMonteeBasVersIntermediaire = TempsMouvementSecondes
                                .DateDebutMouvement = Empty
                                .DateFinMouvement = Empty
                            End If
                        End If
                    End With
                
                Case Else
            
            End Select
            
            '********************************************************************************************
            '*                                  MONTEE DU NIVEAU BAS VERS HAUT
            '********************************************************************************************
            
            '--- initialisation de la s�quence ---
            If .NiveauActuel = NIVEAUX_PONTS.N_BAS And .NiveauDestination = NIVEAUX_PONTS.N_BAS Then
                TAnalyseMonteeBasVersHaut(a).EtatMouvement = ETATS_MOUVEMENTS.E_PAS_DE_MOUVEMENT
            End If
            If .UnDefautAuMoinsSignale = True Then
                TAnalyseMonteeBasVersHaut(a).DateDebutMouvement = Empty
                TAnalyseMonteeBasVersHaut(a).DateFinMouvement = Empty
            End If
            
            '--- analyse des mouvements ---
            Select Case TAnalyseMonteeBasVersHaut(a).EtatMouvement
            
                Case ETATS_MOUVEMENTS.E_PAS_DE_MOUVEMENT
                    '--- pas de mouvement ---
                    If .NiveauActuel = NIVEAUX_PONTS.N_BAS And .NiveauDestination = NIVEAUX_PONTS.N_HAUT Then
                        TAnalyseMonteeBasVersHaut(a).DateDebutMouvement = Now
                        TAnalyseMonteeBasVersHaut(a).EtatMouvement = ETATS_MOUVEMENTS.E_MOUVEMENT_EN_COURS
                    End If
        
                Case ETATS_MOUVEMENTS.E_MOUVEMENT_EN_COURS
                    '--- mouvement en cours ---
                    If .NiveauActuel = NIVEAUX_PONTS.N_HAUT Then
                        TAnalyseMonteeBasVersHaut(a).DateFinMouvement = Now
                        TAnalyseMonteeBasVersHaut(a).EtatMouvement = ETATS_MOUVEMENTS.E_FIN_DU_MOUVEMENT
                    End If
                
                Case ETATS_MOUVEMENTS.E_FIN_DU_MOUVEMENT
                    '--- fin du mouvement (calcul du temps du mouvement) ---
                    With TAnalyseMonteeBasVersHaut(a)
                        If .DateDebutMouvement <> Empty And .DateFinMouvement <> Empty Then
                            TempsMouvementSecondes = DateDiff("s", .DateDebutMouvement, .DateFinMouvement)
                            If TempsMouvementSecondes > 0 Then
                                TEtatsPonts(a).TTempsMouvements.TempsMonteeBasVersHaut = TempsMouvementSecondes
                                .DateDebutMouvement = Empty
                                .DateFinMouvement = Empty
                            End If
                        End If
                    End With
                
                Case Else
            
            End Select
    
            '********************************************************************************************
            '*                       TRANSLATION D'UN POSTE DE DEPART VERS ARRIVEE
            '********************************************************************************************
    
            '--- initialisation de la s�quence ---
            If .UnDefautAuMoinsSignale = True Then
                TAnalyseTranslation(a).EtatMouvement = ETATS_MOUVEMENTS.E_PAS_DE_MOUVEMENT
                TAnalyseTranslation(a).DateDebutMouvement = Empty
                TAnalyseTranslation(a).DateFinMouvement = Empty
                TPostesDepart(a) = 0
                TPostesArrivee(a) = 0
            End If
            
            '--- analyse des mouvements ---
            Select Case TAnalyseTranslation(a).EtatMouvement
            
                Case ETATS_MOUVEMENTS.E_PAS_DE_MOUVEMENT
                    '--- pas de mouvement ---
                    If .PosteActuel <> .PosteDestination Then
                        TAnalyseTranslation(a).DateDebutMouvement = Now
                        TPostesDepart(a) = .PosteActuel
                        TPostesArrivee(a) = .PosteDestination
                        TAnalyseTranslation(a).EtatMouvement = ETATS_MOUVEMENTS.E_MOUVEMENT_EN_COURS
                    End If
        
                Case ETATS_MOUVEMENTS.E_MOUVEMENT_EN_COURS
                    '--- mouvement en cours ---
                    If .PosteActuel = .PosteDestination Then
                        TAnalyseTranslation(a).DateFinMouvement = Now
                        TAnalyseTranslation(a).EtatMouvement = ETATS_MOUVEMENTS.E_FIN_DU_MOUVEMENT
                    End If
                
                Case ETATS_MOUVEMENTS.E_FIN_DU_MOUVEMENT
                    '--- fin du mouvement (calcul du temps du mouvement) ---
                    With TAnalyseTranslation(a)
                        If .DateDebutMouvement <> Empty And .DateFinMouvement <> Empty And _
                           TPostesDepart(a) <> 0 And TPostesArrivee(a) <> 0 Then
                            
                            TempsMouvementSecondes = DateDiff("s", .DateDebutMouvement, .DateFinMouvement)
                            If TempsMouvementSecondes > 0 Then
                                
                                '--- calculer le temps du mouvements ---
                                TEtatsPonts(a).TTempsMouvements.TTempsTranslation(TPostesDepart(a), TPostesArrivee(a)) = TempsMouvementSecondes
                                .DateDebutMouvement = Empty
                                .DateFinMouvement = Empty
                                
                                '--- relancer l'analyse ---
                                TAnalyseTranslation(a).EtatMouvement = ETATS_MOUVEMENTS.E_PAS_DE_MOUVEMENT
                            
                            End If
                        
                        Else
                            
                            '--- relancer l'analyse ---
                            TAnalyseTranslation(a).EtatMouvement = ETATS_MOUVEMENTS.E_PAS_DE_MOUVEMENT
                        
                        End If
                    End With
                
                Case Else
            
            End Select
    
        End With
    
    Next a
    
End Sub

'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' R�le      : Analyse des charges en lignes pour les PONTS
' Entr�es :
' Retours :
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Sub AnalyseChargesEnLignePonts()
    
    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next

    '--- d�claration ---
    Dim a As Integer
    Static TCopieEtatsPonts(PONTS.P_1 To PONTS.P_2) As EtatsPonts
    
    '--- analyse en fonction du PC ---
    'If TypePC <> TYPES_PC. Then Exit Sub

    '*************************************************************************************************
    '                                              ANALYSE POUR LES PONTS
    '*************************************************************************************************
    For a = PONTS.P_1 To PONTS.P_2
    
        With TEtatsPonts(a)
            
            '*************************************************************************************************
            '                                            PRISE D'UNE CHARGE PAR LE PONT
            '*************************************************************************************************
            If .NumCharge >= CHARGES.C_NUM_MINI And _
               .NumCharge <= CHARGES.C_NUM_MAXI And _
               TCopieEtatsPonts(a).NumCharge = 0 Then
    
            End If
    
            '*************************************************************************************************
            '                                                 CHARGE DEJA SUR LE PONT
            '*************************************************************************************************
            If .NumCharge >= CHARGES.C_NUM_MINI And .NumCharge <= CHARGES.C_NUM_MAXI Then
                If .NumCharge = TCopieEtatsPonts(a).NumCharge Then
    
                    With TEtatsCharges(.NumCharge)
                    
                        '--- remplissage de la fiche pour les alarmes de la ligne ---
                        .AlarmesLigne = AjoutNumDefautsSansDoublons(.AlarmesLigne, AlarmesLigneEnCours)
                        
                        '--- remplissage de la fiche de production pour l'�gouttage ---
                        If .NbrPostesTraites > 0 Then
                            If TEtatsPonts(a).NiveauActuel = NIVEAUX_PONTS.N_HAUT And _
                               TEtatsPonts(a).TEntreesAPI.M_MoteurTourneLevPont = False And _
                               TEtatsPonts(a).TEntreesAPI.M_MoteurTourneTrlPont = False And _
                               .TDetailsFichesProduction(.NbrPostesTraites).NumPoste = TEtatsPonts(a).PosteActuel Then
                                    If .TDetailsFichesProduction(.NbrPostesTraites).DateDebutEgouttage = Empty Then
                                        .TDetailsFichesProduction(.NbrPostesTraites).DateDebutEgouttage = Now  'd�but de l'�gouttage
                                    Else
                                        .TDetailsFichesProduction(.NbrPostesTraites).DateFinEgouttage = Now       'fin de l'�gouttage
                                    End If
                            End If
                        End If
    
                    End With
    
                End If
            End If
    
            '*************************************************************************************************
            '                                       DEPOSE D'UNE CHARGE PAR LE PONT
            '*************************************************************************************************
            If .NumCharge = 0 And _
               TCopieEtatsPonts(a).NumCharge >= CHARGES.C_NUM_MINI And _
               TCopieEtatsPonts(a).NumCharge <= CHARGES.C_NUM_MAXI Then
    
                    With TEtatsCharges(TCopieEtatsPonts(a).NumCharge)
                        
            
                    End With
            
            End If
            
            '--- affectation ---
            TCopieEtatsPonts(a).NumCharge = .NumCharge
    
        End With
    
    Next a

End Sub

'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' R�le      : Analyse la fin de cycle de l'�tuve
' Entr�es :
' Retours :
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Sub AnalyseFinDeCycleEtuve()
    
    '--- aiguillage en cas d'erreurs ---
'    On Error Resume Next
'
'    '--- d�claration ---
'    Dim a As Integer, _
'           NumCharge As Integer
'
'    '--- analyse en fonction du PC ---
'    If TypePC <> TYPES_PC. Then Exit Sub
'
'    '--- analyse de la charge au poste ---
'    With TEtatsPostes(POSTES.P_A18)
'
'        If .NumCharge >= CHARGES.C_NUM_MINI And _
'           .NumCharge <= CHARGES.C_NUM_MAXI Then
'
'            '--- affectation du num�ro de charge ---
'            NumCharge = .NumCharge
'
'            If TEtatsEtuveA18.CycleTermine = True Then
'
'                With TEtatsCharges(NumCharge)
'
'                    If .PtrZoneGammeAnodisation > 0 Then
'
'                        '--- fin du temps du poste r�el de la gamme d'anodisation ---
'                        .TGammesAnodisation.TDetailsGammesAnodisation(.PtrZoneGammeAnodisation).FinDuTempsPosteReel = True
'
'                        '--- anti-rebond ---
'                        TEtatsEtuveA18.CycleTermine = False
'
'                    End If
'
'                End With
'
'            End If
'
'        End If
'
'    End With
    
End Sub


Private Function MakeTrue( _
                 ByRef bValue As Boolean) As Boolean
    MakeTrue = True
    bValue = True
End Function


'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' R�le      : Analyse des charges en lignes pour les POSTES
' Entr�es :
' Retours :
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Sub AnalyseChargesEnLignePostes()
    
    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next

    '--- d�claration ---
    Dim a As Integer, _
           NumCuve As Integer
    Static TCopieEtatsPostes(POSTES.P_CHGT_1 To DERNIER_POSTE) As EtatsPostes
    
    '--- analyse en fonction du PC ---
    'If TypePC <> TYPES_PC. Then Exit Sub
    
    '*************************************************************************************************
    '                                               ANALYSE POUR LES POSTES
    '*************************************************************************************************
    For a = POSTES.P_CHGT_1 To DERNIER_POSTE
                    
        '--- recherche du n� de cuve correspondant au poste ---
        NumCuve = CorrespondancePostesCuvesAPI(a)
    
        With TEtatsPostes(a)
        
            '*************************************************************************************************
            '                                        ENTREE D'UNE CHARGE DANS LE POSTE
            '*************************************************************************************************
            If .NumCharge >= CHARGES.C_NUM_MINI And _
               .NumCharge <= CHARGES.C_NUM_MAXI And _
               TCopieEtatsPostes(a).NumCharge = 0 Then
    
                With TEtatsCharges(.NumCharge)
                
                    '--- incr�mentation du nombre de postes trait�s ---
                    Inc .NbrPostesTraites
                    If .NbrPostesTraites > NBR_LIGNES_DETAILS_FICHES_PRODUCTION Then
                        .NbrPostesTraites = NBR_LIGNES_DETAILS_FICHES_PRODUCTION
                    End If
                    
                    '--- incr�mentation du pointeur de la zone d'anodisation ---
                    IncrementationPtrZoneGammeAnodisation a
                    
                    '--- enregistrer le poste r�el dans la gamme d'anodisation de la charge ---
                    EnregistreNumPosteReelGamme a
                    
                    '--- remplissage de la fiche de production pour le n� de poste et la date d'entr�e ---
                    If .NbrPostesTraites > 0 Then
                        .TDetailsFichesProduction(.NbrPostesTraites).NumPoste = a
                        .TDetailsFichesProduction(.NbrPostesTraites).DateEntreePoste = Now
                    End If
                    
                    '--- remplissage de la fiche de production pour la temp�rature en entr�e ---
                    If .NbrPostesTraites > 0 And NumCuve > 0 Then
                        .TDetailsFichesProduction(.NbrPostesTraites).TemperatureEnEntree = TEtatsCuves(NumCuve).Temperatures.TempActuelle
                    End If
                    
                    '--- remplissage de la fiche de production pour l'analyseur d'anodisation en entr�e ---
                    If .NbrPostesTraites > 0 And NumCuve > 0 Then
                        If TEtatsCuves(NumCuve).DefinitionCuve.PresenceAnalyseurAnodisation = True Then
                            .TDetailsFichesProduction(.NbrPostesTraites).AnalyseurEnEntree = TEtatsCuves(NumCuve).TEntreesAPI.E_Analogique_Analyseur
                        End If
                    End If
                    If a >= POSTES.P_D1 And a <= POSTES.P_D2 Then
                        .DateArriveeAuDechargement = Now
                        
                    End If
                   
                   

                End With
            
            End If
    
            '*************************************************************************************************
            '                                                CHARGE DEJA DANS LE POSTE
            '*************************************************************************************************
            If .NumCharge >= CHARGES.C_NUM_MINI And .NumCharge <= CHARGES.C_NUM_MAXI Then
                If .NumCharge = TCopieEtatsPostes(a).NumCharge Then
    
                    With TEtatsCharges(.NumCharge)
                    
                        '--- remplissage de la fiche pour les alarmes de la ligne ---
                        .AlarmesLigne = AjoutNumDefautsSansDoublons(.AlarmesLigne, AlarmesLigneEnCours)
                        
                        '--- remplissage de la fiche de production pour la date de sortie (comptage en temps r�el) ---
                        If .NbrPostesTraites > 0 Then
                            .TDetailsFichesProduction(.NbrPostesTraites).DateSortiePoste = Now
                        End If
    
                        '--- remplissage de la fiche de production pour les alarmes dans le poste ---
                        If .NbrPostesTraites > 0 And NumCuve > 0 Then
                            .TDetailsFichesProduction(.NbrPostesTraites).AlarmesPoste = TEtatsCuves(NumCuve).ListeNumDefautsSiCharge
                        End If
                        '--- remplissage de la fiche de production pour les valeurs du redresseur ---

                        If .NbrPostesTraites > 0 Then

                            Select Case a

                                Case POSTES.P_C13

                                    '--- premier poste d'anodisation ---
                                    If .TDetailsPhasesProduction(PHASES_GAMMES_REDRESSEURS.PH_T4).TempsPhase > 0 Then

                                        '--- prendre la mesure uniquement sur la phase 4 ---

                                        With .TDetailsFichesProduction(.NbrPostesTraites)

                                            If TEtatsRedresseurs(REDRESSEURS.R_C13).NumPhaseEnCours = PHASES_GAMMES_REDRESSEURS.PH_T4 Then

                                                .IRedresseur = TEtatsRedresseurs(REDRESSEURS.R_C13).I

                                                .URedresseur = TEtatsRedresseurs(REDRESSEURS.R_C13).U

                                            End If

                                        End With

                                    End If


                                Case POSTES.P_C14
                                    '--- second poste d'anodisation ---
                                    If .TDetailsPhasesProduction(PHASES_GAMMES_REDRESSEURS.PH_T4).TempsPhase > 0 Then

                                        '--- prendre la mesure uniquement sur la phase 4 ---
                                        With .TDetailsFichesProduction(.NbrPostesTraites)

                                            If TEtatsRedresseurs(REDRESSEURS.R_C14).NumPhaseEnCours = PHASES_GAMMES_REDRESSEURS.PH_T4 Then
                                                .IRedresseur = TEtatsRedresseurs(REDRESSEURS.R_C14).I
                                                .URedresseur = TEtatsRedresseurs(REDRESSEURS.R_C14).U
                                            End If
                                        End With
                                    End If



                                Case POSTES.P_C15
                                    '--- troisi?me poste d'anodisation ---
                                    If .TDetailsPhasesProduction(PHASES_GAMMES_REDRESSEURS.PH_T4).TempsPhase > 0 Then
                                        '--- prendre la mesure uniquement sur la phase 4 ---

                                        With .TDetailsFichesProduction(.NbrPostesTraites)

                                            If TEtatsRedresseurs(REDRESSEURS.R_C15).NumPhaseEnCours = PHASES_GAMMES_REDRESSEURS.PH_T4 Then
                                                .IRedresseur = TEtatsRedresseurs(REDRESSEURS.R_C15).I
                                                .URedresseur = TEtatsRedresseurs(REDRESSEURS.R_C15).U
                                            End If

                                        End With



                                    End If



                                Case POSTES.P_C16

                                    '--- quatri?me poste d'anodisation ---
                                    If .TDetailsPhasesProduction(PHASES_GAMMES_REDRESSEURS.PH_T4).TempsPhase > 0 Then

                                        '--- prendre la mesure uniquement sur la phase 4 ---

                                        With .TDetailsFichesProduction(.NbrPostesTraites)
                                            If TEtatsRedresseurs(REDRESSEURS.R_C16).NumPhaseEnCours = PHASES_GAMMES_REDRESSEURS.PH_T4 Then
                                                .IRedresseur = TEtatsRedresseurs(REDRESSEURS.R_C16).I
                                                .URedresseur = TEtatsRedresseurs(REDRESSEURS.R_C16).U
                                            End If
                                        End With

                                    End If
                                Case Else
                            End Select
                        '202501
                        'enregistreRedresseursAno .NumCharge, a
                        End If
                    
                    End With
    
                    '--- d�compte dans la gamme du temps au poste (en secondes) ---
                    DecompteDuTempsAuPosteSecondes a
                    
                    '--- d�compte dans la gamme du temps d'alerte au poste (en secondes) ---
                    DecompteDuTempsAlerteAuPosteSecondes a
                
                End If
            End If
          
            '*************************************************************************************************
            '                                               SORTIE D'UNE CHARGE DU POSTE
            '*************************************************************************************************
            If .NumCharge = 0 And _
               TCopieEtatsPostes(a).NumCharge >= CHARGES.C_NUM_MINI And _
               TCopieEtatsPostes(a).NumCharge <= CHARGES.C_NUM_MAXI Then
    
                    With TEtatsCharges(TCopieEtatsPostes(a).NumCharge)
                        
                        '--- remplissage de la fiche de production pour la date de sortie ---
                        If .NbrPostesTraites > 0 Then
                            .TDetailsFichesProduction(.NbrPostesTraites).DateSortiePoste = Now    'date de sortie
                        End If
                        
                        '--- remplissage de la fiche de production pour la temp�rature en sortie ---
                        If .NbrPostesTraites > 0 And NumCuve > 0 Then
                            .TDetailsFichesProduction(.NbrPostesTraites).TemperatureEnSortie = TEtatsCuves(NumCuve).Temperatures.TempActuelle
                        End If
                    
                        '--- remplissage de la fiche de production pour l'analyseur d'anodisation en sortie ---
                        If .NbrPostesTraites > 0 And NumCuve > 0 Then
                            If TEtatsCuves(NumCuve).DefinitionCuve.PresenceAnalyseurAnodisation = True Then
                                .TDetailsFichesProduction(.NbrPostesTraites).AnalyseurEnSortie = TEtatsCuves(NumCuve).TEntreesAPI.E_Analogique_Analyseur
                            End If
                        End If
 
 
                        
                         
    
                    
                    End With
                    If a >= POSTES.P_D1 And a <= POSTES.P_D2 Then
                       'SZP 02/2025
                            
                        If MODE_DECONNECTE = False Then
                            insertionClipperPointage (TCopieEtatsPostes(a).NumCharge)
                            EnregistrementProductionLocal (TCopieEtatsPostes(a).NumCharge)
                        End If
                
                    End If
            
            End If
        
            '--- affectation ---
            TCopieEtatsPostes(a).NumCharge = .NumCharge
    
        End With

    Next a

End Sub

Public Sub enregistreRedresseursAno(NumCharge As Integer, NumPoste As Integer)

    With TEtatsCharges(NumCharge)
        .FinPhase4 = True
    End With

    

End Sub

'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' R�le      : Gestion des commandes de l'op�rateur
' Entr�es :
' Retours :
' D�tails  : Cette fonction ne sert que lorsque l'op�rateur a pris le contr�le des 2 ponts. Elle permet de faire des
'                 d�placements et transferts complexes en g�rant automatiquement l'anti-collision
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Sub GestionCommandesOperateur()
    
    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next
    
    '--- constantes priv�es ---
    
    '--- d�claration ---
    Dim a As Integer                                                                                    'r�serv� pour les boucles FOR ... NEXT
    Dim NumPont As Integer                                                                       'num�ro d'un pont
    Dim NumPontReelDuTransfert  As Integer                                            'num�ro du pont r�el qui va effectuer le transfert
    
    Dim NumPosteDepart As Integer                                                           'num�ro d'un POSTE de DEPART dans le cas d'une PREMISSE (transfert d'une charge)
    Dim NumPosteArrivee As Integer                                                          'num�ro d'un POSTE d'ARRIVEE dans le cas d'une PREMISSE (transfert d'une charge)
    
    Dim TypeCollision As Integer                                                                'repr�sente le n� du type de collision  utilis� pour le contr�le anti-collision
    Dim NumPontOppose As Integer                                                           'num�ro du pont oppos� utilis� pour le contr�le anti-collision
    Dim NumPosteAssurantSecurite As Integer                                          'num�ro du poste assurant la s�curit� pour le contr�le anti-collision
    
    Dim TempsEgouttageSecondes As Integer                                           'temps d'�gouttage en secondes � implanter dans une pr�misse avant
                                                                                                                   'le transfert vers l'API (ce temps est celui fourni dans la gamme)
    Dim DelaiSupStabilisationChargeSecondes As Integer                        'd�lai suppl�mentaire de stabilisation de la charge
    
    Dim CouleurReponseAntiCollision As Long                                           'couleur d'une r�ponse � une gestion d'anti-collision
    Dim CouleurReponseDeplacementOuTransfert As Long                       'couleur d'une r�ponse � un d�placement de pont ou un transfert
    
    Dim ReponseAntiCollision As String                                                      'r�ponse � une gestion d'anti-collision
    Dim ReponseDeplacementOuTransfert As String                                  'r�ponse � un d�placement de pont ou un transfert
    
    Dim TypeCycle As TYPES_CYCLES                                                      'type de cycle fonction de l'�num�ration TYPES_CYCLES
    
    Dim FicheVideCommandesOperateur As VarCommandesOperateur    'fiche vide des commandes op�rateur
    
    Static MemReponseAntiCollision As String                                           'm�moire de la r�ponse � une gestion d'anti-collision
    Static MemReponseDeplacementOuTransfert As String                       'm�moire de la r�ponse � un d�placement de pont ou un transfert
    
    '--- analyse en fonction du PC ---
    'If TypePC <> TYPES_PC. Then Exit Sub

    '--- v�rification du contr�le des 2 ponts par l'op�rateur ---
    If TEtatsPonts(PONTS.P_1).ControleParOperateur = True And _
       TEtatsPonts(PONTS.P_2).ControleParOperateur = True Then

        '--- transfert de la commande dans les variables locales ---
        With TCommandesOperateur(1)
            TypeCycle = .TypeCycle
            NumPont = .NumPont
            NumPosteDepart = .NumPosteDepart
            NumPosteArrivee = .NumPosteArrivee
            TempsEgouttageSecondes = .TempsEgouttageSecondes
        End With
        DelaiSupStabilisationChargeSecondes = 0
                
        If TypeCycle <> TYPES_CYCLES.TC_INCONNU Then
                
            If NumPont >= PONTS.P_1 And NumPont <= PONTS.P_2 And _
               NumPosteDepart >= POSTES.P_CHGT_1 And NumPosteDepart <= DERNIER_POSTE And _
               NumPosteArrivee >= POSTES.P_CHGT_1 And NumPosteArrivee <= DERNIER_POSTE Then
             
                '****************************************************************************************************
                '                                           analyse de la commande DEPLACEMENT
                '****************************************************************************************************
                If TypeCycle = TYPES_CYCLES.TC_DEPLACEMENT_PONT Then
                    
                    '--- gestion de l'anti-collision ---
                    ReponseAntiCollision = ControleAntiCollision(NumPont, _
                                                                                             NumPosteDepart, _
                                                                                             NumPosteArrivee, _
                                                                                             TypeCollision, _
                                                                                             NumPontOppose, _
                                                                                             NumPosteAssurantSecurite, _
                                                                                             CouleurReponseAntiCollision)
                    
                    If TypeCollision = TYPES_COLLISION.AUCUN_RISQUE Then
                        
                        '************************************************************************
                        '        pas de risque de collision / lancement du d�placement
                        '************************************************************************
                        ReponseDeplacementOuTransfert = AutomatiqueDeplacementPont(NumPont, _
                                                                                                                                   NumPosteArrivee, _
                                                                                                                                   CouleurReponseDeplacementOuTransfert)
                    
                        '************************************************************************
                        '            contr�le du d�placement et d�calage des commandes
                        '************************************************************************
                        If ReponseDeplacementOuTransfert = OK Then
                            For a = LBound(TCommandesOperateur()) To (UBound(TCommandesOperateur()) - 1)
                                TCommandesOperateur(a) = TCommandesOperateur(a + 1)
                            Next a
                            TCommandesOperateur(UBound(TCommandesOperateur())) = FicheVideCommandesOperateur
                        End If
                    
                    Else
                        
                        '************************************************************************
                        '            risque de collision / d�placement du pont oppos�
                        '************************************************************************
                        If NumPontOppose > 0 And NumPosteAssurantSecurite > 0 Then
                            ReponseDeplacementOuTransfert = AutomatiqueDeplacementPont(NumPontOppose, _
                                                                                                                                        NumPosteAssurantSecurite, _
                                                                                                                                        CouleurReponseDeplacementOuTransfert)
                        End If
                        
                    End If
                
                End If
                
                '****************************************************************************************************
                '                                              analyse de la commande TRANSFERT
                '****************************************************************************************************
                If TypeCycle = TYPES_CYCLES.TC_TRANSFERT_CHARGE Then
                        
                    '--- gestion de l'anti-collision ---
                    ReponseAntiCollision = ControleAntiCollision(NumPont, _
                                                                                              NumPosteDepart, _
                                                                                              NumPosteArrivee, _
                                                                                              TypeCollision, _
                                                                                              NumPontOppose, _
                                                                                              NumPosteAssurantSecurite, _
                                                                                              CouleurReponseAntiCollision)
                    
                    If TypeCollision = TYPES_COLLISION.AUCUN_RISQUE Then
                        
                        '************************************************************************
                        '            pas de risque de collision / lancement de transfert
                        '************************************************************************
                        ReponseDeplacementOuTransfert = AutomatiqueTransfertCharge(NumPont, _
                                                                                                                                  NumPosteDepart, _
                                                                                                                                  NumPosteArrivee, _
                                                                                                                                  TempsEgouttageSecondes, _
                                                                                                                                  DelaiSupStabilisationChargeSecondes, _
                                                                                                                                  NumPontReelDuTransfert:=NumPontReelDuTransfert, _
                                                                                                                                  CouleurReponse:=CouleurReponseDeplacementOuTransfert)
                    
                        '************************************************************************
                        '              contr�le du transfert et d�calage des commandes
                        '************************************************************************
                        If ReponseDeplacementOuTransfert = OK Then
                            For a = LBound(TCommandesOperateur()) To (UBound(TCommandesOperateur()) - 1)
                                TCommandesOperateur(a) = TCommandesOperateur(a + 1)
                            Next a
                            TCommandesOperateur(UBound(TCommandesOperateur())) = FicheVideCommandesOperateur
                        End If
                    
                    Else
                        
                        '************************************************************************
                        '            risque de collision / d�placement du pont oppos�
                        '************************************************************************
                        If NumPontOppose > 0 And NumPosteAssurantSecurite > 0 Then
                            ReponseDeplacementOuTransfert = AutomatiqueDeplacementPont(NumPontOppose, _
                                                                                                                                        NumPosteAssurantSecurite, _
                                                                                                                                        CouleurReponseDeplacementOuTransfert)
                        End If
                        
                    End If
                    
                End If
        
            End If
                
        End If
    
    End If
                
End Sub

'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' R�le      : Constitue le moteur inf�rence
' Entr�es :
' Retours : Les divers valeurs de retours sont m�moris�es dans la variable publique du moteur d'inf�rence
'                 appel�e TMoteurInference
'                 Cette variable permet l'affichage des divers valeurs dans la fenetre FMoteurInference
' D�tails  : Le moteur d'inf�rence g�re la totalit� de la ligne (gammes, charges, ordonnancement, etc...)
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Sub MoteurInference()
    
    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next
    
    '--- constantes priv�es ---
    
    '--- d�claration ---
    Dim TravailAvecMI(PONTS.P_1 To PONTS.P_2) As Boolean            'indique que le syst�me vient de travailler
                                                                                                               'avec le moteur d'inf�rence pour un pont
    Dim ChargePrioritaireSiAnodisationC13Impose As Boolean            'indique qu'une charge avec Anodisation C13 impos�
                                                                                                               'a �t� valid� comme prioritaire
    Dim ChargePrioritaireSiAnodisationC14Impose As Boolean            'indique qu'une charge avec Anodisation C14 impos�
                                                                                                               'a �t� valid� comme prioritaire
    Dim ChargePrioritaireSiAnodisationC15Impose As Boolean            'indique qu'une charge avec Anodisation C15 impos�
                                                                                                               'a �t� valid� comme prioritaire
    Dim ChargePrioritaireSiAnodisationC16Impose As Boolean            'indique qu'une charge avec Anodisation C16 impos�
                                                                                                               'a �t� valid� comme prioritaire
    
    Dim ChargePrioritaireSiAnodisationAutomatique As Boolean          'indique qu'une charge avec Anodisation sur AUTOMATIQUE
                                                                                                               'a �t� valid� comme prioritaire
    
    Dim a As Integer                                                                                'r�serv� pour les boucles FOR ... NEXT
    
    Dim NumPontPrioritaire As Integer                                                   'n� du pont prioritaire lorsque un des ponts est condammn�
                                                                                                               'ex : si le PONT 1 est condammn�, NumPontPrioritaire = PONT 2
                                                                                                               '       si le PONT 2 est condammn�, NumPontPrioritaire = PONT 1
                                                                                                               '       si le PONT 1 et le PONT 2 fonctionne, NumPontPrioritaire = 0
                                                                                                               '       si le PONT 1 et le PONT 2 sont condammn�s, NumPontPrioritaire = 0
    Dim NumPontReelDuTransfert  As Integer                                        'num�ro du pont r�el qui va effectuer le transfert
    
    Dim NumCharge As Integer                                                               'indique un num�ro de charge
    Dim NumZoneDepart As Integer                                                        'num�ro d'une ZONE de  DEPART d'une GAMME
    Dim NumZoneArrivee As Integer                                                       'num�ro d'une ZONE d'ARRIVEE d'une GAMME
    Dim NumPosteDepart As Integer                                                       'num�ro d'un POSTE de DEPART dans le cas d'une PREMISSE (transfert d'une charge)
    Dim NumPosteArrivee As Integer                                                      'num�ro d'un POSTE d'ARRIVEE dans le cas d'une PREMISSE (transfert d'une charge)
    
    Dim TypeCollision As Integer                                                            'repr�sente le n� du type de collision  utilis� pour le contr�le anti-collision
    Dim NumPontOppose As Integer                                                       'num�ro du pont oppos� utilis� pour le contr�le anti-collision
    Dim NumPosteAssurantSecurite As Integer                                      'num�ro du poste assurant la s�curit� pour le contr�le anti-collision
    
    Dim TempsEgouttageSecondes As Integer                                       'temps d'�gouttage en secondes � implanter dans une pr�misse avant
                                                                                                               'le transfert vers l'API (ce temps est celui fourni dans la gamme)
    Dim DelaiSupStabilisationChargeSecondes As Integer                    'd�lai suppl�mentaire de stabilisation de la charge
    Dim CouleurReponse As Long                                                           'couleur d'une r�ponse � une demande (pour la zone des renseignements)
    
    Dim ReponseAntiCollision As String                                                 'r�ponse � une gestion de l'anti-collision
    Dim ReponseDeplacementPont As String                                         'r�ponse donn�e lors de l'envoi de la commande deplacement d'un pont
    Dim ReponseControleBainsPrioritaires As String                            'r�ponse donn�e lors du contr�le des bains prioritaires
    Dim ReponseTransfertCharge As String                                           'r�ponse donn�e lors de l'envoi de la commande transfert
    
    Dim logMoteurInference As Boolean
    
    Static MemReponseTransfertCharge As String                                 'm�moire de la r�ponse donn�e lors de l'envoi de la commande transfert

    Static TDatesDerniersTransfertsCharges(PONTS.P_1 To PONTS.P_2) As Date   'indique la date du dernier transfert de charge de chaque pont
    Static TDatesDerniersDeplacementsAVide(PONTS.P_1 To PONTS.P_2) As Date 'indique la date du dernier d�placement � vide de chaque pont

    '--- analyse en fonction du PC ---
    'If TypePC <> TYPES_PC. Then Exit Sub
    
    logMoteurInference = False
    
    '**********************************************************************************************************
    '**********************************************************************************************************
    '*     Remplir le tableau sur l'ordre de sortie des charges dans le tableau du moteur d'inf�rence
    '**********************************************************************************************************
    '**********************************************************************************************************
    RechercheOrdreSortieCharges
  
    
    With TMoteurInference
    
        '********************************************************************************************
        '********************************************************************************************
        '*                           V�rification de la condamnation de l'un des ponts
        '********************************************************************************************
        '********************************************************************************************
        'regarder si un des ponts est condamn� auquel cas la totalit� du travail passe sur
        'l'autre pont qui devient le pont prioritaire
        
        If TEtatsPonts(PONTS.P_1).Condamnation = True And TEtatsPonts(PONTS.P_2).Condamnation = False Then
            NumPontPrioritaire = PONTS.P_2
        ElseIf TEtatsPonts(PONTS.P_1).Condamnation = False And TEtatsPonts(PONTS.P_2).Condamnation = True Then
            NumPontPrioritaire = PONTS.P_1
        Else
            NumPontPrioritaire = 0
        End If
        
        '********************************************************************************************
        '********************************************************************************************
        '         Analyse complet du chargement (prochaine charge � rentrer dans la ligne)
        '                               ATTENTION aux charges qui sont prioritaires
        '********************************************************************************************
        '********************************************************************************************
        's�lection du poste de chargement en fonction des cas
        
        '--- recherche du prochain poste de chargement si le poste d'anodisation C13 est impos� dans la gamme ---
        .ProchainNumPosteChargementSiAnodisationC13Impose = ProchainNumeroPosteChargementSiAnodisationC13Impose(ChargePrioritaireSiAnodisationC13Impose)
        
        '--- recherche du prochain poste de chargement si le poste d'anodisation C14 est impos� dans la gamme ---
        .ProchainNumPosteChargementSiAnodisationC14Impose = ProchainNumeroPosteChargementSiAnodisationC14Impose(ChargePrioritaireSiAnodisationC14Impose)
        
        '--- recherche du prochain poste de chargement si le poste d'anodisation C15 est impos� dans la gamme ---
        .ProchainNumPosteChargementSiAnodisationC15Impose = ProchainNumeroPosteChargementSiAnodisationC15Impose(ChargePrioritaireSiAnodisationC15Impose)
        
        '--- recherche du prochain poste de chargement si le poste d'anodisation C16 est impos� dans la gamme ---
        .ProchainNumPosteChargementSiAnodisationC16Impose = ProchainNumeroPosteChargementSiAnodisationC16Impose(ChargePrioritaireSiAnodisationC16Impose)
        
        '--- recherche du prochain poste de chargement si le choix du poste d'anodisation est automatique dans la gamme ---
        .ProchainNumPosteChargementSiAnodisationAutomatique = ProchainNumeroPosteChargementSiAnodisationAutomatique(ChargePrioritaireSiAnodisationAutomatique)
        
        If ExistenceChargeEnLigneHorsChargementDechargement = False Then
    
            '********************************************************************************************
            '*         PAS DE CHARGE DANS LA LIGNE (hormis chargement ou d�chargement)
            '********************************************************************************************
            If ChargePrioritaireSiAnodisationC13Impose = True Or _
               ChargePrioritaireSiAnodisationC14Impose = True Or _
               ChargePrioritaireSiAnodisationC15Impose = True Or _
               ChargePrioritaireSiAnodisationC16Impose = True Or _
               ChargePrioritaireSiAnodisationAutomatique = True Then
                
                '--- IL Y A UNE CHARGE PRIORITAIRE / affectation du prochain n� de poste de chargement ---
                If ChargePrioritaireSiAnodisationC13Impose = True Then
                    .ProchainNumPosteChargement = .ProchainNumPosteChargementSiAnodisationC13Impose
                ElseIf ChargePrioritaireSiAnodisationC14Impose = True Then
                    .ProchainNumPosteChargement = .ProchainNumPosteChargementSiAnodisationC14Impose
                ElseIf ChargePrioritaireSiAnodisationC15Impose = True Then
                    .ProchainNumPosteChargement = .ProchainNumPosteChargementSiAnodisationC15Impose
                ElseIf ChargePrioritaireSiAnodisationC16Impose = True Then
                    .ProchainNumPosteChargement = .ProchainNumPosteChargementSiAnodisationC16Impose
                ElseIf ChargePrioritaireSiAnodisationAutomatique = True Then
                    .ProchainNumPosteChargement = .ProchainNumPosteChargementSiAnodisationAutomatique
                Else
                    .ProchainNumPosteChargement = 0
                End If
                
           Else

                '--- IL N'Y A PAS DE CHARGE PRIORITAIRE / affectation du prochain n� de poste de chargement ---
                If .ProchainNumPosteChargementSiAnodisationC13Impose > 0 Then
                    .ProchainNumPosteChargement = .ProchainNumPosteChargementSiAnodisationC13Impose
                ElseIf .ProchainNumPosteChargementSiAnodisationC14Impose > 0 Then
                    .ProchainNumPosteChargement = .ProchainNumPosteChargementSiAnodisationC14Impose
                ElseIf .ProchainNumPosteChargementSiAnodisationC15Impose > 0 Then
                    .ProchainNumPosteChargement = .ProchainNumPosteChargementSiAnodisationC15Impose
                ElseIf .ProchainNumPosteChargementSiAnodisationC16Impose > 0 Then
                    .ProchainNumPosteChargement = .ProchainNumPosteChargementSiAnodisationC16Impose
                ElseIf .ProchainNumPosteChargementSiAnodisationAutomatique > 0 Then
                    .ProchainNumPosteChargement = .ProchainNumPosteChargementSiAnodisationAutomatique
                Else
                    .ProchainNumPosteChargement = 0
                End If
                
           End If
                
            '--- passer le pointeur de la zone de la gamme d'anodisation � 1 (autorisation de lancement) ---
            If .ProchainNumPosteChargement > 0 Then
                NumCharge = TEtatsPostes(.ProchainNumPosteChargement).NumCharge
                If TEtatsCharges(NumCharge).PtrZoneGammeAnodisation = 0 Then
                    TEtatsCharges(NumCharge).PtrZoneGammeAnodisation = 1
                End If
            End If
        
        Else
        
            '**********************************************************************************************************
            '*  IL Y A DES CHARGES EN LIGNE (les charges au chargement, d�chargement ne comptent pas)
            '**********************************************************************************************************
        
            '--- v�rification du travail du pont 1 sur la partie pr�paration ---
            'si le pont 1 n'a plus de transfert � faire on peut lancer une gamme
            'contr�le l'occupation des postes pour �viter les conflits de postes et de lib�ration des ponts
            If EntreeAutomatiqueCharges = True Then
                VerificationLignePourEntreeCharge
            End If
        
        End If
    
    End With
    
 
    
    '**********************************************************************************************************
    '**********************************************************************************************************
    '*                   Analyse des charges en ligne pour d�terminer les mouvements des ponts
    '**********************************************************************************************************
    '**********************************************************************************************************
    For a = POSTES.P_CHGT_1 To DERNIER_POSTE

        '--- affectation du n� de charge ---
        NumCharge = TEtatsPostes(a).NumCharge

        If NumCharge >= CHARGES.C_NUM_MINI And NumCharge <= CHARGES.C_NUM_MAXI Then

            With TEtatsCharges(NumCharge)
                
                '----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
                
                '**********************************************************************************************************
                '*                                 Le pointeur de la zone de la gamme d'anodisation est � 0
                '**********************************************************************************************************
                'If .PtrZoneGammeAnodisation = 0 Then
                    'la charge est au chargement mais le moteur d'inf�rence d�cide de ne pas la rentrer en ligne
                    'car ce n'est pas le bon moment
                    'ou un num�ro de charge a �t� tap� par l'op�rateur dans la ligne sans une gamme
                'End If
                
                '**********************************************************************************************************
                '*                 Extraction de la zone de d�part et d'arriv�e ainsi que du temps d'�gouttage
                '**********************************************************************************************************
                If .PtrZoneGammeAnodisation > 0 Then    'ne prendre en compte que lorsque le pointeur est sup�rieur � 0
                    
                    '--- affectation ---
                    NumZoneDepart = .TGammesAnodisation.TDetailsGammesAnodisation(.PtrZoneGammeAnodisation).NumZone
                    NumZoneArrivee = .TGammesAnodisation.TDetailsGammesAnodisation(Succ(.PtrZoneGammeAnodisation)).NumZone
                    TempsEgouttageSecondes = .TGammesAnodisation.TDetailsGammesAnodisation(.PtrZoneGammeAnodisation).TempsEgouttageSecondes
                    DelaiSupStabilisationChargeSecondes = .DelaiSupStabilisationChargeSecondes
                    
                    '--- affichage des zones dans les renseignements ---
                    If NumZoneDepart > 0 And NumZoneArrivee > 0 Then
                        AfficheRenseignements VERT_4, "Charge " & NumCharge & _
                                                                              " - Zone d�part " & TZones(NumZoneDepart).Codezone & _
                                                                              " Zone arriv�e " & TZones(NumZoneArrivee).Codezone & vbCrLf
                    End If
                
                End If
                
                '----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
                
                '**********************************************************************************************************
                '*                                 Le pointeur de la zone de la gamme d'anodisation est � 1
                '**********************************************************************************************************
                If .PtrZoneGammeAnodisation = 1 And a >= POSTES.P_CHGT_1 And a <= POSTES.P_CHGT_2 Then
                    'le moteur d'inf�rence vient de d�cider l'entr�e d'une charge dans la ligne
                    'tous les cas de mouvements des ponts ont �t� analys� dans la gestion du chargement
                    'il faut donc rentrer cette charge le plus t�t possible
                    If NumZoneDepart > 0 And NumZoneArrivee > 0 Then
                
                        '--- d�termination du num�ro du poste de d�part et d'arriv�e ---
                        If TEtatsPostes(a).Condamnation = True Then
                            NumPosteDepart = 0                                          'si le poste est condamn� il ne faut pas lancer la charge
                        Else
                            NumPosteDepart = a
                        End If
                        NumPosteArrivee = ProchainNumeroPosteValide(NumCharge, NumZoneArrivee, True)
                        
                        '--- analyse uniquement si les deux num�ros de postes sont d�termin�s ---
                        If NumPosteDepart > 0 And NumPosteArrivee > 0 Then
                            
                            '--- gestion de l'anti-collision ---
                            ReponseAntiCollision = ControleAntiCollision(TPremisses(NumPosteDepart, NumPosteArrivee).NumPontIA, _
                                                                                                      NumPosteDepart, _
                                                                                                      NumPosteArrivee, _
                                                                                                      TypeCollision, _
                                                                                                      NumPontOppose, _
                                                                                                      NumPosteAssurantSecurite, _
                                                                                                      CouleurReponse)
                            
                            If TypeCollision = TYPES_COLLISION.AUCUN_RISQUE Then
                                
                                '************************************************************************
                                '               Contr�le des bains prioritaires avant transfert
                                '************************************************************************
                                ReponseControleBainsPrioritaires = ControleBainsPrioritaires(NumPontImpose:=NumPontPrioritaire, _
                                                                                                                                      NumPosteDepart:=NumPosteDepart, _
                                                                                                                                      NumPosteArrivee:=NumPosteArrivee, _
                                                                                                                                      CouleurReponse:=CouleurReponse)
                                AfficheRenseignements CouleurReponse, ReponseControleBainsPrioritaires & vbCrLf
                                
                                If ReponseControleBainsPrioritaires = OK Then
                                
                                    '************************************************************************
                                    '            Pas de risque de collision / lancement de transfert
                                    '************************************************************************
                                    ReponseTransfertCharge = AutomatiqueTransfertCharge(NumPontImpose:=NumPontPrioritaire, _
                                                                                                                                NumPosteDepart:=NumPosteDepart, _
                                                                                                                                NumPosteArrivee:=NumPosteArrivee, _
                                                                                                                                TempsEgouttageSecondes:=TempsEgouttageSecondes, _
                                                                                                                                DelaiSupStabilisationChargeSecondes:=DelaiSupStabilisationChargeSecondes, _
                                                                                                                                NumPontReelDuTransfert:=NumPontReelDuTransfert, _
                                                                                                                                CouleurReponse:=CouleurReponse)
                                    AfficheRenseignements CouleurReponse, ReponseTransfertCharge & vbCrLf
    
                                    '--- affectation ---
                                    TravailAvecMI(NumPontReelDuTransfert) = True     'signaler dans ce cas le travail avec le moteur d'inf�rence
                                    
                                    '************************************************************************
                                    ' Construction du prochain cycle en fonction du r�sultat du transfert
                                    '************************************************************************
                                    If ReponseTransfertCharge = OK Then
                                    
                                        'SZP 2021
                                        'on initialise le temps de traitement de la charge
                                        TEtatsCharges(NumCharge).DateEntreeEnLigne = Now
                                        
                                        TEtatsCharges(NumCharge).FinPhase4 = False
                                        
                                        Bidon = ConstruitProchainCyclePont(ViderProchainCycle:=False, _
                                                                                                    TypeCycle:=TC_TRANSFERT_CHARGE, _
                                                                                                    NumPont:=NumPontReelDuTransfert, _
                                                                                                    NumPosteDepart:=NumPosteDepart, _
                                                                                                    NumPosteArrivee:=NumPosteArrivee)
                                    
                                  
                                    End If
                                                                
                                End If
                                
                            Else
                                
                                '--- affichage dans les renseignements ---
                                AfficheRenseignements CouleurReponse, ReponseAntiCollision & vbCrLf
                            
                                '************************************************************************
                                '            Risque de collision / d�placement du pont oppos�
                                '************************************************************************
                                'cas extr�me se produisant si le pont 2 est tr�s proche du chargement
                                If NumPontOppose > 0 And NumPosteAssurantSecurite > 0 Then
                                    TravailAvecMI(NumPontOppose) = True     'signaler dans ce cas le travail avec le moteur d'inf�rence
                                    Bidon = ConstruitProchainCyclePont(ViderProchainCycle:=False, _
                                                                                                TypeCycle:=TC_DEPLACEMENT_PONT, _
                                                                                                NumPont:=NumPontOppose, _
                                                                                                NumPosteDepart:=TEtatsPonts(NumPontOppose).PosteActuel, _
                                                                                                NumPosteArrivee:=NumPosteAssurantSecurite)
                                    ReponseDeplacementPont = AutomatiqueDeplacementPont(NumPontOppose, NumPosteAssurantSecurite, CouleurReponse)
                                    
                                    Call LogPourCPO("Cas extr�me du d�placement du PONT " & NumPontOppose & " car risque de collision" & Chr(13) & "ReponseDeplacementPont=" & ReponseDeplacementPont)
                                    AfficheRenseignements CouleurReponse, ReponseDeplacementPont & vbCrLf
                             
                                End If
                                
                            End If
                        
                        End If
                        
                    End If
                
                End If
                
                '----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
                
                '**********************************************************************************************************
                '*                                Le pointeur de la zone de la gamme d'anodisation est � X
                '**********************************************************************************************************
                If .PtrZoneGammeAnodisation > 1 Then
                    'les gammes se d�roulent normalement, il faut g�rer les ponts afin de repartir chaque transfert
                    'et �viter les collisions
                    
                                       
                    If NumZoneDepart > 0 And NumZoneArrivee > 0 Then
                
                        '--- d�termination du num�ro du poste de d�part et d'arriv�e ---
                        If NumZoneDepart = NumZoneArrivee Then
                            
                            '--- la zone de d�part est �gale � celle d'arriv�e ---
                            'dans ce cas attendre la fin du temps au poste et forcer poste de d�part = poste d'arriv�e
                            NumPosteDepart = ProchainNumeroPosteValide(NumCharge, NumZoneDepart, False)
                            NumPosteArrivee = NumPosteDepart
                        
                        Else
                        
                            '--- d�termination du num�ro du poste de d�part et d'arriv�e ---
                            NumPosteDepart = ProchainNumeroPosteValide(NumCharge, NumZoneDepart, False)
                            NumPosteArrivee = ProchainNumeroPosteValide(NumCharge, NumZoneArrivee, True)
                                                  
                        End If
                        
                        '--- analyse uniquement si les deux num�ros de postes sont d�termin�s ---
                        If NumPosteDepart > 0 And NumPosteArrivee > 0 Then

                            '--- gestion de l'anti-collision ---
                            ReponseAntiCollision = ControleAntiCollision(TPremisses(NumPosteDepart, NumPosteArrivee).NumPontIA, _
                                                                                                     NumPosteDepart, _
                                                                                                     NumPosteArrivee, _
                                                                                                     TypeCollision, _
                                                                                                     NumPontOppose, _
                                                                                                     NumPosteAssurantSecurite, _
                                                                                                     CouleurReponse)
                            
                            If TypeCollision = TYPES_COLLISION.AUCUN_RISQUE Then
                                
                                '************************************************************************
                                '               Contr�le des bains prioritaires avant transfert
                                '************************************************************************
                                ReponseControleBainsPrioritaires = ControleBainsPrioritaires(NumPontImpose:=NumPontPrioritaire, _
                                                                                                                                      NumPosteDepart:=NumPosteDepart, _
                                                                                                                                      NumPosteArrivee:=NumPosteArrivee, _
                                                                                                                                      CouleurReponse:=CouleurReponse)
                                AfficheRenseignements CouleurReponse, ReponseControleBainsPrioritaires & vbCrLf
                                
                                If ReponseControleBainsPrioritaires = OK Then
                                
                                    '************************************************************************
                                    '            Pas de risque de collision / lancement de transfert
                                    '************************************************************************
                                    ReponseTransfertCharge = AutomatiqueTransfertCharge(NumPontImpose:=NumPontPrioritaire, _
                                                                                                                                NumPosteDepart:=NumPosteDepart, _
                                                                                                                                NumPosteArrivee:=NumPosteArrivee, _
                                                                                                                                TempsEgouttageSecondes:=TempsEgouttageSecondes, _
                                                                                                                                DelaiSupStabilisationChargeSecondes:=DelaiSupStabilisationChargeSecondes, _
                                                                                                                                NumPontReelDuTransfert:=NumPontReelDuTransfert, _
                                                                                                                                CouleurReponse:=CouleurReponse)
                                    AfficheRenseignements CouleurReponse, ReponseTransfertCharge & vbCrLf
                                    
                                    '************************************************************************
                                    'Affectation de la date du dernier transfert de charge en fonction du pont
                                    '************************************************************************
                                    If NumPontReelDuTransfert >= PONTS.P_1 And NumPontReelDuTransfert <= PONTS.P_2 Then
                                        TDatesDerniersTransfertsCharges(NumPontReelDuTransfert) = Now
                                    End If
                                    
                                    '************************************************************************
                                    ' Construction du prochain cycle en fonction du r�sultat du transfert
                                    '************************************************************************
                                    '--- v�rifier avec la premi�re sortie dans l'ordre de sortie des charges ---
                                    
                                    'Call Log("NumPosteDepart: " + NumPosteDepart + ", TMoteurInference.TOrdreSortieCharges(1).NumPoste:" + TMoteurInference.TOrdreSortieCharges(1).NumPoste)
                                    If NumPosteDepart = TMoteurInference.TOrdreSortieCharges(1).NumPoste Then
                                        If MemReponseTransfertCharge <> ReponseTransfertCharge Then
                                            Bidon = ConstruitProchainCyclePont(ViderProchainCycle:=False, _
                                                                                                        TypeCycle:=TC_TRANSFERT_CHARGE, _
                                                                                                        NumPont:=NumPontReelDuTransfert, _
                                                                                                        NumPosteDepart:=NumPosteDepart, _
                                                                                                        NumPosteArrivee:=NumPosteArrivee)
                                            MemReponseTransfertCharge = NumPontReelDuTransfert & "/" & NumPosteDepart & "/" & NumPosteArrivee
                                        End If
                                    Else
                                       ' Call Log(" NO BUILD CYCLE PONT")
                                    End If
                                
                                End If
                                
                            Else
                            
                               '--- affichage du message ---
                               AfficheRenseignements CouleurReponse, ReponseAntiCollision & vbCrLf
        
                               'Call Log("Risque collision, d�placement  pont  oppos�: " & NumPontOppose & " , poste s�curit�: " & NumPosteAssurantSecurite)
                               
                               '************************************************************************
                               '            Risque de collision / d�placement du pont oppos�
                               '************************************************************************
                               If NumPontOppose > 0 And NumPosteAssurantSecurite > 0 Then
                                    TravailAvecMI(NumPontOppose) = True     'signaler dans ce cas le travail avec le moteur d'inf�rence
                                    Bidon = ConstruitProchainCyclePont(ViderProchainCycle:=False, _
                                                                                                TypeCycle:=TC_DEPLACEMENT_PONT, _
                                                                                                NumPont:=NumPontOppose, _
                                                                                                NumPosteDepart:=TEtatsPonts(NumPontOppose).PosteActuel, _
                                                                                                NumPosteArrivee:=NumPosteAssurantSecurite)
                                    ReponseDeplacementPont = AutomatiqueDeplacementPont(NumPontOppose, _
                                                                                                                                    NumPosteAssurantSecurite, _
                                                                                                                                    CouleurReponse)
                                    
                                    Call LogPourCPO("D�placement du PONT " & NumPontOppose & " car risque de collision" & Chr(13) & "ReponseDeplacementPont=" & ReponseDeplacementPont)
                                    AfficheRenseignements CouleurReponse, ReponseDeplacementPont & vbCrLf
                                
                                    '************************************************************************
                                    'affectation de la date du dernier d�placement en fonction du pont
                                    '************************************************************************
                                    If NumPontOppose >= PONTS.P_1 And NumPontOppose <= PONTS.P_2 Then
                                        TDatesDerniersDeplacementsAVide(NumPontOppose) = Now
                                    End If
                                Else
                                   'Call Log("ici 1")
                               
                               End If

                            End If
                        
                        End If
                    
                    End If
                
                End If
                
                '----------------------------------------------------------------------------------------------------------
                
                '**********************************************************************************************************
                '*                       D�placement des ponts avant le terme du temps au poste de prise
                '**********************************************************************************************************
                
                'Call Log("ReponseTransfertCharge =" & ReponseTransfertCharge & Chr(13) & "ReponseDeplacementPont=" & ReponseDeplacementPont & Chr(13) & _
                '    "ReponseAntiCollision=" & ReponseAntiCollision, logMoteurInference)
                'Call Log(".PtrZoneGammeAnodisation=" & .PtrZoneGammeAnodisation & "NumZoneDepart =" & NumZoneDepart & ", NumZoneArrivee=" & NumZoneArrivee)
                'With TMoteurInference.TOrdreSortiePonts(PONTS.P_1, 1)
                '    Call Log("d�compte sortie P1:" & .DecompteDuTempsAuPosteReelSecondes & ", numposte � sortir:" & .NumPoste & ", poste actuel:" & TEtatsPonts(PONTS.P_1).PosteActuel & Chr(13) & _
                '    ", depuis DatesDerniersTransfertsCharges P1 =" & DateDiff("s", TDatesDerniersTransfertsCharges(PONTS.P_1), Now) & Chr(13) & _
                '    ", depuis  DatesDerniersDeplacementsAVide P1=" & DateDiff("s", TDatesDerniersDeplacementsAVide(PONTS.P_1), Now), logMoteurInference)
                'End With
                'With TMoteurInference.TOrdreSortiePonts(PONTS.P_2, 1)
                '    Call Log("d�compte sortie P2:" & .DecompteDuTempsAuPosteReelSecondes & ", numposte � sortir:" & .NumPoste & ", poste actuel:" & TEtatsPonts(PONTS.P_2).PosteActuel & Chr(13) & _
                '    ", depuis DatesDerniersTransfertsCharges P2=" & DateDiff("s", TDatesDerniersTransfertsCharges(PONTS.P_2), Now) & Chr(13) & _
                '    ", depuis  DatesDerniersDeplacementsAVide P2=" & DateDiff("s", TDatesDerniersDeplacementsAVide(PONTS.P_2), Now), logMoteurInference)
                'End With
                
                'Call Log("TEtatsPonts(PONTS.P_1).PosteActuel=" & TEtatsPonts(PONTS.P_1).PosteActuel & " NumZoneArrivee=" & NumZoneArrivee & _
                '    " dateDiff(s, TDatesDerniersTransfertsCharges(.NumPont), Now)=" & DateDiff("s", TDatesDerniersTransfertsCharges(1), Now))
                
                
                  If .PtrZoneGammeAnodisation > 1 Then
                    'SZP 20241014
                    With TMoteurInference.TOrdreSortiePonts(PONTS.P_1, 1)
                    If NumZoneDepart = NUMZONE_ANO Then
                    'If .NumPont = PONTS.P_1 Then
                    'If IsNumeric(.DecompteDuTempsAuPosteReelSecondes) = True Then
                    If TEtatsPonts(PONTS.P_1).PosteActuel > POSTES.P_C12 Then
                    If DateDiff("s", TDatesDerniersTransfertsCharges(1), Now) >= 15 Then
                    
                       
                        Bidon = ConstruitProchainCyclePont(ViderProchainCycle:=False, _
                                                                                                TypeCycle:=TC_DEPLACEMENT_PONT, _
                                                                                                NumPont:=PONTS.P_1, _
                                                                                                NumPosteDepart:=TEtatsPonts(PONTS.P_1).PosteActuel, _
                                                                                                NumPosteArrivee:=POSTES.P_C08)
                       
                        ReponseDeplacementPont = AutomatiqueDeplacementPont(PONTS.P_1, POSTES.P_C08, CouleurReponse)
                        
                        'Call Log("D�placement du PONT 1 en C08 avant terme du temps en ANODISATION" & Chr(13) & "ReponseDeplacementPont=" & ReponseDeplacementPont, logMoteurInference)
                        AfficheRenseignements CouleurReponse, ReponseDeplacementPont & vbCrLf
                        
                        'TEtatsPonts(PONTS.P_1).PosteActuel = P_C08
                        
                        '************************************************************************
                        'affectation de la date du dernier d�placement du PONT 1
                        '************************************************************************
                        TDatesDerniersDeplacementsAVide(PONTS.P_1) = Now
                    
                    End If
                    'End If
                    'End If
                    End If
                    End If
                    End With
                    
                    With TMoteurInference.TOrdreSortiePonts(PONTS.P_2, 1)
                    If NumZoneDepart = NUMZONE_ANO Then
                    If TEtatsPonts(PONTS.P_2).PosteActuel = .NumPoste Then
                    If TEtatsPonts(PONTS.P_2).PosteActuel >= P_C13 And TEtatsPonts(PONTS.P_2).PosteActuel <= P_C15 Then
                    If IsNumeric(.DecompteDuTempsAuPosteReelSecondes) = True Then
                    If CLng(.DecompteDuTempsAuPosteReelSecondes) <= 30 Then
                        Dim NomGroupe As String
                        
                        Dim NumChargeRed As Integer
                        
                        NumChargeRed = -1
                        
                        '--- d�claration ---
                        Dim ValeurRetourneeAPI As Long          'valeur retourn�e par une fonction concernant le dialogue avec l'automate
                        Dim NomVariable As String
                                   
     
                        If POSTES.P_C13 = .NumPoste Then
                            NumChargeRed = TEtatsRedresseurs(1).NumCharge
                        End If
                        If POSTES.P_C14 = .NumPoste Then
                            NumChargeRed = TEtatsRedresseurs(2).NumCharge
                        End If
                        If POSTES.P_C15 = .NumPoste Then
                            NumChargeRed = TEtatsRedresseurs(3).NumCharge
                        End If
                        
                        If NumChargeRed <> -1 Then
                        
                            'il faut d�j� enregistr� les donn�es redresseurs car
                            'elles vont �tre perdues ensuite
                            enregistreRedresseursAno NumChargeRed, .NumPoste
                            
                            NomGroupe = "CHARGE_" & Right("00" & NumChargeRed, 2)
                            
                            Call Log("HORS TENSION: " & Chr(13) & NomGroupe, logMoteurInference)
                                    
                            '--- �criture dans l'automate ---
                            ValeurRetourneeAPI = APIEcritureVariableNommee(NomGroupe, "UPhase4", 0)
                            If ValeurRetourneeAPI <> 0 Then
                                Bidon = MessageErreur("erreur couper redreseur", MESSAGE_500)             'lancer un message d'erreur
                            End If
                             ValeurRetourneeAPI = APIEcritureVariableNommee(NomGroupe, "IPhase4", 0)
                            If ValeurRetourneeAPI <> 0 Then
                                Bidon = MessageErreur("erreur couper redreseur", MESSAGE_500)             'lancer un message d'erreur
                            End If
                            ValeurRetourneeAPI = APIEcritureVariableNommee(NomGroupe, "TpsPhase4", 0)
                            If ValeurRetourneeAPI <> 0 Then
                                Bidon = MessageErreur("erreur couper redreseur", MESSAGE_500)             'lancer un message d'erreur
                            End If
                        Else
                            'Call Log("IMPOSSIBLE DE TROUVER LA CHARGE ERREUR COUPER TENSION REDRESSEUR ")
                        End If
                        
                        
                       
                    End If
                    End If
                    End If
                    End If
                    End If
                    End With
                  
                    ' FIN ---------------------------------------------------------------------
                   
                    
                    
                    If ReponseTransfertCharge <> OK And ReponseDeplacementPont <> OK And ReponseAntiCollision <> OK Then
                        
                        'les gammes se d�roulent normalement, il faut g�rer les ponts afin de d�placer chaque pont avant le terme
                        'du temps au poste de prise et �viter les collisions
                        If NumZoneDepart > 0 And NumZoneArrivee > 0 Then
                                
                            '**********************************************************************************************************
                            '*                                         D�placement du PONT 1 avant terme du temps
                            '**********************************************************************************************************
                            If TEtatsPonts(PONTS.P_1).SensX = S_AU_POSTE Then
                            
                                With TMoteurInference.TOrdreSortiePonts(PONTS.P_1, 1)
                                    
                                    If .NumPont = PONTS.P_1 Then
                                    If IsNumeric(.DecompteDuTempsAuPosteReelSecondes) = True Then
                                    If (TEtatsPonts(PONTS.P_1).PosteActuel <> .NumPoste) Then
                                    If DateDiff("s", TDatesDerniersTransfertsCharges(.NumPont), Now) >= 40 Then
                                    If DateDiff("s", TDatesDerniersDeplacementsAVide(.NumPont), Now) >= 20 Then
                                                    
                                                        
                                        '--- gestion de l'anti-collision ---
                                        ReponseAntiCollision = ControleAntiCollision(PONTS.P_1, TEtatsPonts(PONTS.P_1).PosteActuel, _
                                                                                                .NumPoste, TypeCollision, NumPontOppose, _
                                                                                                NumPosteAssurantSecurite, CouleurReponse)
                        
                                        '--- aucun risque alors d�placement du pont ---
                                        If TypeCollision = TYPES_COLLISION.AUCUN_RISQUE Then
                                            
                                            ReponseDeplacementPont = AutomatiqueDeplacementPontOptimisation(PONTS.P_1, .NumPoste, CouleurReponse)
                                            'Call Log("D�placement du PONT 1 avant terme du temps -  ReponseDeplacementPont=" & ReponseDeplacementPont, logMoteurInference)
                                            AfficheRenseignements CouleurReponse, ReponseDeplacementPont & vbCrLf
                                        
                                            '************************************************************************
                                            'affectation de la date du dernier d�placement du PONT 1
                                            '************************************************************************
                                            TDatesDerniersDeplacementsAVide(PONTS.P_1) = Now
                                        Else
                                            Call LogPourCPO("avant terme du temps:Risque collision avec P2 " & Chr(13) & "ReponseAntiCollision=" & ReponseAntiCollision)
                                        
                                        End If
                                             
                                    End If
                                    End If
                                    End If
                                    End If
                                    End If
                                
                                End With
                            
                            End If
                            
                            '**********************************************************************************************************
                            '*                                         D�placement du PONT 2 avant terme du temps
                            '**********************************************************************************************************
                            If TEtatsPonts(PONTS.P_2).SensX = S_AU_POSTE Then
                            
                                With TMoteurInference.TOrdreSortiePonts(PONTS.P_2, 1)
                                
                                    If IsNull(TDatesDerniersTransfertsCharges(.NumPont)) Then
                                        TDatesDerniersTransfertsCharges(.NumPont) = DateAdd("m", -3, Date)
                                    End If
                                    If IsNull(TDatesDerniersTransfertsCharges(.NumPont)) Then
                                        TDatesDerniersDeplacementsAVide(.NumPont) = DateAdd("m", -3, Date)
                                    End If
                                
                                   
                                    If .NumPont = PONTS.P_2 Then
                                    If IsNumeric(.DecompteDuTempsAuPosteReelSecondes) = True Then
                                    If CLng(.DecompteDuTempsAuPosteReelSecondes) <= 40 Then
                                    If TEtatsPonts(PONTS.P_2).PosteActuel <> .NumPoste Then
                                    If DateDiff("s", TDatesDerniersTransfertsCharges(.NumPont), Now) >= 40 Then
                                    If DateDiff("s", TDatesDerniersDeplacementsAVide(.NumPont), Now) >= 20 Then
                                        
                                        
                                        'Call Log("D�placement du PONT 2 avant terme du temps - deb check collision", logMoteurInference)
                                        '--- gestion de l'anti-collision ---
                                        ReponseAntiCollision = ControleAntiCollision(PONTS.P_2, TEtatsPonts(PONTS.P_2).PosteActuel, .NumPoste, TypeCollision, _
                                                                                    NumPontOppose, NumPosteAssurantSecurite, CouleurReponse)
                        
                                        
                                        'TypeCollision = TYPES_COLLISION.AUCUN_RISQUE
                                        '--- aucun risque alors d�placement du pont ---
                                        If TypeCollision = TYPES_COLLISION.AUCUN_RISQUE Then
                                        
                                            
                                            ReponseDeplacementPont = AutomatiqueDeplacementPontOptimisation(PONTS.P_2, .NumPoste, CouleurReponse)
                                            AfficheRenseignements CouleurReponse, ReponseDeplacementPont & vbCrLf
                                            'Call Log("D�placement du PONT 2 avant terme du temps - OK -ReponseDeplacementPont:" & ReponseDeplacementPont, logMoteurInference)
                                            
                                            '************************************************************************
                                            'affectation de la date du dernier d�placement du PONT 2
                                            '************************************************************************
                                            TDatesDerniersDeplacementsAVide(PONTS.P_2) = Now
                                        Else
                                            Call LogPourCPO("avant terme du temps: risque collision avec P1 " & Chr(13) & "ReponseAntiCollision=" & ReponseAntiCollision)
                                        End If
                                        
                                    End If
                                    End If
                                    End If
                                    End If
                                    End If
                                    End If
                                
                                End With
                        
                            End If
                        
                        End If
                        
                    End If
            
                End If
        
            End With
        
        End If

    Next a
        
    '********************************************************************************************
    '********************************************************************************************
    '*                                           Forcer le type de s�quence
    '********************************************************************************************
    '********************************************************************************************
    'par d�faut en automatique la ligne suit pr�cis�ment les gammes d'anodisation (mode
    'cyclique optimis�), le moteur d'inf�rence intervient dans le choix du travail de chaque
    'pont pour l'anti-collision, les transferts et l'optimisation des charges en entr�e de ligne
    For a = LBound(TEtatsPonts()) To UBound(TEtatsPonts())
        With TEtatsPonts(a)
            If .ModePont = MODES_PONTS.M_AUTOMATIQUE Then
                If TravailAvecMI(a) = True Then
                    .TypeSequence = TYPES_SEQUENCES.TS_ALEATOIRE
                Else
                    .TypeSequence = TYPES_SEQUENCES.TS_CYCLIQUE_OPTIMISE
                End If
            End If
        End With
    Next a
    
End Sub

'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' R�le      : Contr�le la gestion de la pompe en fonction du programmateur cyclique
' Entr�es : NumCuve -> Num�ro de la cuve trait�e
' Retours :
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Sub AutomatiquePompe(ByVal NumCuve As Integer)
    
    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next
    
    '--- constantes priv�es ---
    Const NOM_GROUPE As String = "CYCLES_AUTO_POMPES_CUVES"
    
    '--- d�claration ---
    Dim ValeurRetourneeAPI As Long                  'valeur retourn�e par une fonction concernant le dialogue avec l'automate
    Dim NomVariable As String                            'nom de la variable
    Dim NumCuveAutomate As Integer
    '--- analyse en fonction du PC ---
    'If TypePC <> TYPES_PC. Then Exit Sub

    With TEtatsCuves(NumCuve)
         NumCuveAutomate = .IndexAutomate
        '--- cuves avec pompe ---
        If .API_CyclePompe <> .CyclePompe Or PremierPassageNoyauCentral = False Then
            
            '--- �criture dans l'API ---
            If PROGRAMME_AVEC_AUTOMATE = True Then

                '--- affectation du nom de la variable ---
                NomVariable = "CycleAutoPompePCCuve" & Right("0" & NumCuveAutomate, 2)
                
                '--- �criture ---
                ValeurRetourneeAPI = APIEcritureVariableNommee(NOM_GROUPE, NomVariable, .CyclePompe)

            End If

        End If

    End With
    
End Sub

'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' R�le      : Contr�le la gestion du chauffage en fonction du programmateur cyclique
' Entr�es : NumCuve -> Num�ro de la cuve trait�e
' Retours :
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Sub AutomatiqueChauffage(ByVal NumCuve As Integer)
    
    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next
    
    '--- constantes priv�es ---
    Const NOM_GROUPE As String = "MODES_CHAUFFAGES_CUVES"
    
    '--- d�claration ---
    Dim ValeurRetourneeAPI As Long                  'valeur retourn�e par une fonction concernant le dialogue avec l'automate
    Dim NomVariable As String                            'nom de la variable
    
    '--- analyse en fonction du PC ---
    'If TypePC <> TYPES_PC. Then Exit Sub


    Dim NumCuveAutomate As Integer
    With TEtatsCuves(NumCuve)
        
        
    NumCuveAutomate = .IndexAutomate
        '--- mode de production ---
        If .API_ModeProduction <> .ModeProduction Or PremierPassageNoyauCentral = False Then
            
            '--- �criture dans l'API ---
            If PROGRAMME_AVEC_AUTOMATE = True Then

                '--- affectation du nom de la variable ---
                NomVariable = "ModeChauffagePCCuve" & Right("0" & NumCuveAutomate, 2)
                
                '--- �criture ---
                ValeurRetourneeAPI = APIEcritureVariableNommee(NOM_GROUPE, NomVariable, .ModeProduction)

            End If

        End If

    End With

End Sub

'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' R�le        : Signalisation des d�fauts sur le gyrophare et le klaxon
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Sub SignalisationDefautsGyrophareKlaxonVersAPI()
    
    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next

    '--- d�claration ---

    '--- analyse en fonction du PC ---
    'If TypePC <> TYPES_PC. Then Exit Sub

    If PROGRAMME_AVEC_AUTOMATE = True Then
    
        If SignalerDefautSurGyrophare = True Or SignalerDefautSurKlaxon = True Then
    
            '--- �criture dans l'API pour enclencher le gyrophare ---
            If SignalerDefautSurGyrophare = True Then
                Bidon = APIEcritureVariableNommee("DEFAUTS", "M_Dem_PC_Gyrophare", True)                  'pour le gyrophare
                SignalerDefautSurGyrophare = False
            End If
                
            '--- �criture dans l'API pour enclencher le klaxon ---
            If SignalerDefautSurKlaxon = True Then
                Bidon = APIEcritureVariableNommee("DEFAUTS", "M_Dem_PC_Klaxon", True)                        'pour le klaxon
                SignalerDefautSurKlaxon = False
            End If
                
        End If
    
    End If
    
End Sub

'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' R�le      : Effectue l'enregistrement des valeurs U, I et de la temp�rature + divers valeurs
' Entr�es :
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Sub EffectueTra�abiliteRedresseurs()
    
    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next

    '--- d�claration ---
    Static PremierPassage As Boolean
    
    Static TEnregistrementDesPoints(REDRESSEURS.R_C13 To REDRESSEURS.R_C16) As Boolean              'tableau contenant l'enregistrement des points
    
    Dim a As Integer                                                                                                                                                'pour les boucles FOR...NEXT
    
    Dim NumCuve As Integer                                                                                                                                  'repr�sente un num�ro de cuve
    
    Dim UEnCours As Integer                                                                                                                                  'tension en cours � l'instant t
    Dim IEnCours As Integer                                                                                                                                    'intensit� en cours � l'instant t
    
    Static TNumChargesRedresseurs(REDRESSEURS.R_C13 To REDRESSEURS.R_C16) As Integer                'tableau des num�ros de charges des redresseurs

    Static TPtrPointsTra�abilite(REDRESSEURS.R_C13 To REDRESSEURS.R_C16) As Long                              'tableau des pointeur des points � tracer
    
    Dim TemperatureActuelle As Single                                                                                                                  'temperature actuelle d'une cuve
    
    Static TValeursInferieures(REDRESSEURS.R_C13 To REDRESSEURS.R_C16) As Single                              'tableau des valeurs inf�rieures
    Static TValeursSuperieures(REDRESSEURS.R_C13 To REDRESSEURS.R_C16) As Single                            'tableau des valeurs sup�rieures
    
    Static TNomsFichiersTra�abilite(REDRESSEURS.R_C13 To REDRESSEURS.R_C16) As String                     'm�morisation des noms des fichiers de tra�abilit�
    
    Dim TTra�abilite As Tra�abilite                                                                                                                          'tableau image de la tra�abilit�
    Static TPointsTra�abilite(REDRESSEURS.R_C13 To REDRESSEURS.R_C16) As Tra�abilite                          'tableau contenant les points de la tra�abilit�
    Static TMemPointsTra�abilite(REDRESSEURS.R_C13 To REDRESSEURS.R_C16) As Tra�abilite                  'tableau m�moire des points de la tra�abilit�
    
    '----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    
    For a = REDRESSEURS.R_C13 To REDRESSEURS.R_C15
    
        With TEtatsRedresseurs(a)
        
            '----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
            
            '--- une charge vient de rentrer dans le poste, le redresseur se met en marche ---
            If .NumCharge > CHARGES.PAS_DE_CHARGE And _
               TNumChargesRedresseurs(a) = CHARGES.PAS_DE_CHARGE And _
               .NumCharge <> TNumChargesRedresseurs(a) And _
               .DebutCycle = True And _
               .ControleFinCycle = False Then
    
                '--- m�morisation du num�ros de charges ---
                TNumChargesRedresseurs(a) = .NumCharge
                
                '--- affectation du nom du fichiers de tra�abilit� ---
                TNomsFichiersTra�abilite(a) = RepGraphesProductionLocal & "AnalyseRedresseurCharge" & Right("0" & TNumChargesRedresseurs(a), 2) & ".FIC"
    
                '--- vidange du fichier ---
                Close TCanauxFichiersTra�abilite(a)
                If FileExist(TNomsFichiersTra�abilite(a)) = True Then
                    'Kill TNomsFichiersTra�abilite(a)
                    Open TNomsFichiersTra�abilite(a) For Output As 1
                    Close 1
                End If
                
                '--- ouverture du fichier ---
                Open TNomsFichiersTra�abilite(a) For Random Shared As #TCanauxFichiersTra�abilite(a) Len = Len(TTra�abilite)

                '--- affectation ---
                TPtrPointsTra�abilite(a) = 1                    'premier point du trac�
                TPointsTra�abilite(a) = TTra�abilite       'affectation du tableau vide de la tra�abilit�
                
            End If
    
            '----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    
            '--- enregistrement des points ---
            If TNumChargesRedresseurs(a) >= CHARGES.C_NUM_MINI And TNumChargesRedresseurs(a) <= CHARGES.C_NUM_MAXI And _
               .DebutCycle = True And _
               .ControleFinCycle = False Then

                '--- num�ro de la phase ---
                If TPointsTra�abilite(a).NumPhase <> .NumPhaseEnCours Then
                    TEnregistrementDesPoints(a) = True
                End If

                '--- �tat du redresseur ---
                If TPointsTra�abilite(a).EtatRedresseur <> .EtatRedresseur Then
                    TEnregistrementDesPoints(a) = True
                End If

                '--- contr�le pour U ---
                UEnCours = .U * 10
                TValeursInferieures(a) = UEnCours * (1 - POURCENT_AVANT_TRACABILITE)
                TValeursSuperieures(a) = UEnCours * (1 + POURCENT_AVANT_TRACABILITE)
                With TPointsTra�abilite(a)
                    If .Tension < TValeursInferieures(a) Or .Tension > TValeursSuperieures(a) Then
                        TEnregistrementDesPoints(a) = True
                    End If
                End With
                
                '--- contr�le pour I ---
                IEnCours = .I
                TValeursInferieures(a) = IEnCours * (1 - POURCENT_AVANT_TRACABILITE)
                TValeursSuperieures(a) = IEnCours * (1 + POURCENT_AVANT_TRACABILITE)
                With TPointsTra�abilite(a)
                    If (.Intensite < TValeursInferieures(a) Or .Intensite > TValeursSuperieures(a)) Then
                        TEnregistrementDesPoints(a) = True
                    End If
                End With

                '--- contr�le pour la temp�rature ---
                NumCuve = CorrespondanceRedresseursCuvesAPI(a)
                TemperatureActuelle = TEtatsCuves(NumCuve).Temperatures.TempActuelle
                If TemperatureActuelle >= 0 And TemperatureActuelle <= 99 Then
                    With TPointsTra�abilite(a)
                        If .Temperature <> TemperatureActuelle * 10 Then
                            TEnregistrementDesPoints(a) = True
                        End If
                    End With
                End If
            
                '----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
                
                '--- enregistrement du point ---
                If TEnregistrementDesPoints(a) = True Then
                
                    '--- affectation de valeurs en cours ---
                    TPointsTra�abilite(a).DateDuPoint = Now
                    TPointsTra�abilite(a).NumPhase = .NumPhaseEnCours
                    TPointsTra�abilite(a).EtatRedresseur = .EtatRedresseur
                    TPointsTra�abilite(a).Tension = UEnCours
                    TPointsTra�abilite(a).Intensite = IEnCours
                    TPointsTra�abilite(a).Temperature = TemperatureActuelle * 10

                    '--- enregistrement dans le fichier ---
                    If TPtrPointsTra�abilite(a) < NBR_POINTS_MAXI_TRACABILITE Then
                    
                        '--- enregistrement dans le fichier ---
                        Put #TCanauxFichiersTra�abilite(a), TPtrPointsTra�abilite(a), TPointsTra�abilite(a)

                        '--- incr�mentation du pointeur pour le prochain point ---
                        Inc TPtrPointsTra�abilite(a)

                    End If

                    '--- anti-rebond d'enegistrement ---
                    TEnregistrementDesPoints(a) = False
                
                End If

            End If

            '----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

            '--- fermeture du fichier en fin d'anodisation ---
            If TNumChargesRedresseurs(a) <> CHARGES.PAS_DE_CHARGE Then
                
                If TPtrPointsTra�abilite(a) > 10 Then               'tra�age de 10 points au moins  (10 secondes de traitement)
                
                    If .DebutCycle = False And .ControleFinCycle = True Then
                    
                        '--- fermeture du fichier ---
                        Close TCanauxFichiersTra�abilite(a)
                    
                        '--- affectation de la charge vide ---
                        TNumChargesRedresseurs(a) = CHARGES.PAS_DE_CHARGE
    
                    End If

                End If

            End If
                
            '----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    
            '--- vider la m�moire des num�ros de charges des redresseurs si pas de charge dans le poste ---
            If .NumCharge = CHARGES.PAS_DE_CHARGE Then
                
                If TNumChargesRedresseurs(a) <> CHARGES.PAS_DE_CHARGE Then
                    
                    '--- fermeture du fichier ---
                    Close TCanauxFichiersTra�abilite(a)
                    
                    '--- affectation de la charge vide ---
                    TNumChargesRedresseurs(a) = CHARGES.PAS_DE_CHARGE
                
                End If
            
            End If
    
        End With
    
    Next a
    
End Sub



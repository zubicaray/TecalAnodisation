Attribute VB_Name = "MAlarmes"
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' R�le                    : MODULE AIDANT A LA GESTION DES ALARMES
' Nom                    : MAlarmes.bas
' Date de cr�ation : 26/06/2002
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

'--- d�clarations obligatoires ---
Option Explicit

'--- options g�n�rales ---
Option Base 1
DefVar A-Z

'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' R�le      : Acquitte les alarmes (coupe le gyrophare et le klaxon)
' Entr�es :
' Retours :
' D�tails  :
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Sub AcquittementAlarmes()

    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next

    '--- d�claration ---

    '--- analyse en fonction du PC ---
    'If TypePC <> TYPES_PC. Then Exit Sub
    
    '--- �criture dans l'API ---
    If PROGRAMME_AVEC_AUTOMATE = True Then
        Bidon = APIEcritureVariableNommee("DEFAUTS", "M_Acquittement_Defauts", True)                   'pour le gyrophare et le klaxon
    End If

End Sub

'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' R�le      : Attribution d'un num�ro de d�faut en fonction de chaque d�faut
' Entr�es :
' Retours :
' D�tails  :
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function AttributionNumDefauts() As String
    
    '--- aiguillage en cas d'erreurs ---
    On Error GoTo GestionErreurs
    
    '--- d�claration ---
    Dim a As Integer
            
    '--- affectation ---
    AttributionNumDefauts = ""
    
    '--- affichage du type de t�che ---
    AfficheTypeTache ("Attribution des num�ros de d�fauts")
    
    '*********************************************************************************************************************
    '*                                                                       ETATS DE LA LIGNE
    '*********************************************************************************************************************
    
    With TEtatsLigne.TNumDefauts
        
        '--- arr�t g�n�rale ---
        .NumDefautArretGeneral = 1
        
        '--- arr�t d'urgence ---
        .NumDefautArretUrgence = 2
        
        '--- portillons et ligne de vie ---
        .NumDefautPortillonsLigneVie = 3
        
        '--- s�curit� du pont 1 ---
        .NumDefautSecuriteP1 = 4
        
        '--- s�curit� du pont 2 ---
        .NumDefautSecuriteP2 = 5
            
        '--- manque de tension ---
        .NumDefautManqueTension = 6
            
        '--- manque d'air ---
        .NumDefautManqueAir = 7
    
        '--- stop ligne ---
        .NumDefautStopLigne = 8
    
    End With
    
    '*********************************************************************************************************************
    '*                                                                        PONTS
    '*********************************************************************************************************************
    
    For a = LBound(TEtatsPonts()) To UBound(TEtatsPonts())
    
        With TEtatsPonts(a).TNumDefauts
            
            '--- d�faut du variateur de la translation ---
            .NumDefautDefautVariateurTrLPont = Choose(a, 101, 201)
                
            '--- d�faut du variateur du levage ---
            .NumDefautDefautVariateurLevPont = Choose(a, 104, 204)
                
            '--- surcourse levage haut ---
            .NumDefautSurcourseLevHaut = Choose(a, 105, 205)
            
            '--- surcourse levage bas ---
            .NumDefautSurcourseLevBas = Choose(a, 106, 206)
                
            '--- axe non r�f�renc� de la translation ---
            .NumDefautAxeNonReferenceTrlPont = Choose(a, 110, 210)
            
            '--- axe non r�f�renc� du levage ---
            .NumDefautAxeNonReferenceLevPont = Choose(a, 111, 211)
                
            '--- anti-collision ---
            .NumDefautAntiCollision = Choose(a, 109, 209)
                
            '--- fin de zone ---
            .NumDefautFinDeZone = Choose(a, 110, 210)
                
            '--- pr�sence pi�ce ---
            .NumDefautPresencePiece = Choose(a, 115, 215)
                
            '--- d�faut laser ---
            .NumDefautDefautLaser = Choose(a, 116, 216)
        
            '--- d�lai trop long de descente des accroches de la charge ---
            .NumDefautDelaiTropLongDescenteAccroches = Choose(a, 120, 220)
            
            '--- d�lai trop long de mont�e des accroches de la charge ---
            .NumDefautDelaiTropLongMonteeAccroches = Choose(a, 121, 221)
        
        End With
    
    Next a
    
    '*********************************************************************************************************************
    '*                                                                        CUVE C00
    '*********************************************************************************************************************
    
    With TEtatsCuves(CUVES_REGULATION.C_C00).TNumDefauts

        '--- niveau tr�s bas ---
        .NumDefautNiveauTresBas = 300
        
        '--- niveau tr�s haut ---
        .NumDefautNiveauTresHaut = 301

        '--- d�faut chauffage ---
        .NumDefautDefautChauffage = 302

        '--- d�fauts temp�ratures ---
        .NumDefautDefautPT100 = 303
        .NumDefautTemperatureTropBasse = 304
        .NumDefautTemperatureTropHaute = 305
        
        
      
'        '--- disjonction des �lectro-vannes ouverture/fermeture des couvercles ---
'        .NumDefautDisjonctionEVCouvercles = 301
'
'        '--- disjonction chauffage ---
        '.NumDefautDisjonctionChauffage = 302
'
'        '--- disjonction �lectro-vanne d'arriv�e d'eau ---
'        .NumDefautDisjonctionEVEau = 304
'
'        '--- disjonction de la pompe ---
'        .NumDefautDisjonctionPompe = 305
'
'
'        '--- d�lai trop long d'ouverture des couvercles ---
'        .NumDefautDelaiTropLongOuvertureCouvercles = 312
'
'        '--- d�lai trop long de fermeture des couvercles ---
'        .NumDefautDelaiTropLongFermetureCouvercles = 313
'
'        '--- d�lai trop long de l'�lectro-vanne d'arriv�e d'eau ---
'        .NumDefautDelaiTropLongEVEau = 314
'
    End With
    
    '*********************************************************************************************************************
    '*                                                                        CUVE C01
    '*********************************************************************************************************************
    If False Then
          'With TEtatsCuves(CUVES_REGULATION.C_SAT).TNumDefauts
           With TEtatsCuves(1).TNumDefauts
        
            '--- niveau tr�s bas ---
            .NumDefautNiveauTresBas = 325
            
            '--- niveau tr�s haut ---
            .NumDefautNiveauTresHaut = 326
    
            '--- d�faut chauffage ---
            .NumDefautDefautChauffage = 327
    
            '--- d�fauts temp�ratures ---
            .NumDefautDefautPT100 = 328
            .NumDefautTemperatureTropBasse = 329
            .NumDefautTemperatureTropHaute = 330
    
            '        '--- disjonction des �lectro-vannes ouverture/fermeture des couvercles ---
            '        .NumDefautDisjonctionEVCouvercles = 331
            '
            '        '--- disjonction chauffage ---
            '        .NumDefautDisjonctionChauffage = 332
            '
            '        '--- disjonction �lectro-vanne d'arriv�e d'eau ---
            '        .NumDefautDisjonctionEVEau = 334
            '
            '        '--- d�lai trop long d'ouverture des couvercles ---
            '        .NumDefautDelaiTropLongOuvertureCouvercles = 341
            '
            '        '--- d�lai trop long de fermeture des couvercles ---
            '        .NumDefautDelaiTropLongFermetureCouvercles = 342
            '
            '        '--- d�lai trop long de l'�lectro-vanne d'arriv�e d'eau ---
            '        .NumDefautDelaiTropLongEVEau = 343
            '
        End With
    
    
    End If
  
        
    
 
   
    
    '*********************************************************************************************************************
    '*                                                                        CUVE C02
    '*********************************************************************************************************************
    
    With TEtatsCuves(CUVES_REGULATION.C_DEC).TNumDefauts

        '--- niveau tr�s bas ---
        .NumDefautNiveauTresBas = 350
        
        '--- niveau tr�s haut ---
        .NumDefautNiveauTresHaut = 351

        '--- d�faut chauffage ---
        .NumDefautDefautChauffage = 352
        
        '--- d�fauts temp�ratures ---
        .NumDefautDefautPT100 = 353
        .NumDefautTemperatureTropBasse = 354
        .NumDefautTemperatureTropHaute = 355
    
    End With
    
    '*********************************************************************************************************************
    '*                                                                        CUVE C03
    '*********************************************************************************************************************
    If False Then
        
        'With TEtatsCuves(CUVES_REGULATION.C_C03).TNumDefauts
        With TEtatsCuves(1).TNumDefauts
         
        '--- niveau tr�s bas ---
        .NumDefautNiveauTresBas = 375
        
        '--- niveau tr�s haut ---
        .NumDefautNiveauTresHaut = 376
    
        '--- d�faut chauffage ---
        .NumDefautDefautChauffage = 377
    
        '--- d�fauts temp�ratures ---
        .NumDefautDefautPT100 = 378
        .NumDefautTemperatureTropBasse = 379
        .NumDefautTemperatureTropHaute = 380
    
        End With
      
        
        '*********************************************************************************************************************
        '*                                                                        CUVE C05
        '*********************************************************************************************************************
        
        'With TEtatsCuves(CUVES_REGULATION.C_C05).TNumDefauts
         With TEtatsCuves(1).TNumDefauts
    
            '--- niveau tr�s bas ---
            .NumDefautNiveauTresBas = 400
            
            '--- niveau tr�s haut ---
            .NumDefautNiveauTresHaut = 401
    
            '--- d�faut chauffage ---
            .NumDefautDefautChauffage = 402
            
            '--- d�fauts temp�ratures ---
            .NumDefautDefautPT100 = 403
            .NumDefautTemperatureTropBasse = 404
            .NumDefautTemperatureTropHaute = 405
        
        End With
        
        '*********************************************************************************************************************
        '*                                                                        CUVE C06
        '*********************************************************************************************************************
        
        'With TEtatsCuves(CUVES_REGULATION.C_C06).TNumDefauts
     With TEtatsCuves(1).TNumDefauts
            '--- niveau tr�s bas ---
            .NumDefautNiveauTresBas = 425
            
            '--- niveau tr�s haut ---
            .NumDefautNiveauTresHaut = 426
    
            '--- d�faut chauffage ---
            .NumDefautDefautChauffage = 427
            
            '--- d�fauts temp�ratures ---
            .NumDefautDefautPT100 = 428
            .NumDefautTemperatureTropBasse = 429
            .NumDefautTemperatureTropHaute = 430
        
        End With
    
    End If
    
    '*********************************************************************************************************************
    '*                                                                        CUVE C07
    '*********************************************************************************************************************
    
    With TEtatsCuves(CUVES_REGULATION.C_C07).TNumDefauts

        '--- niveau tr�s bas ---
        .NumDefautNiveauTresBas = 450
        
        '--- niveau tr�s haut ---
        .NumDefautNiveauTresHaut = 451

        '--- d�faut chauffage ---
        .NumDefautDefautChauffage = 452
    
        '--- d�fauts temp�ratures ---
        .NumDefautDefautPT100 = 453
        .NumDefautTemperatureTropBasse = 454
        .NumDefautTemperatureTropHaute = 455
    
    End With
    
    '*********************************************************************************************************************
    '*                                                                        CUVE C13
    '*********************************************************************************************************************
    
    With TEtatsCuves(CUVES_REGULATION.C_C13).TNumDefauts

        '--- niveau tr�s bas ---
        .NumDefautNiveauTresBas = 475
        
        '--- niveau tr�s haut ---
        .NumDefautNiveauTresHaut = 476

        '--- d�faut chauffage ---
        .NumDefautDefautChauffage = 477
    
        '--- d�fauts temp�ratures ---
        .NumDefautDefautPT100 = 478
        .NumDefautTemperatureTropBasse = 479
        .NumDefautTemperatureTropHaute = 480
    
        '--- d�faut refroidissement ---
        .NumDefautDefautRefroidissement = 481
    
    End With
    
    '*********************************************************************************************************************
    '*                                                                        CUVE C14
    '*********************************************************************************************************************
    
    With TEtatsCuves(CUVES_REGULATION.C_C14).TNumDefauts

        '--- niveau tr�s bas ---
        .NumDefautNiveauTresBas = 500
        
        '--- niveau tr�s haut ---
        .NumDefautNiveauTresHaut = 501

        '--- d�faut chauffage ---
        .NumDefautDefautChauffage = 502
    
        '--- d�fauts temp�ratures ---
        .NumDefautDefautPT100 = 503
        .NumDefautTemperatureTropBasse = 504
        .NumDefautTemperatureTropHaute = 505
        
        '--- d�faut refroidissement ---
        .NumDefautDefautRefroidissement = 506
    
    End With
    
    '*********************************************************************************************************************
    '*                                                                        CUVE C15
    '*********************************************************************************************************************
    
    With TEtatsCuves(CUVES_REGULATION.C_C15).TNumDefauts

        '--- niveau tr�s bas ---
        .NumDefautNiveauTresBas = 525
        
        '--- niveau tr�s haut ---
        .NumDefautNiveauTresHaut = 526

        '--- d�faut chauffage ---
        .NumDefautDefautChauffage = 527
        
        '--- d�fauts temp�ratures ---
        .NumDefautDefautPT100 = 528
        .NumDefautTemperatureTropBasse = 529
        .NumDefautTemperatureTropHaute = 530
        
        '--- d�faut refroidissement ---
        .NumDefautDefautRefroidissement = 531
    
    End With
    
    '*********************************************************************************************************************
    '*                                                                        CUVE C16
    '*********************************************************************************************************************
    If False Then
       'With TEtatsCuves(CUVES_REGULATION.C_C16).TNumDefauts
       With TEtatsCuves(1).TNumDefauts


        '--- niveau tr�s bas ---
        .NumDefautNiveauTresBas = 550
        
        '--- niveau tr�s haut ---
        .NumDefautNiveauTresHaut = 551
    
        '--- d�faut chauffage ---
        .NumDefautDefautChauffage = 552
    
        '--- d�fauts temp�ratures ---
        .NumDefautDefautPT100 = 553
        .NumDefautTemperatureTropBasse = 554
        .NumDefautTemperatureTropHaute = 555
    
        '--- d�faut refroidissement ---
        .NumDefautDefautRefroidissement = 556
    
        End With

    
        
        
        '*********************************************************************************************************************
        '*                                                                        CUVE C19
        '*********************************************************************************************************************
        'With TEtatsCuves(CUVES_REGULATION.C_C19).TNumDefauts
        With TEtatsCuves(1).TNumDefauts
            '--- niveau tr�s bas ---
            .NumDefautNiveauTresBas = 575
            
            '--- niveau tr�s haut ---
            .NumDefautNiveauTresHaut = 576
    
            '--- d�faut chauffage ---
            .NumDefautDefautChauffage = 577
        
            '--- d�fauts temp�ratures ---
            .NumDefautDefautPT100 = 578
            .NumDefautTemperatureTropBasse = 579
            .NumDefautTemperatureTropHaute = 580
        
        End With
    End If
    '*********************************************************************************************************************
    '*                                                                        CUVE C22
    '*********************************************************************************************************************
    
    With TEtatsCuves(CUVES_REGULATION.C_C22).TNumDefauts

        '--- niveau tr�s bas ---
        .NumDefautNiveauTresBas = 600
        
        '--- niveau tr�s haut ---
        .NumDefautNiveauTresHaut = 601

        '--- d�faut chauffage ---
        .NumDefautDefautChauffage = 602
        
        '--- d�fauts temp�ratures ---
        .NumDefautDefautPT100 = 603
        .NumDefautTemperatureTropBasse = 604
        .NumDefautTemperatureTropHaute = 605
    
    End With
    
    '*********************************************************************************************************************
    '*                                                                        CUVE C27
    '*********************************************************************************************************************
    
    With TEtatsCuves(CUVES_REGULATION.C_C27).TNumDefauts

        '--- niveau tr�s bas ---
        .NumDefautNiveauTresBas = 625
        
        '--- niveau tr�s haut ---
        .NumDefautNiveauTresHaut = 626

        '--- d�faut chauffage ---
        .NumDefautDefautChauffage = 627
    
        '--- d�fauts temp�ratures ---
        .NumDefautDefautPT100 = 628
        .NumDefautTemperatureTropBasse = 629
        .NumDefautTemperatureTropHaute = 630
    
    End With
    
    '*********************************************************************************************************************
    '*                                                                        CUVE C28
    '*********************************************************************************************************************
    
    With TEtatsCuves(CUVES_REGULATION.C_C28).TNumDefauts

        '--- niveau tr�s bas ---
        .NumDefautNiveauTresBas = 650
        
        '--- niveau tr�s haut ---
        .NumDefautNiveauTresHaut = 651

        '--- d�faut chauffage ---
        .NumDefautDefautChauffage = 652
        
        '--- d�fauts temp�ratures ---
        .NumDefautDefautPT100 = 653
        .NumDefautTemperatureTropBasse = 654
        .NumDefautTemperatureTropHaute = 655
    
    End With
    
    '*********************************************************************************************************************
    '*                                                                        CUVE C31
    '*********************************************************************************************************************
    
    With TEtatsCuves(CUVES_REGULATION.C_C31).TNumDefauts

        '--- niveau tr�s bas ---
        .NumDefautNiveauTresBas = 675
        
        '--- niveau tr�s haut ---
        .NumDefautNiveauTresHaut = 676

        '--- d�faut chauffage ---
        .NumDefautDefautChauffage = 677
    
        '--- d�fauts temp�ratures ---
        .NumDefautDefautPT100 = 678
        .NumDefautTemperatureTropBasse = 679
        .NumDefautTemperatureTropHaute = 680
    
    End With
    
    '*********************************************************************************************************************
    '*                                                                        CUVE C32
    '*********************************************************************************************************************
    
    With TEtatsCuves(CUVES_REGULATION.C_C32).TNumDefauts

        '--- niveau tr�s bas ---
        .NumDefautNiveauTresBas = 700
        
        '--- niveau tr�s haut ---
        .NumDefautNiveauTresHaut = 701

        '--- d�faut chauffage ---
        .NumDefautDefautChauffage = 702
        
        '--- d�fauts temp�ratures ---
        .NumDefautDefautPT100 = 702
        .NumDefautTemperatureTropBasse = 703
        .NumDefautTemperatureTropHaute = 704
    
    End With
    
    '*********************************************************************************************************************
    '*                                                                    CUVE C33 /C34
    '*********************************************************************************************************************
    
   'With TEtatsCuves(CUVES_REGULATION.C_C33).TNumDefauts
  '
        '--- d�faut chauffage ---
   '    .NumDefautDefautChauffage = 725

    'End With
    
    '*********************************************************************************************************************
    '*                                                                  REDRESSEURS
    '*********************************************************************************************************************
    
    For a = LBound(TEtatsRedresseurs()) To UBound(TEtatsRedresseurs())
        With TEtatsRedresseurs(a).TNumDefauts
            Select Case a
                
                Case REDRESSEURS.R_C13
                    '--- redresseur 1 au poste C13 ---
                    .NumDefautDefautGeneral = 20
                    .NumDefautDelaiTropLongMarcheRedresseur = 25
                    .NumDefautIntensiteNonAtteinte = 30
                    .NumDefautIntensiteInstable = 35
                
                Case REDRESSEURS.R_C14
                    '--- redresseur 2 au poste C14 ---
                    .NumDefautDefautGeneral = 21
                    .NumDefautDelaiTropLongMarcheRedresseur = 26
                    .NumDefautIntensiteNonAtteinte = 31
                    .NumDefautIntensiteInstable = 36
                
                Case REDRESSEURS.R_C15
                    '--- redresseur 3 au poste C15 ---
                    .NumDefautDefautGeneral = 22
                    .NumDefautDelaiTropLongMarcheRedresseur = 27
                    .NumDefautIntensiteNonAtteinte = 32
                    .NumDefautIntensiteInstable = 37
                
                Case REDRESSEURS.R_C16
                    '--- redresseur 4 au poste C16 ---
                    .NumDefautDefautGeneral = 23
                    .NumDefautDelaiTropLongMarcheRedresseur = 28
                    .NumDefautIntensiteNonAtteinte = 33
                    .NumDefautIntensiteInstable = 38

                Case Else
            End Select
        End With
    Next a
    
    '*********************************************************************************************************************
    '*                                                                      ANNEXES
    '*********************************************************************************************************************

    'With TEtatsAnnexes.TNumDefauts
            
    'End With
    
    '--- affichage du type de t�che ---
    AfficheTypeTache ("")

    Exit Function

GestionErreurs:

    AttributionNumDefauts = CStr(Err.Number)

End Function


'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' R�le      : Compl�te le libell� d'un d�faut en fonction de son num�ro
'                 (cas des num�ros d'erreurs dans les variateurs)
' Entr�es :                    NumDefaut -> Repr�sente un num�ro de d�faut
' Retours :       ComplementDefaut -> Contient le texte du compl�ment ajout� au libell� du d�faut
'                 CompleteLibelleDefaut -> Libell� du d�faut compl�t�
' D�tails  :
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function CompleteLibelleDefaut(ByVal NumDefaut As Integer, _
                                                               ByRef ComplementDefaut As String) As String
                                                              
    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next

    '--- d�claration ---
    Dim LibelleDefaut As String                     'libell� d'un d�faut
    
    '--- affectation par d�faut du libell� du d�faut ---
    If NumDefaut >= DEFAUTS.NUM_MINI And NumDefaut <= DEFAUTS.NUM_MAXI Then
        LibelleDefaut = TDefauts(NumDefaut).LibelleDefaut
    End If

    '--- compl�ment du libell� si n�cessaire ---
    Select Case NumDefaut

        Case 100
            '--- translation gauche du pont ---
            'ComplementDefaut = "D�faut variateur n� " & TEntreesSortiesVariateurs6DP(VARIATEURS_6DP.V_TRL_G_PONT).NumDefaut
        
        Case 101
            '--- translation droite du pont ---
            'ComplementDefaut = "D�faut variateur n� " & TEntreesSortiesVariateurs6DP(VARIATEURS_6DP.V_TRL_D_PONT).NumDefaut
        
        Case 102
            '--- levage du pont ---
            'ComplementDefaut = "D�faut variateur n� " & TEntreesSortiesVariateurs6DP(VARIATEURS_6DP.V_LEV_PONT).NumDefaut

        Case 103
            '--- d�faut du variateur de la translation gauche du d�graissage ---
            'ComplementDefaut = "D�faut variateur n� " & TEntreesSortiesVariateurs3DP(VARIATEURS_3DP.V_TRL_G_DEGRAISSAGE).NumDefaut

        Case 104
            '--- d�faut du variateur de la translation droite du d�graissage ---
            'ComplementDefaut = "D�faut variateur n� " & TEntreesSortiesVariateurs3DP(VARIATEURS_3DP.V_TRL_D_DEGRAISSAGE).NumDefaut
        
        Case 105
            '--- d�faut du variateur rotation brosses avant ---
            'ComplementDefaut = "D�faut variateur n� " & TEntreesSortiesVariateurs3DP(VARIATEURS_3DP.V_ROT_BROSSES_AVANT).NumDefaut
        
        Case 106
            '--- d�faut du variateur rotation brosses arri�re ---
            'ComplementDefaut = "D�faut variateur n� " & TEntreesSortiesVariateurs3DP(VARIATEURS_3DP.V_ROT_BROSSES_ARRIERE).NumDefaut

        Case Else
            '--- valeur par d�faut du compl�ment du d�faut ---
            'ComplementDefaut = ""
    
    End Select

    '--- valeur de retour ---
    If ComplementDefaut = "" Then
        CompleteLibelleDefaut = LibelleDefaut
    Else
        CompleteLibelleDefaut = LibelleDefaut & " (" & ComplementDefaut & ")"
    End If

End Function

'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' R�le      : Construit une liste des num�ros de d�fauts en fonction de l'�tat d'un d�faut
' Entr�es :                                     EtatDefaut -> Etat d'un d�faut FALSE = pas de d�faut en cours
'                                                                                                      TRUE = d�faut en cours
'                                                   NumDefaut -> Repr�sente un num�ro de d�faut
' Retours :                         ListeNumDefauts -> Liste de tous les num�ros de d�fauts
'                    ListeNumDefautsPourLaLigne -> Liste des num�ros de d�fauts concernant
'                                                                          uniquement la ligne (sans les cuves)
'                  ListeNumDefautsPourUneCuve -> Liste des num�ros de d�fauts concernant
'                                                                          uniquement les cuves
' D�tails  :
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Sub ConstruitListeNumDefauts(ByVal EtatDefaut As Boolean, _
                                                              ByVal NumDefaut As Integer, _
                                                              ByRef ListeNumDefauts As String, _
                                                              Optional ByRef ListeNumDefautsPourLaLigne As String, _
                                                              Optional ByRef ListeNumDefautsPourUneCuve As String)
    
    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next

    '--- d�claration ---
    
    If NumDefaut > 0 Then
    
        If EtatDefaut = True And TDefauts(NumDefaut).SignalerOuiNon = True Then
            
            '--- affectation de tous les num�ros de d�faut ---
            ListeNumDefauts = ListeNumDefauts & CStr(NumDefaut) & SEPARATEUR_NUM_DEFAUTS
            
            '--- affectation des num�ros de d�fauts uniquement pour la ligne ---
            If IsMissing(ListeNumDefautsPourLaLigne) = False Then
                ListeNumDefautsPourLaLigne = ListeNumDefautsPourLaLigne & CStr(NumDefaut) & SEPARATEUR_NUM_DEFAUTS
            End If
            
            '--- affectation des num�ros de d�fauts uniquement pour une cuve ---
            If IsMissing(ListeNumDefautsPourUneCuve) = False Then
                ListeNumDefautsPourUneCuve = ListeNumDefautsPourUneCuve & CStr(NumDefaut) & SEPARATEUR_NUM_DEFAUTS
            End If
    
        End If

    End If

End Sub

'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' R�le      : Enregistre la date de d�tection et de disparition d'un d�faut en fonction de son �tat
' Entr�es :   EtatDefaut -> Etat d'un d�faut FALSE = pas de d�faut en cours
'                                                                    TRUE = d�faut en cours
'                 NumDefaut -> Repr�sente un num�ro de d�faut
' Retours :
' D�tails  :
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Sub EnregistreDateDetectionDisparitionDefaut(ByVal EtatDefaut As Boolean, _
                                                                                       ByVal NumDefaut As Integer)
    
    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next

    '--- d�claration ---
        
    '--- enregistrement de la date de d�tection du d�faut et la date de sa disparition ---
    If EtatDefaut = True Then
        SignalerDefaut NumDefaut, True
    Else
        SignalerDefaut NumDefaut, False
        InitialiserAntiRebondsDefaut NumDefaut
    End If
        
End Sub

'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' R�le      : Ajoute des num�ros de d�fauts � une liste sans que le resultat contienne des doublons
' Entr�es :            ListeNumDefautsDeBase -> Liste des num�ros de d�fauts de base
'                           ListeNumDefautsAAjouter -> Liste des num�ros de d�fauts � ajouter � la liste de base
' Retours : AjoutNumDefautsSansDoublons -> Contient la l'addition des 2 listes sans doublons
' D�tails  :
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function AjoutNumDefautsSansDoublons(ByVal ListeNumDefautsDeBase As String, _
                                                                               ByVal ListeNumDefautsAAjouter As String) As String
    
    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next
    
    '--- d�claration ---
    Dim a As Integer
    Dim NumDefaut As String, _
            CopieListeNumDefautsDeBase  As String, _
            CopieListeNumDefautsAAjouter As String
    Dim TListeNumDefautsAAjouter As Variant
    Dim TListeNumDefautsASupprimer As Variant
           
    If ListeNumDefautsAAjouter <> "" Then
    
        If ListeNumDefautsDeBase = "" Then
    
            '--- si la liste des n� de d�fauts de base est vide alors affect� directement la liste des ajouts ---
            ListeNumDefautsDeBase = ListeNumDefautsAAjouter
        
        Else
    
            '--- avant toutes op�rations il faut effectuer un contr�le sur les chaines afin d'�viter que le
            '    dernier caract�re soit un s�parateur de n� de d�fauts ---
            If Right(ListeNumDefautsDeBase, 1) = SEPARATEUR_NUM_DEFAUTS Then
                ListeNumDefautsDeBase = Left(ListeNumDefautsDeBase, Pred(Len(ListeNumDefautsDeBase))) 'suppression du dernier s�parateur
            End If
            If Right(ListeNumDefautsAAjouter, 1) = SEPARATEUR_NUM_DEFAUTS Then
                ListeNumDefautsAAjouter = Left(ListeNumDefautsAAjouter, Pred(Len(ListeNumDefautsAAjouter))) 'suppression du dernier s�parateur
            End If
            
            '--- construction du tableau des num�ros de d�fauts � ajouter ---
            TListeNumDefautsAAjouter = Split(ListeNumDefautsAAjouter, SEPARATEUR_NUM_DEFAUTS)
        
            '--- recherche des donn�es dans le tableau ---
            For a = LBound(TListeNumDefautsAAjouter) To UBound(TListeNumDefautsAAjouter)
        
                '--- affectation sur la liste de base afin de pouvoir d�tecter
                    'un num�ro de d�faut avec 2 s�parateurs (ex : -123-,-41-,-523-) pour �viter une mauvaise comparaison ---
                CopieListeNumDefautsDeBase = SEPARATEUR_NUM_DEFAUTS & ListeNumDefautsDeBase & SEPARATEUR_NUM_DEFAUTS
        
                '--- affectation ---
                NumDefaut = TListeNumDefautsAAjouter(a)
        
                '--- recherche du d�faut ---
                If InStr(CopieListeNumDefautsDeBase, SEPARATEUR_NUM_DEFAUTS & NumDefaut & SEPARATEUR_NUM_DEFAUTS) = 0 Then
        
                    '--- le d�faut n'est pas dans la liste alors le rajouter ---
                    ListeNumDefautsDeBase = ListeNumDefautsDeBase & SEPARATEUR_NUM_DEFAUTS & NumDefaut
        
                End If
        
            Next a
            
        End If
    
    End If
        
    '--- affectation ---
    AjoutNumDefautsSansDoublons = ListeNumDefautsDeBase

End Function

'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' R�le      : D�code une suite de n� d'alarmes de la ligne en une suite de libell�s correspondant
' Entr�es :              AlarmesLigne -> Suite de n� d'alarmes avec s�parateur
' Retours : DecodeAlarmesLigne -> Suite des libell�s des alarmes avec divers renseignements
' D�tails  :
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function DecodeAlarmesLigne(ByVal AlarmesLigne As String) As String
    
    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next
    
    '--- d�claration ---
    Dim a As Integer
    Dim NumDefaut As Integer
    Dim TAlarmesLigne As Variant
           
    '--- affectation ---
    DecodeAlarmesLigne = ""
    
    If AlarmesLigne <> "" Then
                    
        '--- construction du tableau contenant les num�ros d'alarmes ---
        TAlarmesLigne = Split(AlarmesLigne, SEPARATEUR_NUM_DEFAUTS)
                        
        '--- construction de la chaine des libell�s ---
        For a = LBound(TAlarmesLigne) To UBound(TAlarmesLigne)
            If IsNumeric(TAlarmesLigne(a)) = True Then
                
                '--- affectation du n� du d�faut ---
                NumDefaut = TAlarmesLigne(a)
                
                If NumDefaut >= DEFAUTS.NUM_MINI And NumDefaut <= DEFAUTS.NUM_MAXI Then
                    DecodeAlarmesLigne = DecodeAlarmesLigne & _
                                                          "D�faut n� " & TAlarmesLigne(a) & " - " & _
                                                          TDefauts(TAlarmesLigne(a)).LibelleDefaut & vbCrLf
                End If
            
            End If
        Next a
    
    End If
                                                                      
End Function

'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' R�le      : D�code une suite de n� d'alarmes d'un poste en une suite de libell�s correspondant
' Entr�es :             AlarmesPoste -> Suite de n� d'alarmes avec s�parateur
' Retours : DecodeAlarmesPoste -> Suite des libell�s des alarmes avec divers renseignements
' D�tails  :
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function DecodeAlarmesPoste(ByVal AlarmesPoste As String) As String
    
    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next
    
    '--- d�claration ---
    Dim a As Integer
    Dim NumDefaut As Integer
    Dim TAlarmesPoste As Variant
           
    '--- affectation ---
    DecodeAlarmesPoste = ""
    
    If AlarmesPoste <> "" Then
                    
        '--- construction du tableau contenant les num�ros d'alarmes ---
        TAlarmesPoste = Split(AlarmesPoste, SEPARATEUR_NUM_DEFAUTS)
                        
        '--- construction de la chaine des libell�s ---
        For a = LBound(TAlarmesPoste) To UBound(TAlarmesPoste)
            If IsNumeric(TAlarmesPoste(a)) = True Then
                
                '--- affectation du n� du d�faut ---
                NumDefaut = TAlarmesPoste(a)
                
                If NumDefaut >= DEFAUTS.NUM_MINI And NumDefaut <= DEFAUTS.NUM_MAXI Then
                    DecodeAlarmesPoste = DecodeAlarmesPoste & _
                                                          "D�faut n� " & TAlarmesPoste(a) & " - " & _
                                                          TDefauts(TAlarmesPoste(a)).LibelleDefaut & vbCrLf
                End If
            
            End If
        Next a
    
    End If
                                                                      
End Function

'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' R�le      : Initialise les bits d'anti-rebonds
' Entr�es : NumDefaut -> n� du d�faut faisant l'objet de la demande
' Retours :
' D�tails  :
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Sub InitialiserAntiRebondsDefaut(ByVal NumDefaut As Long)
    
    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next
    
    '--- d�claration ---
           
    If NumDefaut >= DEFAUTS.NUM_MINI And NumDefaut <= DEFAUTS.NUM_MAXI Then
    
         With TDefauts(NumDefaut)
            
            '--- initialisation des anti-rebonds ---
            .AntiRebondGyrophare = False
            .AntiRebondKlaxon = False
            .AntiRebondTra�abiliteAlarmes = False
        
        End With
    
    End If

End Sub

'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' R�le      : Signale un d�faut sur le gyrophare et le klaxon en fonction des valeurs de la table des d�fauts
' Entr�es : NumDefaut -> n� du d�faut faisant l'objet de la demande
' Retours :
' D�tails  :
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Sub SignalerDefaut(ByVal NumDefaut As Long, EtatsDefaut As Boolean)
    
    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next
    
    '--- d�claration ---
           
    If NumDefaut >= DEFAUTS.NUM_MINI And NumDefaut <= DEFAUTS.NUM_MAXI Then
    
         With TDefauts(NumDefaut)
            
            '--- signalisation dans le fichier de tra�abilit� des alarmes ---
            If .AntiRebondTra�abiliteAlarmes = False And EtatsDefaut = True Then
                Bidon = EnregistrementDefautDansTra�abiliteAlarmes(NumDefaut, EtatsDefaut)
                .AntiRebondTra�abiliteAlarmes = True
            End If
            
            '--- signalisation dans le fichier de tra�abilit� des alarmes ---
            If .AntiRebondTra�abiliteAlarmes = True And EtatsDefaut = False Then
                Bidon = EnregistrementDefautDansTra�abiliteAlarmes(NumDefaut, EtatsDefaut)
                .AntiRebondTra�abiliteAlarmes = False
            End If
            
            If EtatsDefaut = True Then
            
                '--- signalisation au niveau du gyrophare ---
                If .GyrophareOuiNon = True And .AntiRebondGyrophare = False Then
                    Bidon = APIEcritureVariableNommee("DEFAUTS", "M_Dem_PC_Gyrophare", True)
                    .AntiRebondGyrophare = True
                End If
                
                '--- signalisation au niveau du klaxon ---
                If .KlaxonOuiNon = True And .AntiRebondKlaxon = False Then
                    Bidon = APIEcritureVariableNommee("DEFAUTS", "M_Dem_PC_Klaxon", True)
                    .AntiRebondKlaxon = True
                End If
            
            End If
            
        End With
            
    End If
           
End Sub

'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' R�le      :  Analyse de la totalit� des alarmes et visualisation sur la ligne en bas � gauche de l'�cran principal
'                 des alarmes en cours
' D�tails  :
' Retours :
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Sub VisualisationLigneAlarmes()
    
    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next

    '--- constantes priv�es ---
    
    '--- d�claration ---
    Dim UnDefautAuMoinsSignale As Boolean       'indique un d�faut au moins de signaler par section (pont, cuves, etc...)
    Dim EtatDefaut As Boolean                               'repr�sente l'�tat d'un d�faut (FALSE = pas de d�faut, TRUE = d�faut en cours)
    
    Dim a As Integer                                                'pour les boucles FOR...NEXT
    Dim NumDefaut As Integer, _
            NumDefautAAfficher As Integer
    Static CptAppels As Integer                                'compteur d'appels de la routine
    Static CptDefauts As Integer                              'compteur de d�fauts
    Dim NumPoste1 As Integer                                'repr�sente le poste 1 d'une cuve � postes multiples
    Dim NumPoste2 As Integer                                'repr�sente le poste 2 d'une cuve � postes multiples
    Dim NumChargeACePoste As Integer                'repr�sente le num�ro de charge dans un poste
    
    Dim NumCuve As Long                                      'repr�sente un num�ro de cuve quelconque
    
    Dim ComplementDefaut As String                     'contient le texte du compl�ment ajout� au libell� du d�faut
    Dim LibelleCompleteDefaut As String               'repr�sente un libell� compl�t� d'un d�faut (pour les num�ros de d�faut des variateurs, etc ...)
    Dim LibelleDefautAfficheur As String                 'libell� du d�faut destin� � l'afficheur
    
    Dim ListeNumDefauts As String, _
            ListeNumDefautsPourUneCuve As String, _
            ListeNumDefautsPourLaLigne As String
    Dim TNumDefauts As Variant, _
            TNumDefautsPourUneCuve As Variant
   
    '*********************************************************************************************************************
    '*                                                                   ETATS DE LA LIGNE
    '*********************************************************************************************************************

    With TEtatsLigne

        '--- arr�t g�n�rale ---
        NumDefaut = .TNumDefauts.NumDefautArretGeneral
        EtatDefaut = .ArretGeneral
        ConstruitListeNumDefauts EtatDefaut, NumDefaut, ListeNumDefauts, ListeNumDefautsPourLaLigne   'contruction des listes de d�fauts
        EnregistreDateDetectionDisparitionDefaut EtatDefaut, NumDefaut 'enregistrement de la date de d�tection et disparition du d�faut

        '-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

        '--- arr�t d'urgence g�n�ral ---
        NumDefaut = .TNumDefauts.NumDefautArretUrgence
        EtatDefaut = .ArretUrgence
        ConstruitListeNumDefauts EtatDefaut, NumDefaut, ListeNumDefauts, ListeNumDefautsPourLaLigne    'contruction des listes de d�fauts
        EnregistreDateDetectionDisparitionDefaut EtatDefaut, NumDefaut 'enregistrement de la date de d�tection et disparition du d�faut

        '-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
        
        '--- portillons et ligne de vie ---
        NumDefaut = .TNumDefauts.NumDefautPortillonsLigneVie
        EtatDefaut = .PortillonsLigneVie
        ConstruitListeNumDefauts EtatDefaut, NumDefaut, ListeNumDefauts, ListeNumDefautsPourLaLigne    'contruction des listes de d�fauts
        EnregistreDateDetectionDisparitionDefaut EtatDefaut, NumDefaut 'enregistrement de la date de d�tection et disparition du d�faut

        '-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

        '--- chaine de s�curit� du pont 1 ---
        NumDefaut = .TNumDefauts.NumDefautSecuriteP1
        EtatDefaut = .SecuriteP1
        ConstruitListeNumDefauts EtatDefaut, NumDefaut, ListeNumDefauts, ListeNumDefautsPourLaLigne    'contruction des listes de d�fauts
        EnregistreDateDetectionDisparitionDefaut EtatDefaut, NumDefaut 'enregistrement de la date de d�tection et disparition du d�faut

        '-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

        '--- chaine de s�curit� du pont 2 ---
        NumDefaut = .TNumDefauts.NumDefautSecuriteP2
        EtatDefaut = .SecuriteP2
        ConstruitListeNumDefauts EtatDefaut, NumDefaut, ListeNumDefauts, ListeNumDefautsPourLaLigne    'contruction des listes de d�fauts
        EnregistreDateDetectionDisparitionDefaut EtatDefaut, NumDefaut 'enregistrement de la date de d�tection et disparition du d�faut

        '-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

        '--- manque de tension ---
        NumDefaut = .TNumDefauts.NumDefautManqueTension
        EtatDefaut = .ManqueTension
        ConstruitListeNumDefauts EtatDefaut, NumDefaut, ListeNumDefauts, ListeNumDefautsPourLaLigne    'contruction des listes de d�fauts
        EnregistreDateDetectionDisparitionDefaut EtatDefaut, NumDefaut     'enregistrement de la date de d�tection et disparition du d�faut

        '-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

        '--- manque d'air ---
        NumDefaut = .TNumDefauts.NumDefautManqueAir
        EtatDefaut = .ManqueAir
        ConstruitListeNumDefauts EtatDefaut, NumDefaut, ListeNumDefauts, ListeNumDefautsPourLaLigne    'contruction des listes de d�fauts
        EnregistreDateDetectionDisparitionDefaut EtatDefaut, NumDefaut     'enregistrement de la date de d�tection et disparition du d�faut

        '-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

        '--- arr�t par stop ligne ---
        NumDefaut = .TNumDefauts.NumDefautStopLigne
        EtatDefaut = .StopLigne
        ConstruitListeNumDefauts EtatDefaut, NumDefaut, ListeNumDefauts, ListeNumDefautsPourLaLigne    'contruction des listes de d�fauts
        EnregistreDateDetectionDisparitionDefaut EtatDefaut, NumDefaut     'enregistrement de la date de d�tection et disparition du d�faut

        '-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

'        '--- anti collision sur les d�fauts des ponts ---
'        NumDefaut = .TNumDefauts.NumDefautAntiCollisionDefautsPonts
'        EtatDefaut = .AntiCollisionDefautsPonts
'        ConstruitListeNumDefauts EtatDefaut, NumDefaut, ListeNumDefauts, ListeNumDefautsPourLaLigne    'contruction des listes de d�fauts
'        EnregistreDateDetectionDisparitionDefaut EtatDefaut, NumDefaut     'enregistrement de la date de d�tection et disparition du d�faut
'
'        '-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'
'        '--- anti collision sur la somme des lasers des 2 ponts ---
'        NumDefaut = .TNumDefauts.NumDefautAntiCollisionLasersPonts
'        EtatDefaut = .AntiCollisionLasersPonts
'        ConstruitListeNumDefauts EtatDefaut, NumDefaut, ListeNumDefauts, ListeNumDefautsPourLaLigne    'contruction des listes de d�fauts
'        EnregistreDateDetectionDisparitionDefaut EtatDefaut, NumDefaut     'enregistrement de la date de d�tection et disparition du d�faut
'
'        '-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'
'        '--- poste occup� ---
'        NumDefaut = .TNumDefauts.NumDefautPosteOccupe
'        EtatDefaut = .PosteOccupe
'        ConstruitListeNumDefauts EtatDefaut, NumDefaut, ListeNumDefauts, ListeNumDefautsPourLaLigne    'contruction des listes de d�fauts
'        EnregistreDateDetectionDisparitionDefaut EtatDefaut, NumDefaut     'enregistrement de la date de d�tection et disparition du d�faut

    End With
        
    '*********************************************************************************************************************
    '*                                                                   LES PONTS
    '*********************************************************************************************************************
    
    For a = LBound(TEtatsPonts()) To UBound(TEtatsPonts())

        With TEtatsPonts(a)

            '--- initialisation � FALSE par d�faut ---
            UnDefautAuMoinsSignale = False

            '-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

            '--- Poste OQP cellule du PONT ---
            NumDefaut = .TNumDefauts.NumDefautPresencePiece
            EtatDefaut = .TEntreesAPI.M_DefautPresencePicece
            UnDefautAuMoinsSignale = UnDefautAuMoinsSignale Or EtatDefaut
            ConstruitListeNumDefauts EtatDefaut, NumDefaut, ListeNumDefauts, ListeNumDefautsPourLaLigne    'contruction des listes de d�fauts
            EnregistreDateDetectionDisparitionDefaut EtatDefaut, NumDefaut     'enregistrement de la date de d�tection et disparition du d�faut

            '-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

            '--- d�faut variateur de la translation du pont ---
            NumDefaut = .TNumDefauts.NumDefautDefautVariateurTrLPont
            EtatDefaut = .TEntreesAPI.M_DefautVariateurTrlPont
            UnDefautAuMoinsSignale = UnDefautAuMoinsSignale Or EtatDefaut
            ConstruitListeNumDefauts EtatDefaut, NumDefaut, ListeNumDefauts, ListeNumDefautsPourLaLigne    'contruction des listes de d�fauts
            EnregistreDateDetectionDisparitionDefaut EtatDefaut, NumDefaut     'enregistrement de la date de d�tection et disparition du d�faut

            '-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

            '--- axe non r�f�renc� de la translation du pont ---
            NumDefaut = .TNumDefauts.NumDefautAxeNonReferenceTrlPont
            EtatDefaut = .TEntreesAPI.M_AxeNonReferenceTrlPont
            UnDefautAuMoinsSignale = UnDefautAuMoinsSignale Or EtatDefaut
            ConstruitListeNumDefauts EtatDefaut, NumDefaut, ListeNumDefauts, ListeNumDefautsPourLaLigne    'contruction des listes de d�fauts
            EnregistreDateDetectionDisparitionDefaut EtatDefaut, NumDefaut     'enregistrement de la date de d�tection et disparition du d�faut

            '-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

            '--- d�faut variateur du levage du pont ---
            NumDefaut = .TNumDefauts.NumDefautDefautVariateurLevPont
            EtatDefaut = .TEntreesAPI.M_DefautVariateurLevPont
            UnDefautAuMoinsSignale = UnDefautAuMoinsSignale Or EtatDefaut
            ConstruitListeNumDefauts EtatDefaut, NumDefaut, ListeNumDefauts, ListeNumDefautsPourLaLigne    'contruction des listes de d�fauts
            EnregistreDateDetectionDisparitionDefaut EtatDefaut, NumDefaut     'enregistrement de la date de d�tection et disparition du d�faut
            
            '-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

            '--- surcourse haut du levage ---
            NumDefaut = .TNumDefauts.NumDefautSurcourseLevHaut
            EtatDefaut = .TEntreesAPI.M_SurcourseLevHaut
            UnDefautAuMoinsSignale = UnDefautAuMoinsSignale Or EtatDefaut
            ConstruitListeNumDefauts EtatDefaut, NumDefaut, ListeNumDefauts, ListeNumDefautsPourLaLigne    'contruction des listes de d�fauts
            EnregistreDateDetectionDisparitionDefaut EtatDefaut, NumDefaut     'enregistrement de la date de d�tection et disparition du d�faut
            
            '-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

            '--- surcourse bas du levage ---
            NumDefaut = .TNumDefauts.NumDefautSurcourseLevBas
            EtatDefaut = .TEntreesAPI.M_SurcourseLevBas
            UnDefautAuMoinsSignale = UnDefautAuMoinsSignale Or EtatDefaut
            ConstruitListeNumDefauts EtatDefaut, NumDefaut, ListeNumDefauts, ListeNumDefautsPourLaLigne    'contruction des listes de d�fauts
            EnregistreDateDetectionDisparitionDefaut EtatDefaut, NumDefaut     'enregistrement de la date de d�tection et disparition du d�faut

            '-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
            
            '--- axe non r�f�renc� du levage du pont ---
            NumDefaut = .TNumDefauts.NumDefautAxeNonReferenceLevPont
            EtatDefaut = .TEntreesAPI.M_AxeNonReferenceLevPont
            UnDefautAuMoinsSignale = UnDefautAuMoinsSignale Or EtatDefaut
            ConstruitListeNumDefauts EtatDefaut, NumDefaut, ListeNumDefauts, ListeNumDefautsPourLaLigne    'contruction des listes de d�fauts
            EnregistreDateDetectionDisparitionDefaut EtatDefaut, NumDefaut     'enregistrement de la date de d�tection et disparition du d�faut

            '-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
            
            '--- d�lai trop long de descente des accroches ---
            NumDefaut = .TNumDefauts.NumDefautDelaiTropLongDescenteAccroches
            EtatDefaut = .TEntreesAPI.M_DelaiTropLongDescenteAccroches
            UnDefautAuMoinsSignale = UnDefautAuMoinsSignale Or EtatDefaut
            ConstruitListeNumDefauts EtatDefaut, NumDefaut, ListeNumDefauts, ListeNumDefautsPourLaLigne    'contruction des listes de d�fauts
            EnregistreDateDetectionDisparitionDefaut EtatDefaut, NumDefaut     'enregistrement de la date de d�tection et disparition du d�faut

            '-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
            
            '--- d�lai trop long de mont�e des accroches ---
            NumDefaut = .TNumDefauts.NumDefautDelaiTropLongMonteeAccroches
            EtatDefaut = .TEntreesAPI.M_DelaiTropLongMonteeAccroches
            UnDefautAuMoinsSignale = UnDefautAuMoinsSignale Or EtatDefaut
            ConstruitListeNumDefauts EtatDefaut, NumDefaut, ListeNumDefauts, ListeNumDefautsPourLaLigne    'contruction des listes de d�fauts
            EnregistreDateDetectionDisparitionDefaut EtatDefaut, NumDefaut     'enregistrement de la date de d�tection et disparition du d�faut

            '-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

            '--- affectation d�finitive ---
            .UnDefautAuMoinsSignale = UnDefautAuMoinsSignale

        End With

    Next a
    
    '*********************************************************************************************************************
    '*                                                            CUVE DE DEGRAISSAGE - C00
    '*********************************************************************************************************************
    
    '--- affectation ---
    NumCuve = CUVES_REGULATION.C_C00
    ListeNumDefautsPourUneCuve = ""
    
    '--- recherche de la pr�sence d'une charge � ce poste ---
    NumChargeACePoste = TEtatsPostes(CorrespondanceCuvesAPIPostes(NumCuve)).NumCharge

    '--- recherche de la pr�sence d'une charge � ce poste ---
    With TEtatsCuves(NumCuve)

        '--- initialisation � FALSE par d�faut ---
        UnDefautAuMoinsSignale = False

        If TEtatsLigne.MarcheGenerale = True Then

            '-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

            '--- niveau trop bas ---
            NumDefaut = .TNumDefauts.NumDefautNiveauTresBas
            EtatDefaut = .TEntreesAPI.E_NiveauTresBas
            UnDefautAuMoinsSignale = UnDefautAuMoinsSignale Or EtatDefaut
            ConstruitListeNumDefauts EtatDefaut, NumDefaut, ListeNumDefauts, , ListeNumDefautsPourUneCuve  'contruction des listes de d�fauts
            EnregistreDateDetectionDisparitionDefaut EtatDefaut, NumDefaut 'enregistrement de la date de d�tection et disparition du d�faut

            '--- niveau trop haut ---
            NumDefaut = .TNumDefauts.NumDefautNiveauTresHaut
            EtatDefaut = .TEntreesAPI.E_NiveauTresHaut
            UnDefautAuMoinsSignale = UnDefautAuMoinsSignale Or EtatDefaut
            ConstruitListeNumDefauts EtatDefaut, NumDefaut, ListeNumDefauts, , ListeNumDefautsPourUneCuve  'contruction des listes de d�fauts
            EnregistreDateDetectionDisparitionDefaut EtatDefaut, NumDefaut 'enregistrement de la date de d�tection et disparition du d�faut

            '-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

            '--- d�faut du chauffage ---
            NumDefaut = .TNumDefauts.NumDefautDefautChauffage
            EtatDefaut = .TEntreesAPI.E_DefautChauffage
            UnDefautAuMoinsSignale = UnDefautAuMoinsSignale Or EtatDefaut
            ConstruitListeNumDefauts EtatDefaut, NumDefaut, ListeNumDefauts, , ListeNumDefautsPourUneCuve  'contruction des listes de d�fauts
            EnregistreDateDetectionDisparitionDefaut EtatDefaut, NumDefaut 'enregistrement de la date de d�tection et disparition du d�faut
            
            '-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
            
            '--- d�faut PT100 ---
            NumDefaut = .TNumDefauts.NumDefautDefautPT100
            EtatDefaut = .TEntreesAPI.DefautPT100
            UnDefautAuMoinsSignale = UnDefautAuMoinsSignale Or EtatDefaut
            ConstruitListeNumDefauts EtatDefaut, NumDefaut, ListeNumDefauts, , ListeNumDefautsPourUneCuve  'contruction des listes de d�fauts
            EnregistreDateDetectionDisparitionDefaut EtatDefaut, NumDefaut 'enregistrement de la date de d�tection et disparition du d�faut
            
            '--- temp�rature trop basse ---
            NumDefaut = .TNumDefauts.NumDefautTemperatureTropBasse
            EtatDefaut = .TEntreesAPI.TemperatureTropBasse
            UnDefautAuMoinsSignale = UnDefautAuMoinsSignale Or EtatDefaut
            ConstruitListeNumDefauts EtatDefaut, NumDefaut, ListeNumDefauts, , ListeNumDefautsPourUneCuve  'contruction des listes de d�fauts
            EnregistreDateDetectionDisparitionDefaut EtatDefaut, NumDefaut 'enregistrement de la date de d�tection et disparition du d�faut
            
            '--- temp�rature trop haute ---
            NumDefaut = .TNumDefauts.NumDefautTemperatureTropHaute
            EtatDefaut = .TEntreesAPI.TemperatureTropHaute
            UnDefautAuMoinsSignale = UnDefautAuMoinsSignale Or EtatDefaut
            ConstruitListeNumDefauts EtatDefaut, NumDefaut, ListeNumDefauts, , ListeNumDefautsPourUneCuve  'contruction des listes de d�fauts
            EnregistreDateDetectionDisparitionDefaut EtatDefaut, NumDefaut 'enregistrement de la date de d�tection et disparition du d�faut

            '-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

        End If

        '--- construction de la liste des num�ros de d�fauts pour une cuve ayant une charge ---
        If NumChargeACePoste = 0 Then
            .ListeNumDefautsSiCharge = ""    'vider la liste si pas de charge dans le poste
        Else
            If ListeNumDefautsPourUneCuve <> "" Then  'remplir la liste en �vitant les doublons
                                                                                     'suppression du dernier s�parateur inutile car dans fonction
                .ListeNumDefautsSiCharge = AjoutNumDefautsSansDoublons(.ListeNumDefautsSiCharge, _
                                                                                                                     ListeNumDefautsPourUneCuve) 'remplir la liste en �liminant les doublons
            End If
        End If

        '--- affectation d�finitive ---
        .UnDefautAuMoinsSignale = UnDefautAuMoinsSignale

    End With
    
    '*********************************************************************************************************************
    '*                                                                CUVE C01
    '*********************************************************************************************************************
    If False Then
       '--- affectation ---
        'NumCuve = CUVES_REGULATION.C_SAT
        ListeNumDefautsPourUneCuve = ""
        
        '--- recherche de la pr�sence d'une charge � ce poste ---
        NumChargeACePoste = TEtatsPostes(CorrespondanceCuvesAPIPostes(NumCuve)).NumCharge
    
        With TEtatsCuves(NumCuve)
    
            '--- initialisation � FALSE par d�faut ---
            UnDefautAuMoinsSignale = False
    
            If TEtatsLigne.MarcheGenerale = True Then
    
                '-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    
                '--- niveau trop bas ---
                NumDefaut = .TNumDefauts.NumDefautNiveauTresBas
                EtatDefaut = .TEntreesAPI.E_NiveauTresBas
                UnDefautAuMoinsSignale = UnDefautAuMoinsSignale Or EtatDefaut
                ConstruitListeNumDefauts EtatDefaut, NumDefaut, ListeNumDefauts, , ListeNumDefautsPourUneCuve  'contruction des listes de d�fauts
                EnregistreDateDetectionDisparitionDefaut EtatDefaut, NumDefaut 'enregistrement de la date de d�tection et disparition du d�faut
    
                '--- niveau trop haut ---
                NumDefaut = .TNumDefauts.NumDefautNiveauTresHaut
                EtatDefaut = .TEntreesAPI.E_NiveauTresHaut
                UnDefautAuMoinsSignale = UnDefautAuMoinsSignale Or EtatDefaut
                ConstruitListeNumDefauts EtatDefaut, NumDefaut, ListeNumDefauts, , ListeNumDefautsPourUneCuve  'contruction des listes de d�fauts
                EnregistreDateDetectionDisparitionDefaut EtatDefaut, NumDefaut 'enregistrement de la date de d�tection et disparition du d�faut
    
                '-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    
                '--- d�faut du chauffage ---
                NumDefaut = .TNumDefauts.NumDefautDefautChauffage
                EtatDefaut = .TEntreesAPI.E_DefautChauffage
                UnDefautAuMoinsSignale = UnDefautAuMoinsSignale Or EtatDefaut
                ConstruitListeNumDefauts EtatDefaut, NumDefaut, ListeNumDefauts, , ListeNumDefautsPourUneCuve  'contruction des listes de d�fauts
                EnregistreDateDetectionDisparitionDefaut EtatDefaut, NumDefaut 'enregistrement de la date de d�tection et disparition du d�faut
                
                '-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
                
                '--- d�faut PT100 ---
                NumDefaut = .TNumDefauts.NumDefautDefautPT100
                EtatDefaut = .TEntreesAPI.DefautPT100
                UnDefautAuMoinsSignale = UnDefautAuMoinsSignale Or EtatDefaut
                ConstruitListeNumDefauts EtatDefaut, NumDefaut, ListeNumDefauts, , ListeNumDefautsPourUneCuve  'contruction des listes de d�fauts
                EnregistreDateDetectionDisparitionDefaut EtatDefaut, NumDefaut 'enregistrement de la date de d�tection et disparition du d�faut
                
                '--- temp�rature trop basse ---
                NumDefaut = .TNumDefauts.NumDefautTemperatureTropBasse
                EtatDefaut = .TEntreesAPI.TemperatureTropBasse
                UnDefautAuMoinsSignale = UnDefautAuMoinsSignale Or EtatDefaut
                ConstruitListeNumDefauts EtatDefaut, NumDefaut, ListeNumDefauts, , ListeNumDefautsPourUneCuve  'contruction des listes de d�fauts
                EnregistreDateDetectionDisparitionDefaut EtatDefaut, NumDefaut 'enregistrement de la date de d�tection et disparition du d�faut
                
                '--- temp�rature trop haute ---
                NumDefaut = .TNumDefauts.NumDefautTemperatureTropHaute
                EtatDefaut = .TEntreesAPI.TemperatureTropHaute
                UnDefautAuMoinsSignale = UnDefautAuMoinsSignale Or EtatDefaut
                ConstruitListeNumDefauts EtatDefaut, NumDefaut, ListeNumDefauts, , ListeNumDefautsPourUneCuve  'contruction des listes de d�fauts
                EnregistreDateDetectionDisparitionDefaut EtatDefaut, NumDefaut 'enregistrement de la date de d�tection et disparition du d�faut
    
                '-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    
            End If
    
            '--- construction de la liste des num�ros de d�fauts pour une cuve ayant une charge ---
            If NumChargeACePoste = 0 Then
                .ListeNumDefautsSiCharge = ""    'vider la liste si pas de charge dans le poste
            Else
                If ListeNumDefautsPourUneCuve <> "" Then  'remplir la liste en �vitant les doublons
                                                                                         'suppression du dernier s�parateur inutile car dans fonction
                    .ListeNumDefautsSiCharge = AjoutNumDefautsSansDoublons(.ListeNumDefautsSiCharge, _
                                                                                                                         ListeNumDefautsPourUneCuve) 'remplir la liste en �liminant les doublons
                End If
            End If
    
            '--- affectation d�finitive ---
            .UnDefautAuMoinsSignale = UnDefautAuMoinsSignale
    
        End With
    
    End If
     
        
    
    '*********************************************************************************************************************
    '*                                                                CUVE C02
    '*********************************************************************************************************************
    
    '--- affectation ---
    NumCuve = CUVES_REGULATION.C_DEC
    ListeNumDefautsPourUneCuve = ""
    
    '--- recherche de la pr�sence d'une charge � ce poste ---
    NumChargeACePoste = TEtatsPostes(CorrespondanceCuvesAPIPostes(NumCuve)).NumCharge

    With TEtatsCuves(NumCuve)

        '--- initialisation � FALSE par d�faut ---
        UnDefautAuMoinsSignale = False

        If TEtatsLigne.MarcheGenerale = True Then

            '-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
            
            '--- niveau trop bas ---
            NumDefaut = .TNumDefauts.NumDefautNiveauTresBas
            EtatDefaut = .TEntreesAPI.E_NiveauTresBas
            UnDefautAuMoinsSignale = UnDefautAuMoinsSignale Or EtatDefaut
            ConstruitListeNumDefauts EtatDefaut, NumDefaut, ListeNumDefauts, , ListeNumDefautsPourUneCuve  'contruction des listes de d�fauts
            EnregistreDateDetectionDisparitionDefaut EtatDefaut, NumDefaut 'enregistrement de la date de d�tection et disparition du d�faut

            '--- niveau trop haut ---
            NumDefaut = .TNumDefauts.NumDefautNiveauTresHaut
            EtatDefaut = .TEntreesAPI.E_NiveauTresHaut
            UnDefautAuMoinsSignale = UnDefautAuMoinsSignale Or EtatDefaut
            ConstruitListeNumDefauts EtatDefaut, NumDefaut, ListeNumDefauts, , ListeNumDefautsPourUneCuve  'contruction des listes de d�fauts
            EnregistreDateDetectionDisparitionDefaut EtatDefaut, NumDefaut 'enregistrement de la date de d�tection et disparition du d�faut

            '-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

            '--- d�faut du chauffage ---
            NumDefaut = .TNumDefauts.NumDefautDefautChauffage
            EtatDefaut = .TEntreesAPI.E_DefautChauffage
            UnDefautAuMoinsSignale = UnDefautAuMoinsSignale Or EtatDefaut
            ConstruitListeNumDefauts EtatDefaut, NumDefaut, ListeNumDefauts, , ListeNumDefautsPourUneCuve  'contruction des listes de d�fauts
            EnregistreDateDetectionDisparitionDefaut EtatDefaut, NumDefaut 'enregistrement de la date de d�tection et disparition du d�faut
            
            '-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
            
            '--- d�faut PT100 ---
            NumDefaut = .TNumDefauts.NumDefautDefautPT100
            EtatDefaut = .TEntreesAPI.DefautPT100
            UnDefautAuMoinsSignale = UnDefautAuMoinsSignale Or EtatDefaut
            ConstruitListeNumDefauts EtatDefaut, NumDefaut, ListeNumDefauts, , ListeNumDefautsPourUneCuve  'contruction des listes de d�fauts
            EnregistreDateDetectionDisparitionDefaut EtatDefaut, NumDefaut 'enregistrement de la date de d�tection et disparition du d�faut
            
            '--- temp�rature trop basse ---
            NumDefaut = .TNumDefauts.NumDefautTemperatureTropBasse
            EtatDefaut = .TEntreesAPI.TemperatureTropBasse
            UnDefautAuMoinsSignale = UnDefautAuMoinsSignale Or EtatDefaut
            ConstruitListeNumDefauts EtatDefaut, NumDefaut, ListeNumDefauts, , ListeNumDefautsPourUneCuve  'contruction des listes de d�fauts
            EnregistreDateDetectionDisparitionDefaut EtatDefaut, NumDefaut 'enregistrement de la date de d�tection et disparition du d�faut
            
            '--- temp�rature trop haute ---
            NumDefaut = .TNumDefauts.NumDefautTemperatureTropHaute
            EtatDefaut = .TEntreesAPI.TemperatureTropHaute
            UnDefautAuMoinsSignale = UnDefautAuMoinsSignale Or EtatDefaut
            ConstruitListeNumDefauts EtatDefaut, NumDefaut, ListeNumDefauts, , ListeNumDefautsPourUneCuve  'contruction des listes de d�fauts
            EnregistreDateDetectionDisparitionDefaut EtatDefaut, NumDefaut 'enregistrement de la date de d�tection et disparition du d�faut

            '-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

        End If

        '--- construction de la liste des num�ros de d�fauts pour une cuve ayant une charge ---
        If NumChargeACePoste = 0 Then
            .ListeNumDefautsSiCharge = ""    'vider la liste si pas de charge dans le poste
        Else
            If ListeNumDefautsPourUneCuve <> "" Then  'remplir la liste en �vitant les doublons
                                                                                     'suppression du dernier s�parateur inutile car dans fonction
                .ListeNumDefautsSiCharge = AjoutNumDefautsSansDoublons(.ListeNumDefautsSiCharge, _
                                                                                                                     ListeNumDefautsPourUneCuve) 'remplir la liste en �liminant les doublons
            End If
        End If

        '--- affectation d�finitive ---
        .UnDefautAuMoinsSignale = UnDefautAuMoinsSignale

    End With
    
    '*********************************************************************************************************************
    '*                                                                CUVE C03
    '*********************************************************************************************************************
    
    '--- affectation ---
    If (False) Then
    
    'NumCuve = CUVES_REGULATION.C_C03
    ListeNumDefautsPourUneCuve = ""

    '--- recherche de la pr�sence d'une charge � ce poste ---
    NumChargeACePoste = TEtatsPostes(CorrespondanceCuvesAPIPostes(NumCuve)).NumCharge

    With TEtatsCuves(NumCuve)

        '--- initialisation � FALSE par d�faut ---
        UnDefautAuMoinsSignale = False

        If TEtatsLigne.MarcheGenerale = True Then

            '-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
            
            '--- niveau trop bas ---
            NumDefaut = .TNumDefauts.NumDefautNiveauTresBas
            EtatDefaut = .TEntreesAPI.E_NiveauTresBas
            UnDefautAuMoinsSignale = UnDefautAuMoinsSignale Or EtatDefaut
            ConstruitListeNumDefauts EtatDefaut, NumDefaut, ListeNumDefauts, , ListeNumDefautsPourUneCuve  'contruction des listes de d�fauts
            EnregistreDateDetectionDisparitionDefaut EtatDefaut, NumDefaut 'enregistrement de la date de d�tection et disparition du d�faut

            '--- niveau trop haut ---
            NumDefaut = .TNumDefauts.NumDefautNiveauTresHaut
            EtatDefaut = .TEntreesAPI.E_NiveauTresHaut
            UnDefautAuMoinsSignale = UnDefautAuMoinsSignale Or EtatDefaut
            ConstruitListeNumDefauts EtatDefaut, NumDefaut, ListeNumDefauts, , ListeNumDefautsPourUneCuve  'contruction des listes de d�fauts
            EnregistreDateDetectionDisparitionDefaut EtatDefaut, NumDefaut 'enregistrement de la date de d�tection et disparition du d�faut

            '-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

            '--- d�faut du chauffage ---
            NumDefaut = .TNumDefauts.NumDefautDefautChauffage
            EtatDefaut = .TEntreesAPI.E_DefautChauffage
            UnDefautAuMoinsSignale = UnDefautAuMoinsSignale Or EtatDefaut
            ConstruitListeNumDefauts EtatDefaut, NumDefaut, ListeNumDefauts, , ListeNumDefautsPourUneCuve  'contruction des listes de d�fauts
            EnregistreDateDetectionDisparitionDefaut EtatDefaut, NumDefaut 'enregistrement de la date de d�tection et disparition du d�faut
            
            '-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
            
            '--- d�faut PT100 ---
            NumDefaut = .TNumDefauts.NumDefautDefautPT100
            EtatDefaut = .TEntreesAPI.DefautPT100
            UnDefautAuMoinsSignale = UnDefautAuMoinsSignale Or EtatDefaut
            ConstruitListeNumDefauts EtatDefaut, NumDefaut, ListeNumDefauts, , ListeNumDefautsPourUneCuve  'contruction des listes de d�fauts
            EnregistreDateDetectionDisparitionDefaut EtatDefaut, NumDefaut 'enregistrement de la date de d�tection et disparition du d�faut
            
            '--- temp�rature trop basse ---
            NumDefaut = .TNumDefauts.NumDefautTemperatureTropBasse
            EtatDefaut = .TEntreesAPI.TemperatureTropBasse
            UnDefautAuMoinsSignale = UnDefautAuMoinsSignale Or EtatDefaut
            ConstruitListeNumDefauts EtatDefaut, NumDefaut, ListeNumDefauts, , ListeNumDefautsPourUneCuve  'contruction des listes de d�fauts
            EnregistreDateDetectionDisparitionDefaut EtatDefaut, NumDefaut 'enregistrement de la date de d�tection et disparition du d�faut
            
            '--- temp�rature trop haute ---
            NumDefaut = .TNumDefauts.NumDefautTemperatureTropHaute
            EtatDefaut = .TEntreesAPI.TemperatureTropHaute
            UnDefautAuMoinsSignale = UnDefautAuMoinsSignale Or EtatDefaut
            ConstruitListeNumDefauts EtatDefaut, NumDefaut, ListeNumDefauts, , ListeNumDefautsPourUneCuve  'contruction des listes de d�fauts
            EnregistreDateDetectionDisparitionDefaut EtatDefaut, NumDefaut 'enregistrement de la date de d�tection et disparition du d�faut

            '-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

        End If

        '--- construction de la liste des num�ros de d�fauts pour une cuve ayant une charge ---
        If NumChargeACePoste = 0 Then
            .ListeNumDefautsSiCharge = ""    'vider la liste si pas de charge dans le poste
        Else
            If ListeNumDefautsPourUneCuve <> "" Then  'remplir la liste en �vitant les doublons
                                                                                     'suppression du dernier s�parateur inutile car dans fonction
                .ListeNumDefautsSiCharge = AjoutNumDefautsSansDoublons(.ListeNumDefautsSiCharge, _
                                                                                                                     ListeNumDefautsPourUneCuve) 'remplir la liste en �liminant les doublons
            End If
        End If

        '--- affectation d�finitive ---
        .UnDefautAuMoinsSignale = UnDefautAuMoinsSignale

    End With

    End If
    
 
    
   
   
    
    '*********************************************************************************************************************
    '*                                                                CUVE C07
    '*********************************************************************************************************************
    
    '--- affectation ---
    NumCuve = CUVES_REGULATION.C_C07
    ListeNumDefautsPourUneCuve = ""
    
    '--- recherche de la pr�sence d'une charge � ce poste ---
    NumChargeACePoste = TEtatsPostes(CorrespondanceCuvesAPIPostes(NumCuve)).NumCharge

    With TEtatsCuves(NumCuve)

        '--- initialisation � FALSE par d�faut ---
        UnDefautAuMoinsSignale = False

        If TEtatsLigne.MarcheGenerale = True Then

            '-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
            
            '--- niveau trop bas ---
            NumDefaut = .TNumDefauts.NumDefautNiveauTresBas
            EtatDefaut = .TEntreesAPI.E_NiveauTresBas
            UnDefautAuMoinsSignale = UnDefautAuMoinsSignale Or EtatDefaut
            ConstruitListeNumDefauts EtatDefaut, NumDefaut, ListeNumDefauts, , ListeNumDefautsPourUneCuve  'contruction des listes de d�fauts
            EnregistreDateDetectionDisparitionDefaut EtatDefaut, NumDefaut 'enregistrement de la date de d�tection et disparition du d�faut

            '--- niveau trop haut ---
            NumDefaut = .TNumDefauts.NumDefautNiveauTresHaut
            EtatDefaut = .TEntreesAPI.E_NiveauTresHaut
            UnDefautAuMoinsSignale = UnDefautAuMoinsSignale Or EtatDefaut
            ConstruitListeNumDefauts EtatDefaut, NumDefaut, ListeNumDefauts, , ListeNumDefautsPourUneCuve  'contruction des listes de d�fauts
            EnregistreDateDetectionDisparitionDefaut EtatDefaut, NumDefaut 'enregistrement de la date de d�tection et disparition du d�faut

            '-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

            '--- d�faut du chauffage ---
            NumDefaut = .TNumDefauts.NumDefautDefautChauffage
            EtatDefaut = .TEntreesAPI.E_DefautChauffage
            UnDefautAuMoinsSignale = UnDefautAuMoinsSignale Or EtatDefaut
            ConstruitListeNumDefauts EtatDefaut, NumDefaut, ListeNumDefauts, , ListeNumDefautsPourUneCuve  'contruction des listes de d�fauts
            EnregistreDateDetectionDisparitionDefaut EtatDefaut, NumDefaut 'enregistrement de la date de d�tection et disparition du d�faut
            
            '-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
            
            '--- d�faut PT100 ---
            NumDefaut = .TNumDefauts.NumDefautDefautPT100
            EtatDefaut = .TEntreesAPI.DefautPT100
            UnDefautAuMoinsSignale = UnDefautAuMoinsSignale Or EtatDefaut
            ConstruitListeNumDefauts EtatDefaut, NumDefaut, ListeNumDefauts, , ListeNumDefautsPourUneCuve  'contruction des listes de d�fauts
            EnregistreDateDetectionDisparitionDefaut EtatDefaut, NumDefaut 'enregistrement de la date de d�tection et disparition du d�faut
            
            '--- temp�rature trop basse ---
            NumDefaut = .TNumDefauts.NumDefautTemperatureTropBasse
            EtatDefaut = .TEntreesAPI.TemperatureTropBasse
            UnDefautAuMoinsSignale = UnDefautAuMoinsSignale Or EtatDefaut
            ConstruitListeNumDefauts EtatDefaut, NumDefaut, ListeNumDefauts, , ListeNumDefautsPourUneCuve  'contruction des listes de d�fauts
            EnregistreDateDetectionDisparitionDefaut EtatDefaut, NumDefaut 'enregistrement de la date de d�tection et disparition du d�faut
            
            '--- temp�rature trop haute ---
            NumDefaut = .TNumDefauts.NumDefautTemperatureTropHaute
            EtatDefaut = .TEntreesAPI.TemperatureTropHaute
            UnDefautAuMoinsSignale = UnDefautAuMoinsSignale Or EtatDefaut
            ConstruitListeNumDefauts EtatDefaut, NumDefaut, ListeNumDefauts, , ListeNumDefautsPourUneCuve  'contruction des listes de d�fauts
            EnregistreDateDetectionDisparitionDefaut EtatDefaut, NumDefaut 'enregistrement de la date de d�tection et disparition du d�faut

            '-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

        End If

        '--- construction de la liste des num�ros de d�fauts pour une cuve ayant une charge ---
        If NumChargeACePoste = 0 Then
            .ListeNumDefautsSiCharge = ""    'vider la liste si pas de charge dans le poste
        Else
            If ListeNumDefautsPourUneCuve <> "" Then  'remplir la liste en �vitant les doublons
                                                                                     'suppression du dernier s�parateur inutile car dans fonction
                .ListeNumDefautsSiCharge = AjoutNumDefautsSansDoublons(.ListeNumDefautsSiCharge, _
                                                                                                                     ListeNumDefautsPourUneCuve) 'remplir la liste en �liminant les doublons
            End If
        End If

        '--- affectation d�finitive ---
        .UnDefautAuMoinsSignale = UnDefautAuMoinsSignale

    End With
    
    '*********************************************************************************************************************
    '*                                                                CUVE C13
    '*********************************************************************************************************************
    
    '--- affectation ---
    NumCuve = CUVES_REGULATION.C_C13
    ListeNumDefautsPourUneCuve = ""
    
    '--- recherche de la pr�sence d'une charge � ce poste ---
    NumChargeACePoste = TEtatsPostes(CorrespondanceCuvesAPIPostes(NumCuve)).NumCharge

    With TEtatsCuves(NumCuve)

        '--- initialisation � FALSE par d�faut ---
        UnDefautAuMoinsSignale = False

        If TEtatsLigne.MarcheGenerale = True Then

            '-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
            
            '--- niveau trop bas ---
            NumDefaut = .TNumDefauts.NumDefautNiveauTresBas
            EtatDefaut = .TEntreesAPI.E_NiveauTresBas
            UnDefautAuMoinsSignale = UnDefautAuMoinsSignale Or EtatDefaut
            ConstruitListeNumDefauts EtatDefaut, NumDefaut, ListeNumDefauts, , ListeNumDefautsPourUneCuve  'contruction des listes de d�fauts
            EnregistreDateDetectionDisparitionDefaut EtatDefaut, NumDefaut 'enregistrement de la date de d�tection et disparition du d�faut

            '--- niveau trop haut ---
            NumDefaut = .TNumDefauts.NumDefautNiveauTresHaut
            EtatDefaut = .TEntreesAPI.E_NiveauTresHaut
            UnDefautAuMoinsSignale = UnDefautAuMoinsSignale Or EtatDefaut
            ConstruitListeNumDefauts EtatDefaut, NumDefaut, ListeNumDefauts, , ListeNumDefautsPourUneCuve  'contruction des listes de d�fauts
            EnregistreDateDetectionDisparitionDefaut EtatDefaut, NumDefaut 'enregistrement de la date de d�tection et disparition du d�faut

            '-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

            '--- d�faut du chauffage ---
            NumDefaut = .TNumDefauts.NumDefautDefautChauffage
            EtatDefaut = .TEntreesAPI.E_DefautChauffage
            UnDefautAuMoinsSignale = UnDefautAuMoinsSignale Or EtatDefaut
            ConstruitListeNumDefauts EtatDefaut, NumDefaut, ListeNumDefauts, , ListeNumDefautsPourUneCuve  'contruction des listes de d�fauts
            EnregistreDateDetectionDisparitionDefaut EtatDefaut, NumDefaut 'enregistrement de la date de d�tection et disparition du d�faut
            
            '-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
            
            '--- d�faut PT100 ---
            NumDefaut = .TNumDefauts.NumDefautDefautPT100
            EtatDefaut = .TEntreesAPI.DefautPT100
            UnDefautAuMoinsSignale = UnDefautAuMoinsSignale Or EtatDefaut
            ConstruitListeNumDefauts EtatDefaut, NumDefaut, ListeNumDefauts, , ListeNumDefautsPourUneCuve  'contruction des listes de d�fauts
            EnregistreDateDetectionDisparitionDefaut EtatDefaut, NumDefaut 'enregistrement de la date de d�tection et disparition du d�faut
            
            '--- temp�rature trop basse ---
            NumDefaut = .TNumDefauts.NumDefautTemperatureTropBasse
            EtatDefaut = .TEntreesAPI.TemperatureTropBasse
            UnDefautAuMoinsSignale = UnDefautAuMoinsSignale Or EtatDefaut
            ConstruitListeNumDefauts EtatDefaut, NumDefaut, ListeNumDefauts, , ListeNumDefautsPourUneCuve  'contruction des listes de d�fauts
            EnregistreDateDetectionDisparitionDefaut EtatDefaut, NumDefaut 'enregistrement de la date de d�tection et disparition du d�faut
            
            '--- temp�rature trop haute ---
            NumDefaut = .TNumDefauts.NumDefautTemperatureTropHaute
            EtatDefaut = .TEntreesAPI.TemperatureTropHaute
            UnDefautAuMoinsSignale = UnDefautAuMoinsSignale Or EtatDefaut
            ConstruitListeNumDefauts EtatDefaut, NumDefaut, ListeNumDefauts, , ListeNumDefautsPourUneCuve  'contruction des listes de d�fauts
            EnregistreDateDetectionDisparitionDefaut EtatDefaut, NumDefaut 'enregistrement de la date de d�tection et disparition du d�faut

            '-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

            '--- d�faut refroidissement ---
            NumDefaut = .TNumDefauts.NumDefautDefautRefroidissement
            EtatDefaut = .TEntreesAPI.E_DefautRefroidissement
            UnDefautAuMoinsSignale = UnDefautAuMoinsSignale Or EtatDefaut
            ConstruitListeNumDefauts EtatDefaut, NumDefaut, ListeNumDefauts, , ListeNumDefautsPourUneCuve  'contruction des listes de d�fauts
            EnregistreDateDetectionDisparitionDefaut EtatDefaut, NumDefaut 'enregistrement de la date de d�tection et disparition du d�faut
        
            '-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
        
        End If

        '--- construction de la liste des num�ros de d�fauts pour une cuve ayant une charge ---
        If NumChargeACePoste = 0 Then
            .ListeNumDefautsSiCharge = ""    'vider la liste si pas de charge dans le poste
        Else
            If ListeNumDefautsPourUneCuve <> "" Then  'remplir la liste en �vitant les doublons
                                                                                     'suppression du dernier s�parateur inutile car dans fonction
                .ListeNumDefautsSiCharge = AjoutNumDefautsSansDoublons(.ListeNumDefautsSiCharge, _
                                                                                                                     ListeNumDefautsPourUneCuve) 'remplir la liste en �liminant les doublons
            End If
        End If

        '--- affectation d�finitive ---
        .UnDefautAuMoinsSignale = UnDefautAuMoinsSignale

    End With
    
    '*********************************************************************************************************************
    '*                                                                CUVE C14
    '*********************************************************************************************************************
    
    '--- affectation ---
    NumCuve = CUVES_REGULATION.C_C14
    ListeNumDefautsPourUneCuve = ""
    
    '--- recherche de la pr�sence d'une charge � ce poste ---
    NumChargeACePoste = TEtatsPostes(CorrespondanceCuvesAPIPostes(NumCuve)).NumCharge

    With TEtatsCuves(NumCuve)

        '--- initialisation � FALSE par d�faut ---
        UnDefautAuMoinsSignale = False

        If TEtatsLigne.MarcheGenerale = True Then

            '-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
            
            '--- niveau trop bas ---
            NumDefaut = .TNumDefauts.NumDefautNiveauTresBas
            EtatDefaut = .TEntreesAPI.E_NiveauTresBas
            UnDefautAuMoinsSignale = UnDefautAuMoinsSignale Or EtatDefaut
            ConstruitListeNumDefauts EtatDefaut, NumDefaut, ListeNumDefauts, , ListeNumDefautsPourUneCuve  'contruction des listes de d�fauts
            EnregistreDateDetectionDisparitionDefaut EtatDefaut, NumDefaut 'enregistrement de la date de d�tection et disparition du d�faut

            '--- niveau trop haut ---
            NumDefaut = .TNumDefauts.NumDefautNiveauTresHaut
            EtatDefaut = .TEntreesAPI.E_NiveauTresHaut
            UnDefautAuMoinsSignale = UnDefautAuMoinsSignale Or EtatDefaut
            ConstruitListeNumDefauts EtatDefaut, NumDefaut, ListeNumDefauts, , ListeNumDefautsPourUneCuve  'contruction des listes de d�fauts
            EnregistreDateDetectionDisparitionDefaut EtatDefaut, NumDefaut 'enregistrement de la date de d�tection et disparition du d�faut

            '-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

            '--- d�faut du chauffage ---
            NumDefaut = .TNumDefauts.NumDefautDefautChauffage
            EtatDefaut = .TEntreesAPI.E_DefautChauffage
            UnDefautAuMoinsSignale = UnDefautAuMoinsSignale Or EtatDefaut
            ConstruitListeNumDefauts EtatDefaut, NumDefaut, ListeNumDefauts, , ListeNumDefautsPourUneCuve  'contruction des listes de d�fauts
            EnregistreDateDetectionDisparitionDefaut EtatDefaut, NumDefaut 'enregistrement de la date de d�tection et disparition du d�faut
            
            '-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
            
            '--- d�faut PT100 ---
            NumDefaut = .TNumDefauts.NumDefautDefautPT100
            EtatDefaut = .TEntreesAPI.DefautPT100
            UnDefautAuMoinsSignale = UnDefautAuMoinsSignale Or EtatDefaut
            ConstruitListeNumDefauts EtatDefaut, NumDefaut, ListeNumDefauts, , ListeNumDefautsPourUneCuve  'contruction des listes de d�fauts
            EnregistreDateDetectionDisparitionDefaut EtatDefaut, NumDefaut 'enregistrement de la date de d�tection et disparition du d�faut
            
            '--- temp�rature trop basse ---
            NumDefaut = .TNumDefauts.NumDefautTemperatureTropBasse
            EtatDefaut = .TEntreesAPI.TemperatureTropBasse
            UnDefautAuMoinsSignale = UnDefautAuMoinsSignale Or EtatDefaut
            ConstruitListeNumDefauts EtatDefaut, NumDefaut, ListeNumDefauts, , ListeNumDefautsPourUneCuve  'contruction des listes de d�fauts
            EnregistreDateDetectionDisparitionDefaut EtatDefaut, NumDefaut 'enregistrement de la date de d�tection et disparition du d�faut
            
            '--- temp�rature trop haute ---
            NumDefaut = .TNumDefauts.NumDefautTemperatureTropHaute
            EtatDefaut = .TEntreesAPI.TemperatureTropHaute
            UnDefautAuMoinsSignale = UnDefautAuMoinsSignale Or EtatDefaut
            ConstruitListeNumDefauts EtatDefaut, NumDefaut, ListeNumDefauts, , ListeNumDefautsPourUneCuve  'contruction des listes de d�fauts
            EnregistreDateDetectionDisparitionDefaut EtatDefaut, NumDefaut 'enregistrement de la date de d�tection et disparition du d�faut

            '-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

            '--- d�faut refroidissement ---
            NumDefaut = .TNumDefauts.NumDefautDefautRefroidissement
            EtatDefaut = .TEntreesAPI.E_DefautRefroidissement
            UnDefautAuMoinsSignale = UnDefautAuMoinsSignale Or EtatDefaut
            ConstruitListeNumDefauts EtatDefaut, NumDefaut, ListeNumDefauts, , ListeNumDefautsPourUneCuve  'contruction des listes de d�fauts
            EnregistreDateDetectionDisparitionDefaut EtatDefaut, NumDefaut 'enregistrement de la date de d�tection et disparition du d�faut
        
            '-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
        
        End If

        '--- construction de la liste des num�ros de d�fauts pour une cuve ayant une charge ---
        If NumChargeACePoste = 0 Then
            .ListeNumDefautsSiCharge = ""    'vider la liste si pas de charge dans le poste
        Else
            If ListeNumDefautsPourUneCuve <> "" Then  'remplir la liste en �vitant les doublons
                                                                                     'suppression du dernier s�parateur inutile car dans fonction
                .ListeNumDefautsSiCharge = AjoutNumDefautsSansDoublons(.ListeNumDefautsSiCharge, _
                                                                                                                     ListeNumDefautsPourUneCuve) 'remplir la liste en �liminant les doublons
            End If
        End If

        '--- affectation d�finitive ---
        .UnDefautAuMoinsSignale = UnDefautAuMoinsSignale

    End With
    
    '*********************************************************************************************************************
    '*                                                                CUVE C15
    '*********************************************************************************************************************
    
    '--- affectation ---
    NumCuve = CUVES_REGULATION.C_C15
    ListeNumDefautsPourUneCuve = ""
    
    '--- recherche de la pr�sence d'une charge � ce poste ---
    NumChargeACePoste = TEtatsPostes(CorrespondanceCuvesAPIPostes(NumCuve)).NumCharge

    With TEtatsCuves(NumCuve)

        '--- initialisation � FALSE par d�faut ---
        UnDefautAuMoinsSignale = False

        If TEtatsLigne.MarcheGenerale = True Then

            '-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

            '--- niveau trop bas ---
            NumDefaut = .TNumDefauts.NumDefautNiveauTresBas
            EtatDefaut = .TEntreesAPI.E_NiveauTresBas
            UnDefautAuMoinsSignale = UnDefautAuMoinsSignale Or EtatDefaut
            ConstruitListeNumDefauts EtatDefaut, NumDefaut, ListeNumDefauts, , ListeNumDefautsPourUneCuve  'contruction des listes de d�fauts
            EnregistreDateDetectionDisparitionDefaut EtatDefaut, NumDefaut 'enregistrement de la date de d�tection et disparition du d�faut

            '--- niveau trop haut ---
            NumDefaut = .TNumDefauts.NumDefautNiveauTresHaut
            EtatDefaut = .TEntreesAPI.E_NiveauTresHaut
            UnDefautAuMoinsSignale = UnDefautAuMoinsSignale Or EtatDefaut
            ConstruitListeNumDefauts EtatDefaut, NumDefaut, ListeNumDefauts, , ListeNumDefautsPourUneCuve  'contruction des listes de d�fauts
            EnregistreDateDetectionDisparitionDefaut EtatDefaut, NumDefaut 'enregistrement de la date de d�tection et disparition du d�faut

            '-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

            '--- d�faut du chauffage ---
            NumDefaut = .TNumDefauts.NumDefautDefautChauffage
            EtatDefaut = .TEntreesAPI.E_DefautChauffage
            UnDefautAuMoinsSignale = UnDefautAuMoinsSignale Or EtatDefaut
            ConstruitListeNumDefauts EtatDefaut, NumDefaut, ListeNumDefauts, , ListeNumDefautsPourUneCuve  'contruction des listes de d�fauts
            EnregistreDateDetectionDisparitionDefaut EtatDefaut, NumDefaut 'enregistrement de la date de d�tection et disparition du d�faut
            
            '-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
            
            '--- d�faut PT100 ---
            NumDefaut = .TNumDefauts.NumDefautDefautPT100
            EtatDefaut = .TEntreesAPI.DefautPT100
            UnDefautAuMoinsSignale = UnDefautAuMoinsSignale Or EtatDefaut
            ConstruitListeNumDefauts EtatDefaut, NumDefaut, ListeNumDefauts, , ListeNumDefautsPourUneCuve  'contruction des listes de d�fauts
            EnregistreDateDetectionDisparitionDefaut EtatDefaut, NumDefaut 'enregistrement de la date de d�tection et disparition du d�faut
            
            '--- temp�rature trop basse ---
            NumDefaut = .TNumDefauts.NumDefautTemperatureTropBasse
            EtatDefaut = .TEntreesAPI.TemperatureTropBasse
            UnDefautAuMoinsSignale = UnDefautAuMoinsSignale Or EtatDefaut
            ConstruitListeNumDefauts EtatDefaut, NumDefaut, ListeNumDefauts, , ListeNumDefautsPourUneCuve  'contruction des listes de d�fauts
            EnregistreDateDetectionDisparitionDefaut EtatDefaut, NumDefaut 'enregistrement de la date de d�tection et disparition du d�faut
            
            '--- temp�rature trop haute ---
            NumDefaut = .TNumDefauts.NumDefautTemperatureTropHaute
            EtatDefaut = .TEntreesAPI.TemperatureTropHaute
            UnDefautAuMoinsSignale = UnDefautAuMoinsSignale Or EtatDefaut
            ConstruitListeNumDefauts EtatDefaut, NumDefaut, ListeNumDefauts, , ListeNumDefautsPourUneCuve  'contruction des listes de d�fauts
            EnregistreDateDetectionDisparitionDefaut EtatDefaut, NumDefaut 'enregistrement de la date de d�tection et disparition du d�faut

            '-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

            '--- d�faut refroidissement ---
            NumDefaut = .TNumDefauts.NumDefautDefautRefroidissement
            EtatDefaut = .TEntreesAPI.E_DefautRefroidissement
            UnDefautAuMoinsSignale = UnDefautAuMoinsSignale Or EtatDefaut
            ConstruitListeNumDefauts EtatDefaut, NumDefaut, ListeNumDefauts, , ListeNumDefautsPourUneCuve  'contruction des listes de d�fauts
            EnregistreDateDetectionDisparitionDefaut EtatDefaut, NumDefaut 'enregistrement de la date de d�tection et disparition du d�faut
        
            '-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
        
        End If

        '--- construction de la liste des num�ros de d�fauts pour une cuve ayant une charge ---
        If NumChargeACePoste = 0 Then
            .ListeNumDefautsSiCharge = ""    'vider la liste si pas de charge dans le poste
        Else
            If ListeNumDefautsPourUneCuve <> "" Then  'remplir la liste en �vitant les doublons
                                                                                     'suppression du dernier s�parateur inutile car dans fonction
                .ListeNumDefautsSiCharge = AjoutNumDefautsSansDoublons(.ListeNumDefautsSiCharge, _
                                                                                                                     ListeNumDefautsPourUneCuve) 'remplir la liste en �liminant les doublons
            End If
        End If

        '--- affectation d�finitive ---
        .UnDefautAuMoinsSignale = UnDefautAuMoinsSignale

    End With
    
    '*********************************************************************************************************************
    '*                                                                CUVE C16
    '*********************************************************************************************************************
    If False Then
    '--- affectation ---
    'NumCuve = CUVES_REGULATION.C_C16
    ListeNumDefautsPourUneCuve = ""
    
    '--- recherche de la pr�sence d'une charge � ce poste ---
    NumChargeACePoste = TEtatsPostes(CorrespondanceCuvesAPIPostes(NumCuve)).NumCharge

    With TEtatsCuves(NumCuve)

        '--- initialisation � FALSE par d�faut ---
        UnDefautAuMoinsSignale = False

        If TEtatsLigne.MarcheGenerale = True Then

            '-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

            '--- niveau trop bas ---
            NumDefaut = .TNumDefauts.NumDefautNiveauTresBas
            EtatDefaut = .TEntreesAPI.E_NiveauTresBas
            UnDefautAuMoinsSignale = UnDefautAuMoinsSignale Or EtatDefaut
            ConstruitListeNumDefauts EtatDefaut, NumDefaut, ListeNumDefauts, , ListeNumDefautsPourUneCuve  'contruction des listes de d�fauts
            EnregistreDateDetectionDisparitionDefaut EtatDefaut, NumDefaut 'enregistrement de la date de d�tection et disparition du d�faut

            '--- niveau trop haut ---
            NumDefaut = .TNumDefauts.NumDefautNiveauTresHaut
            EtatDefaut = .TEntreesAPI.E_NiveauTresHaut
            UnDefautAuMoinsSignale = UnDefautAuMoinsSignale Or EtatDefaut
            ConstruitListeNumDefauts EtatDefaut, NumDefaut, ListeNumDefauts, , ListeNumDefautsPourUneCuve  'contruction des listes de d�fauts
            EnregistreDateDetectionDisparitionDefaut EtatDefaut, NumDefaut 'enregistrement de la date de d�tection et disparition du d�faut

            '-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

            '--- d�faut du chauffage ---
            NumDefaut = .TNumDefauts.NumDefautDefautChauffage
            EtatDefaut = .TEntreesAPI.E_DefautChauffage
            UnDefautAuMoinsSignale = UnDefautAuMoinsSignale Or EtatDefaut
            ConstruitListeNumDefauts EtatDefaut, NumDefaut, ListeNumDefauts, , ListeNumDefautsPourUneCuve  'contruction des listes de d�fauts
            EnregistreDateDetectionDisparitionDefaut EtatDefaut, NumDefaut 'enregistrement de la date de d�tection et disparition du d�faut
            
            '-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
            
            '--- d�faut PT100 ---
            NumDefaut = .TNumDefauts.NumDefautDefautPT100
            EtatDefaut = .TEntreesAPI.DefautPT100
            UnDefautAuMoinsSignale = UnDefautAuMoinsSignale Or EtatDefaut
            ConstruitListeNumDefauts EtatDefaut, NumDefaut, ListeNumDefauts, , ListeNumDefautsPourUneCuve  'contruction des listes de d�fauts
            EnregistreDateDetectionDisparitionDefaut EtatDefaut, NumDefaut 'enregistrement de la date de d�tection et disparition du d�faut
            
            '--- temp�rature trop basse ---
            NumDefaut = .TNumDefauts.NumDefautTemperatureTropBasse
            EtatDefaut = .TEntreesAPI.TemperatureTropBasse
            UnDefautAuMoinsSignale = UnDefautAuMoinsSignale Or EtatDefaut
            ConstruitListeNumDefauts EtatDefaut, NumDefaut, ListeNumDefauts, , ListeNumDefautsPourUneCuve  'contruction des listes de d�fauts
            EnregistreDateDetectionDisparitionDefaut EtatDefaut, NumDefaut 'enregistrement de la date de d�tection et disparition du d�faut
            
            '--- temp�rature trop haute ---
            NumDefaut = .TNumDefauts.NumDefautTemperatureTropHaute
            EtatDefaut = .TEntreesAPI.TemperatureTropHaute
            UnDefautAuMoinsSignale = UnDefautAuMoinsSignale Or EtatDefaut
            ConstruitListeNumDefauts EtatDefaut, NumDefaut, ListeNumDefauts, , ListeNumDefautsPourUneCuve  'contruction des listes de d�fauts
            EnregistreDateDetectionDisparitionDefaut EtatDefaut, NumDefaut 'enregistrement de la date de d�tection et disparition du d�faut

            '-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

            '--- d�faut refroidissement ---
            NumDefaut = .TNumDefauts.NumDefautDefautRefroidissement
            EtatDefaut = .TEntreesAPI.E_DefautRefroidissement
            UnDefautAuMoinsSignale = UnDefautAuMoinsSignale Or EtatDefaut
            ConstruitListeNumDefauts EtatDefaut, NumDefaut, ListeNumDefauts, , ListeNumDefautsPourUneCuve  'contruction des listes de d�fauts
            EnregistreDateDetectionDisparitionDefaut EtatDefaut, NumDefaut 'enregistrement de la date de d�tection et disparition du d�faut
        
            '-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
        
        End If

        '--- construction de la liste des num�ros de d�fauts pour une cuve ayant une charge ---
        If NumChargeACePoste = 0 Then
            .ListeNumDefautsSiCharge = ""    'vider la liste si pas de charge dans le poste
        Else
            If ListeNumDefautsPourUneCuve <> "" Then  'remplir la liste en �vitant les doublons
                                                                                     'suppression du dernier s�parateur inutile car dans fonction
                .ListeNumDefautsSiCharge = AjoutNumDefautsSansDoublons(.ListeNumDefautsSiCharge, _
                                                                                                                     ListeNumDefautsPourUneCuve) 'remplir la liste en �liminant les doublons
            End If
        End If

        '--- affectation d�finitive ---
        .UnDefautAuMoinsSignale = UnDefautAuMoinsSignale

    End With
    

    
    End If
    
    
    
    '*********************************************************************************************************************
    '*                                                                CUVE C19
    '*********************************************************************************************************************
    If (False) Then
     
    '--- affectation ---
    'NumCuve = CUVES_REGULATION.C_C19
    
    ListeNumDefautsPourUneCuve = ""
    
    '--- recherche de la pr�sence d'une charge � ce poste ---
    NumChargeACePoste = TEtatsPostes(CorrespondanceCuvesAPIPostes(NumCuve)).NumCharge

    With TEtatsCuves(NumCuve)

        '--- initialisation � FALSE par d�faut ---
        UnDefautAuMoinsSignale = False

        If TEtatsLigne.MarcheGenerale = True Then
            
            '-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
            
            '--- niveau trop bas ---
            NumDefaut = .TNumDefauts.NumDefautNiveauTresBas
            EtatDefaut = .TEntreesAPI.E_NiveauTresBas
            UnDefautAuMoinsSignale = UnDefautAuMoinsSignale Or EtatDefaut
            ConstruitListeNumDefauts EtatDefaut, NumDefaut, ListeNumDefauts, , ListeNumDefautsPourUneCuve  'contruction des listes de d�fauts
            EnregistreDateDetectionDisparitionDefaut EtatDefaut, NumDefaut 'enregistrement de la date de d�tection et disparition du d�faut

            '--- niveau trop haut ---
            NumDefaut = .TNumDefauts.NumDefautNiveauTresHaut
            EtatDefaut = .TEntreesAPI.E_NiveauTresHaut
            UnDefautAuMoinsSignale = UnDefautAuMoinsSignale Or EtatDefaut
            ConstruitListeNumDefauts EtatDefaut, NumDefaut, ListeNumDefauts, , ListeNumDefautsPourUneCuve  'contruction des listes de d�fauts
            EnregistreDateDetectionDisparitionDefaut EtatDefaut, NumDefaut 'enregistrement de la date de d�tection et disparition du d�faut

            '-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

            '--- d�faut du chauffage ---
            NumDefaut = .TNumDefauts.NumDefautDefautChauffage
            EtatDefaut = .TEntreesAPI.E_DefautChauffage
            UnDefautAuMoinsSignale = UnDefautAuMoinsSignale Or EtatDefaut
            ConstruitListeNumDefauts EtatDefaut, NumDefaut, ListeNumDefauts, , ListeNumDefautsPourUneCuve  'contruction des listes de d�fauts
            EnregistreDateDetectionDisparitionDefaut EtatDefaut, NumDefaut 'enregistrement de la date de d�tection et disparition du d�faut
            
            '-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
            
            '--- d�faut PT100 ---
            NumDefaut = .TNumDefauts.NumDefautDefautPT100
            EtatDefaut = .TEntreesAPI.DefautPT100
            UnDefautAuMoinsSignale = UnDefautAuMoinsSignale Or EtatDefaut
            ConstruitListeNumDefauts EtatDefaut, NumDefaut, ListeNumDefauts, , ListeNumDefautsPourUneCuve  'contruction des listes de d�fauts
            EnregistreDateDetectionDisparitionDefaut EtatDefaut, NumDefaut 'enregistrement de la date de d�tection et disparition du d�faut
            
            '--- temp�rature trop basse ---
            NumDefaut = .TNumDefauts.NumDefautTemperatureTropBasse
            EtatDefaut = .TEntreesAPI.TemperatureTropBasse
            UnDefautAuMoinsSignale = UnDefautAuMoinsSignale Or EtatDefaut
            ConstruitListeNumDefauts EtatDefaut, NumDefaut, ListeNumDefauts, , ListeNumDefautsPourUneCuve  'contruction des listes de d�fauts
            EnregistreDateDetectionDisparitionDefaut EtatDefaut, NumDefaut 'enregistrement de la date de d�tection et disparition du d�faut
            
            '--- temp�rature trop haute ---
            NumDefaut = .TNumDefauts.NumDefautTemperatureTropHaute
            EtatDefaut = .TEntreesAPI.TemperatureTropHaute
            UnDefautAuMoinsSignale = UnDefautAuMoinsSignale Or EtatDefaut
            ConstruitListeNumDefauts EtatDefaut, NumDefaut, ListeNumDefauts, , ListeNumDefautsPourUneCuve  'contruction des listes de d�fauts
            EnregistreDateDetectionDisparitionDefaut EtatDefaut, NumDefaut 'enregistrement de la date de d�tection et disparition du d�faut

            '-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

        End If

        '--- construction de la liste des num�ros de d�fauts pour une cuve ayant une charge ---
        If NumChargeACePoste = 0 Then
            .ListeNumDefautsSiCharge = ""    'vider la liste si pas de charge dans le poste
        Else
            If ListeNumDefautsPourUneCuve <> "" Then  'remplir la liste en �vitant les doublons
                                                                                     'suppression du dernier s�parateur inutile car dans fonction
                .ListeNumDefautsSiCharge = AjoutNumDefautsSansDoublons(.ListeNumDefautsSiCharge, _
                                                                                                                     ListeNumDefautsPourUneCuve) 'remplir la liste en �liminant les doublons
            End If
        End If

        '--- affectation d�finitive ---
        .UnDefautAuMoinsSignale = UnDefautAuMoinsSignale

    End With
    End If
   
    
    '*********************************************************************************************************************
    '*                                                                CUVE C22
    '*********************************************************************************************************************
    
    '--- affectation ---
    NumCuve = CUVES_REGULATION.C_C22
    ListeNumDefautsPourUneCuve = ""
    
    '--- recherche de la pr�sence d'une charge � ce poste ---
    NumChargeACePoste = TEtatsPostes(CorrespondanceCuvesAPIPostes(NumCuve)).NumCharge

    With TEtatsCuves(NumCuve)

        '--- initialisation � FALSE par d�faut ---
        UnDefautAuMoinsSignale = False

        If TEtatsLigne.MarcheGenerale = True Then
            
            '-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
            
            '--- niveau trop bas ---
            NumDefaut = .TNumDefauts.NumDefautNiveauTresBas
            EtatDefaut = .TEntreesAPI.E_NiveauTresBas
            UnDefautAuMoinsSignale = UnDefautAuMoinsSignale Or EtatDefaut
            ConstruitListeNumDefauts EtatDefaut, NumDefaut, ListeNumDefauts, , ListeNumDefautsPourUneCuve  'contruction des listes de d�fauts
            EnregistreDateDetectionDisparitionDefaut EtatDefaut, NumDefaut 'enregistrement de la date de d�tection et disparition du d�faut

            '--- niveau trop haut ---
            NumDefaut = .TNumDefauts.NumDefautNiveauTresHaut
            EtatDefaut = .TEntreesAPI.E_NiveauTresHaut
            UnDefautAuMoinsSignale = UnDefautAuMoinsSignale Or EtatDefaut
            ConstruitListeNumDefauts EtatDefaut, NumDefaut, ListeNumDefauts, , ListeNumDefautsPourUneCuve  'contruction des listes de d�fauts
            EnregistreDateDetectionDisparitionDefaut EtatDefaut, NumDefaut 'enregistrement de la date de d�tection et disparition du d�faut

            '-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

            '--- d�faut du chauffage ---
            NumDefaut = .TNumDefauts.NumDefautDefautChauffage
            EtatDefaut = .TEntreesAPI.E_DefautChauffage
            UnDefautAuMoinsSignale = UnDefautAuMoinsSignale Or EtatDefaut
            ConstruitListeNumDefauts EtatDefaut, NumDefaut, ListeNumDefauts, , ListeNumDefautsPourUneCuve  'contruction des listes de d�fauts
            EnregistreDateDetectionDisparitionDefaut EtatDefaut, NumDefaut 'enregistrement de la date de d�tection et disparition du d�faut
            
            '-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
            
            '--- d�faut PT100 ---
            NumDefaut = .TNumDefauts.NumDefautDefautPT100
            EtatDefaut = .TEntreesAPI.DefautPT100
            UnDefautAuMoinsSignale = UnDefautAuMoinsSignale Or EtatDefaut
            ConstruitListeNumDefauts EtatDefaut, NumDefaut, ListeNumDefauts, , ListeNumDefautsPourUneCuve  'contruction des listes de d�fauts
            EnregistreDateDetectionDisparitionDefaut EtatDefaut, NumDefaut 'enregistrement de la date de d�tection et disparition du d�faut
            
            '--- temp�rature trop basse ---
            NumDefaut = .TNumDefauts.NumDefautTemperatureTropBasse
            EtatDefaut = .TEntreesAPI.TemperatureTropBasse
            UnDefautAuMoinsSignale = UnDefautAuMoinsSignale Or EtatDefaut
            ConstruitListeNumDefauts EtatDefaut, NumDefaut, ListeNumDefauts, , ListeNumDefautsPourUneCuve  'contruction des listes de d�fauts
            EnregistreDateDetectionDisparitionDefaut EtatDefaut, NumDefaut 'enregistrement de la date de d�tection et disparition du d�faut
            
            '--- temp�rature trop haute ---
            NumDefaut = .TNumDefauts.NumDefautTemperatureTropHaute
            EtatDefaut = .TEntreesAPI.TemperatureTropHaute
            UnDefautAuMoinsSignale = UnDefautAuMoinsSignale Or EtatDefaut
            ConstruitListeNumDefauts EtatDefaut, NumDefaut, ListeNumDefauts, , ListeNumDefautsPourUneCuve  'contruction des listes de d�fauts
            EnregistreDateDetectionDisparitionDefaut EtatDefaut, NumDefaut 'enregistrement de la date de d�tection et disparition du d�faut

            '-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

        End If

        '--- construction de la liste des num�ros de d�fauts pour une cuve ayant une charge ---
        If NumChargeACePoste = 0 Then
            .ListeNumDefautsSiCharge = ""    'vider la liste si pas de charge dans le poste
        Else
            If ListeNumDefautsPourUneCuve <> "" Then  'remplir la liste en �vitant les doublons
                                                                                     'suppression du dernier s�parateur inutile car dans fonction
                .ListeNumDefautsSiCharge = AjoutNumDefautsSansDoublons(.ListeNumDefautsSiCharge, _
                                                                                                                     ListeNumDefautsPourUneCuve) 'remplir la liste en �liminant les doublons
            End If
        End If

        '--- affectation d�finitive ---
        .UnDefautAuMoinsSignale = UnDefautAuMoinsSignale

    End With
    
    '*********************************************************************************************************************
    '*                                                                CUVE C27
    '*********************************************************************************************************************
    
    '--- affectation ---
    NumCuve = CUVES_REGULATION.C_C27
    ListeNumDefautsPourUneCuve = ""
    
    '--- recherche de la pr�sence d'une charge � ce poste ---
    NumChargeACePoste = TEtatsPostes(CorrespondanceCuvesAPIPostes(NumCuve)).NumCharge

    With TEtatsCuves(NumCuve)

        '--- initialisation � FALSE par d�faut ---
        UnDefautAuMoinsSignale = False

        If TEtatsLigne.MarcheGenerale = True Then

            '-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

            '--- niveau trop bas ---
            NumDefaut = .TNumDefauts.NumDefautNiveauTresBas
            EtatDefaut = .TEntreesAPI.E_NiveauTresBas
            UnDefautAuMoinsSignale = UnDefautAuMoinsSignale Or EtatDefaut
            ConstruitListeNumDefauts EtatDefaut, NumDefaut, ListeNumDefauts, , ListeNumDefautsPourUneCuve  'contruction des listes de d�fauts
            EnregistreDateDetectionDisparitionDefaut EtatDefaut, NumDefaut 'enregistrement de la date de d�tection et disparition du d�faut

            '--- niveau trop haut ---
            NumDefaut = .TNumDefauts.NumDefautNiveauTresHaut
            EtatDefaut = .TEntreesAPI.E_NiveauTresHaut
            UnDefautAuMoinsSignale = UnDefautAuMoinsSignale Or EtatDefaut
            ConstruitListeNumDefauts EtatDefaut, NumDefaut, ListeNumDefauts, , ListeNumDefautsPourUneCuve  'contruction des listes de d�fauts
            EnregistreDateDetectionDisparitionDefaut EtatDefaut, NumDefaut 'enregistrement de la date de d�tection et disparition du d�faut

            '-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

            '--- d�faut du chauffage ---
            NumDefaut = .TNumDefauts.NumDefautDefautChauffage
            EtatDefaut = .TEntreesAPI.E_DefautChauffage
            UnDefautAuMoinsSignale = UnDefautAuMoinsSignale Or EtatDefaut
            ConstruitListeNumDefauts EtatDefaut, NumDefaut, ListeNumDefauts, , ListeNumDefautsPourUneCuve  'contruction des listes de d�fauts
            EnregistreDateDetectionDisparitionDefaut EtatDefaut, NumDefaut 'enregistrement de la date de d�tection et disparition du d�faut
            
            '-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
            
            '--- d�faut PT100 ---
            NumDefaut = .TNumDefauts.NumDefautDefautPT100
            EtatDefaut = .TEntreesAPI.DefautPT100
            UnDefautAuMoinsSignale = UnDefautAuMoinsSignale Or EtatDefaut
            ConstruitListeNumDefauts EtatDefaut, NumDefaut, ListeNumDefauts, , ListeNumDefautsPourUneCuve  'contruction des listes de d�fauts
            EnregistreDateDetectionDisparitionDefaut EtatDefaut, NumDefaut 'enregistrement de la date de d�tection et disparition du d�faut
            
            '--- temp�rature trop basse ---
            NumDefaut = .TNumDefauts.NumDefautTemperatureTropBasse
            EtatDefaut = .TEntreesAPI.TemperatureTropBasse
            UnDefautAuMoinsSignale = UnDefautAuMoinsSignale Or EtatDefaut
            ConstruitListeNumDefauts EtatDefaut, NumDefaut, ListeNumDefauts, , ListeNumDefautsPourUneCuve  'contruction des listes de d�fauts
            EnregistreDateDetectionDisparitionDefaut EtatDefaut, NumDefaut 'enregistrement de la date de d�tection et disparition du d�faut
            
            '--- temp�rature trop haute ---
            NumDefaut = .TNumDefauts.NumDefautTemperatureTropHaute
            EtatDefaut = .TEntreesAPI.TemperatureTropHaute
            UnDefautAuMoinsSignale = UnDefautAuMoinsSignale Or EtatDefaut
            ConstruitListeNumDefauts EtatDefaut, NumDefaut, ListeNumDefauts, , ListeNumDefautsPourUneCuve  'contruction des listes de d�fauts
            EnregistreDateDetectionDisparitionDefaut EtatDefaut, NumDefaut 'enregistrement de la date de d�tection et disparition du d�faut

            '-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

        End If

        '--- construction de la liste des num�ros de d�fauts pour une cuve ayant une charge ---
        If NumChargeACePoste = 0 Then
            .ListeNumDefautsSiCharge = ""    'vider la liste si pas de charge dans le poste
        Else
            If ListeNumDefautsPourUneCuve <> "" Then  'remplir la liste en �vitant les doublons
                                                                                     'suppression du dernier s�parateur inutile car dans fonction
                .ListeNumDefautsSiCharge = AjoutNumDefautsSansDoublons(.ListeNumDefautsSiCharge, _
                                                                                                                     ListeNumDefautsPourUneCuve) 'remplir la liste en �liminant les doublons
            End If
        End If

        '--- affectation d�finitive ---
        .UnDefautAuMoinsSignale = UnDefautAuMoinsSignale

    End With
    
    '*********************************************************************************************************************
    '*                                                                CUVE C28
    '*********************************************************************************************************************
    
    '--- affectation ---
    NumCuve = CUVES_REGULATION.C_C28
    ListeNumDefautsPourUneCuve = ""
    
    '--- recherche de la pr�sence d'une charge � ce poste ---
    NumChargeACePoste = TEtatsPostes(CorrespondanceCuvesAPIPostes(NumCuve)).NumCharge

    With TEtatsCuves(NumCuve)

        '--- initialisation � FALSE par d�faut ---
        UnDefautAuMoinsSignale = False

        If TEtatsLigne.MarcheGenerale = True Then

            '-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

            '--- niveau trop bas ---
            NumDefaut = .TNumDefauts.NumDefautNiveauTresBas
            EtatDefaut = .TEntreesAPI.E_NiveauTresBas
            UnDefautAuMoinsSignale = UnDefautAuMoinsSignale Or EtatDefaut
            ConstruitListeNumDefauts EtatDefaut, NumDefaut, ListeNumDefauts, , ListeNumDefautsPourUneCuve  'contruction des listes de d�fauts
            EnregistreDateDetectionDisparitionDefaut EtatDefaut, NumDefaut 'enregistrement de la date de d�tection et disparition du d�faut

            '--- niveau trop haut ---
            NumDefaut = .TNumDefauts.NumDefautNiveauTresHaut
            EtatDefaut = .TEntreesAPI.E_NiveauTresHaut
            UnDefautAuMoinsSignale = UnDefautAuMoinsSignale Or EtatDefaut
            ConstruitListeNumDefauts EtatDefaut, NumDefaut, ListeNumDefauts, , ListeNumDefautsPourUneCuve  'contruction des listes de d�fauts
            EnregistreDateDetectionDisparitionDefaut EtatDefaut, NumDefaut 'enregistrement de la date de d�tection et disparition du d�faut

            '-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

            '--- d�faut du chauffage ---
            NumDefaut = .TNumDefauts.NumDefautDefautChauffage
            EtatDefaut = .TEntreesAPI.E_DefautChauffage
            UnDefautAuMoinsSignale = UnDefautAuMoinsSignale Or EtatDefaut
            ConstruitListeNumDefauts EtatDefaut, NumDefaut, ListeNumDefauts, , ListeNumDefautsPourUneCuve  'contruction des listes de d�fauts
            EnregistreDateDetectionDisparitionDefaut EtatDefaut, NumDefaut 'enregistrement de la date de d�tection et disparition du d�faut
            
            '-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
            
            '--- d�faut PT100 ---
            NumDefaut = .TNumDefauts.NumDefautDefautPT100
            EtatDefaut = .TEntreesAPI.DefautPT100
            UnDefautAuMoinsSignale = UnDefautAuMoinsSignale Or EtatDefaut
            ConstruitListeNumDefauts EtatDefaut, NumDefaut, ListeNumDefauts, , ListeNumDefautsPourUneCuve  'contruction des listes de d�fauts
            EnregistreDateDetectionDisparitionDefaut EtatDefaut, NumDefaut 'enregistrement de la date de d�tection et disparition du d�faut
            
            '--- temp�rature trop basse ---
            NumDefaut = .TNumDefauts.NumDefautTemperatureTropBasse
            EtatDefaut = .TEntreesAPI.TemperatureTropBasse
            UnDefautAuMoinsSignale = UnDefautAuMoinsSignale Or EtatDefaut
            ConstruitListeNumDefauts EtatDefaut, NumDefaut, ListeNumDefauts, , ListeNumDefautsPourUneCuve  'contruction des listes de d�fauts
            EnregistreDateDetectionDisparitionDefaut EtatDefaut, NumDefaut 'enregistrement de la date de d�tection et disparition du d�faut
            
            '--- temp�rature trop haute ---
            NumDefaut = .TNumDefauts.NumDefautTemperatureTropHaute
            EtatDefaut = .TEntreesAPI.TemperatureTropHaute
            UnDefautAuMoinsSignale = UnDefautAuMoinsSignale Or EtatDefaut
            ConstruitListeNumDefauts EtatDefaut, NumDefaut, ListeNumDefauts, , ListeNumDefautsPourUneCuve  'contruction des listes de d�fauts
            EnregistreDateDetectionDisparitionDefaut EtatDefaut, NumDefaut 'enregistrement de la date de d�tection et disparition du d�faut

            '-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

        End If

        '--- construction de la liste des num�ros de d�fauts pour une cuve ayant une charge ---
        If NumChargeACePoste = 0 Then
            .ListeNumDefautsSiCharge = ""    'vider la liste si pas de charge dans le poste
        Else
            If ListeNumDefautsPourUneCuve <> "" Then  'remplir la liste en �vitant les doublons
                                                                                     'suppression du dernier s�parateur inutile car dans fonction
                .ListeNumDefautsSiCharge = AjoutNumDefautsSansDoublons(.ListeNumDefautsSiCharge, _
                                                                                                                     ListeNumDefautsPourUneCuve) 'remplir la liste en �liminant les doublons
            End If
        End If

        '--- affectation d�finitive ---
        .UnDefautAuMoinsSignale = UnDefautAuMoinsSignale

    End With
    
    '*********************************************************************************************************************
    '*                                                                CUVE C31
    '*********************************************************************************************************************
    
    '--- affectation ---
    NumCuve = CUVES_REGULATION.C_C31
    ListeNumDefautsPourUneCuve = ""
    
    '--- recherche de la pr�sence d'une charge � ce poste ---
    NumChargeACePoste = TEtatsPostes(CorrespondanceCuvesAPIPostes(NumCuve)).NumCharge

    With TEtatsCuves(NumCuve)

        '--- initialisation � FALSE par d�faut ---
        UnDefautAuMoinsSignale = False

        If TEtatsLigne.MarcheGenerale = True Then

            '-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
            
            '--- niveau trop bas ---
            NumDefaut = .TNumDefauts.NumDefautNiveauTresBas
            EtatDefaut = .TEntreesAPI.E_NiveauTresBas
            UnDefautAuMoinsSignale = UnDefautAuMoinsSignale Or EtatDefaut
            ConstruitListeNumDefauts EtatDefaut, NumDefaut, ListeNumDefauts, , ListeNumDefautsPourUneCuve  'contruction des listes de d�fauts
            EnregistreDateDetectionDisparitionDefaut EtatDefaut, NumDefaut 'enregistrement de la date de d�tection et disparition du d�faut

            '--- niveau trop haut ---
            NumDefaut = .TNumDefauts.NumDefautNiveauTresHaut
            EtatDefaut = .TEntreesAPI.E_NiveauTresHaut
            UnDefautAuMoinsSignale = UnDefautAuMoinsSignale Or EtatDefaut
            ConstruitListeNumDefauts EtatDefaut, NumDefaut, ListeNumDefauts, , ListeNumDefautsPourUneCuve  'contruction des listes de d�fauts
            EnregistreDateDetectionDisparitionDefaut EtatDefaut, NumDefaut 'enregistrement de la date de d�tection et disparition du d�faut

            '-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

            '--- d�faut du chauffage ---
            NumDefaut = .TNumDefauts.NumDefautDefautChauffage
            EtatDefaut = .TEntreesAPI.E_DefautChauffage
            UnDefautAuMoinsSignale = UnDefautAuMoinsSignale Or EtatDefaut
            ConstruitListeNumDefauts EtatDefaut, NumDefaut, ListeNumDefauts, , ListeNumDefautsPourUneCuve  'contruction des listes de d�fauts
            EnregistreDateDetectionDisparitionDefaut EtatDefaut, NumDefaut 'enregistrement de la date de d�tection et disparition du d�faut
            
            '-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
            
            '--- d�faut PT100 ---
            NumDefaut = .TNumDefauts.NumDefautDefautPT100
            EtatDefaut = .TEntreesAPI.DefautPT100
            UnDefautAuMoinsSignale = UnDefautAuMoinsSignale Or EtatDefaut
            ConstruitListeNumDefauts EtatDefaut, NumDefaut, ListeNumDefauts, , ListeNumDefautsPourUneCuve  'contruction des listes de d�fauts
            EnregistreDateDetectionDisparitionDefaut EtatDefaut, NumDefaut 'enregistrement de la date de d�tection et disparition du d�faut
            
            '--- temp�rature trop basse ---
            NumDefaut = .TNumDefauts.NumDefautTemperatureTropBasse
            EtatDefaut = .TEntreesAPI.TemperatureTropBasse
            UnDefautAuMoinsSignale = UnDefautAuMoinsSignale Or EtatDefaut
            ConstruitListeNumDefauts EtatDefaut, NumDefaut, ListeNumDefauts, , ListeNumDefautsPourUneCuve  'contruction des listes de d�fauts
            EnregistreDateDetectionDisparitionDefaut EtatDefaut, NumDefaut 'enregistrement de la date de d�tection et disparition du d�faut
            
            '--- temp�rature trop haute ---
            NumDefaut = .TNumDefauts.NumDefautTemperatureTropHaute
            EtatDefaut = .TEntreesAPI.TemperatureTropHaute
            UnDefautAuMoinsSignale = UnDefautAuMoinsSignale Or EtatDefaut
            ConstruitListeNumDefauts EtatDefaut, NumDefaut, ListeNumDefauts, , ListeNumDefautsPourUneCuve  'contruction des listes de d�fauts
            EnregistreDateDetectionDisparitionDefaut EtatDefaut, NumDefaut 'enregistrement de la date de d�tection et disparition du d�faut

            '-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

        End If

        '--- construction de la liste des num�ros de d�fauts pour une cuve ayant une charge ---
        If NumChargeACePoste = 0 Then
            .ListeNumDefautsSiCharge = ""    'vider la liste si pas de charge dans le poste
        Else
            If ListeNumDefautsPourUneCuve <> "" Then  'remplir la liste en �vitant les doublons
                                                                                     'suppression du dernier s�parateur inutile car dans fonction
                .ListeNumDefautsSiCharge = AjoutNumDefautsSansDoublons(.ListeNumDefautsSiCharge, _
                                                                                                                     ListeNumDefautsPourUneCuve) 'remplir la liste en �liminant les doublons
            End If
        End If

        '--- affectation d�finitive ---
        .UnDefautAuMoinsSignale = UnDefautAuMoinsSignale

    End With
    
    '*********************************************************************************************************************
    '*                                                                CUVE C32
    '*********************************************************************************************************************
    
    '--- affectation ---
    NumCuve = CUVES_REGULATION.C_C32
    ListeNumDefautsPourUneCuve = ""
    
    '--- recherche de la pr�sence d'une charge � ce poste ---
    NumChargeACePoste = TEtatsPostes(CorrespondanceCuvesAPIPostes(NumCuve)).NumCharge

    With TEtatsCuves(NumCuve)

        '--- initialisation � FALSE par d�faut ---
        UnDefautAuMoinsSignale = False

        If TEtatsLigne.MarcheGenerale = True Then

            '-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

            '--- niveau trop bas ---
            NumDefaut = .TNumDefauts.NumDefautNiveauTresBas
            EtatDefaut = .TEntreesAPI.E_NiveauTresBas
            UnDefautAuMoinsSignale = UnDefautAuMoinsSignale Or EtatDefaut
            ConstruitListeNumDefauts EtatDefaut, NumDefaut, ListeNumDefauts, , ListeNumDefautsPourUneCuve  'contruction des listes de d�fauts
            EnregistreDateDetectionDisparitionDefaut EtatDefaut, NumDefaut 'enregistrement de la date de d�tection et disparition du d�faut

            '--- niveau trop haut ---
            NumDefaut = .TNumDefauts.NumDefautNiveauTresHaut
            EtatDefaut = .TEntreesAPI.E_NiveauTresHaut
            UnDefautAuMoinsSignale = UnDefautAuMoinsSignale Or EtatDefaut
            ConstruitListeNumDefauts EtatDefaut, NumDefaut, ListeNumDefauts, , ListeNumDefautsPourUneCuve  'contruction des listes de d�fauts
            EnregistreDateDetectionDisparitionDefaut EtatDefaut, NumDefaut 'enregistrement de la date de d�tection et disparition du d�faut

            '-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

            '--- d�faut du chauffage ---
            NumDefaut = .TNumDefauts.NumDefautDefautChauffage
            EtatDefaut = .TEntreesAPI.E_DefautChauffage
            UnDefautAuMoinsSignale = UnDefautAuMoinsSignale Or EtatDefaut
            ConstruitListeNumDefauts EtatDefaut, NumDefaut, ListeNumDefauts, , ListeNumDefautsPourUneCuve  'contruction des listes de d�fauts
            EnregistreDateDetectionDisparitionDefaut EtatDefaut, NumDefaut 'enregistrement de la date de d�tection et disparition du d�faut
            
            '-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
            
            '--- d�faut PT100 ---
            NumDefaut = .TNumDefauts.NumDefautDefautPT100
            EtatDefaut = .TEntreesAPI.DefautPT100
            UnDefautAuMoinsSignale = UnDefautAuMoinsSignale Or EtatDefaut
            ConstruitListeNumDefauts EtatDefaut, NumDefaut, ListeNumDefauts, , ListeNumDefautsPourUneCuve  'contruction des listes de d�fauts
            EnregistreDateDetectionDisparitionDefaut EtatDefaut, NumDefaut 'enregistrement de la date de d�tection et disparition du d�faut
            
            '--- temp�rature trop basse ---
            NumDefaut = .TNumDefauts.NumDefautTemperatureTropBasse
            EtatDefaut = .TEntreesAPI.TemperatureTropBasse
            UnDefautAuMoinsSignale = UnDefautAuMoinsSignale Or EtatDefaut
            ConstruitListeNumDefauts EtatDefaut, NumDefaut, ListeNumDefauts, , ListeNumDefautsPourUneCuve  'contruction des listes de d�fauts
            EnregistreDateDetectionDisparitionDefaut EtatDefaut, NumDefaut 'enregistrement de la date de d�tection et disparition du d�faut
            
            '--- temp�rature trop haute ---
            NumDefaut = .TNumDefauts.NumDefautTemperatureTropHaute
            EtatDefaut = .TEntreesAPI.TemperatureTropHaute
            UnDefautAuMoinsSignale = UnDefautAuMoinsSignale Or EtatDefaut
            ConstruitListeNumDefauts EtatDefaut, NumDefaut, ListeNumDefauts, , ListeNumDefautsPourUneCuve  'contruction des listes de d�fauts
            EnregistreDateDetectionDisparitionDefaut EtatDefaut, NumDefaut 'enregistrement de la date de d�tection et disparition du d�faut

            '-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

        End If

        '--- construction de la liste des num�ros de d�fauts pour une cuve ayant une charge ---
        If NumChargeACePoste = 0 Then
            .ListeNumDefautsSiCharge = ""    'vider la liste si pas de charge dans le poste
        Else
            If ListeNumDefautsPourUneCuve <> "" Then  'remplir la liste en �vitant les doublons
                                                                                     'suppression du dernier s�parateur inutile car dans fonction
                .ListeNumDefautsSiCharge = AjoutNumDefautsSansDoublons(.ListeNumDefautsSiCharge, _
                                                                                                                     ListeNumDefautsPourUneCuve) 'remplir la liste en �liminant les doublons
            End If
        End If

        '--- affectation d�finitive ---
        .UnDefautAuMoinsSignale = UnDefautAuMoinsSignale

    End With

    '*********************************************************************************************************************
    '*                                                                  REDRESSEURS
    '*********************************************************************************************************************
    
    For a = LBound(TEtatsRedresseurs()) To UBound(TEtatsRedresseurs())
        
        With TEtatsRedresseurs(a)

            '--- d�faut g�n�ral ---
            NumDefaut = .TNumDefauts.NumDefautDefautGeneral
            EtatDefaut = .TEntreesAPI.M_DefautGeneral
            UnDefautAuMoinsSignale = UnDefautAuMoinsSignale Or EtatDefaut
            ConstruitListeNumDefauts EtatDefaut, NumDefaut, ListeNumDefauts, , ListeNumDefautsPourUneCuve  'contruction des listes de d�fauts
            EnregistreDateDetectionDisparitionDefaut EtatDefaut, NumDefaut 'enregistrement de la date de d�tection et disparition du d�faut
            
            '--- d�lai trop long de mise en marche ---
            NumDefaut = .TNumDefauts.NumDefautDelaiTropLongMarcheRedresseur
            EtatDefaut = .TEntreesAPI.M_DelaiTropLongMarcheRedresseur
            UnDefautAuMoinsSignale = UnDefautAuMoinsSignale Or EtatDefaut
            ConstruitListeNumDefauts EtatDefaut, NumDefaut, ListeNumDefauts, , ListeNumDefautsPourUneCuve  'contruction des listes de d�fauts
            EnregistreDateDetectionDisparitionDefaut EtatDefaut, NumDefaut 'enregistrement de la date de d�tection et disparition du d�faut
                            
            '--- intensit� non atteinte ---
            NumDefaut = .TNumDefauts.NumDefautIntensiteNonAtteinte
            EtatDefaut = .TEntreesAPI.M_IntensiteNonAtteinte
            UnDefautAuMoinsSignale = UnDefautAuMoinsSignale Or EtatDefaut
            ConstruitListeNumDefauts EtatDefaut, NumDefaut, ListeNumDefauts, , ListeNumDefautsPourUneCuve  'contruction des listes de d�fauts
            EnregistreDateDetectionDisparitionDefaut EtatDefaut, NumDefaut 'enregistrement de la date de d�tection et disparition du d�faut
                            
            '--- intensit� instable ---
            NumDefaut = .TNumDefauts.NumDefautIntensiteInstable
            EtatDefaut = .TEntreesAPI.M_IntensiteInstable
            UnDefautAuMoinsSignale = UnDefautAuMoinsSignale Or EtatDefaut
            ConstruitListeNumDefauts EtatDefaut, NumDefaut, ListeNumDefauts, , ListeNumDefautsPourUneCuve  'contruction des listes de d�fauts
            EnregistreDateDetectionDisparitionDefaut EtatDefaut, NumDefaut 'enregistrement de la date de d�tection et disparition du d�faut
    
        End With
    
    Next a
    
    '*********************************************************************************************************************
    '*                                                ANALYSE DES ALARMES LIGNE  EN COURS
    '*********************************************************************************************************************

    '--- affectation ---
    If ListeNumDefautsPourLaLigne = "" Then
        AlarmesLigneEnCours = ""
    Else
        AlarmesLigneEnCours = Left(ListeNumDefautsPourLaLigne, Pred(Len(ListeNumDefautsPourLaLigne))) 'suppression du dernier s�parateur
    End If

    '*********************************************************************************************************************
    '*                                       DECODAGE DE LA LISTE DES NUMEROS DE DEFAUTS
    '*********************************************************************************************************************

    If ListeNumDefauts <> "" Then

        '--- passage en couleur rouge ---
        If OccFPrincipale.LMessages.BackColor <> ROUGE_DEFAUT Then
            OccFPrincipale.LMessages.BackColor = ROUGE_DEFAUT
            OccFPrincipale.LMessages.ForeColor = COULEURS.JAUNE_3
            CptAppels = 0
        End If

        If CptAppels = 0 Then

            '--- construction du tableau contenant les num�ros de d�fauts ---
            TNumDefauts = Split(ListeNumDefauts, SEPARATEUR_NUM_DEFAUTS)
            If CptDefauts >= UBound(TNumDefauts) Then 'ATTENTION la derni�re valeur du tableau est
                                                                                      'automatiquement une chaine vide � cause
                                                                                      'du dernier s�parateur ajout�
                CptDefauts = 0
                CptAppels = 0
            End If

            '--- affichage du d�faut ---
            If IsNumeric(TNumDefauts(CptDefauts)) = True Then
                
                NumDefautAAfficher = CLng(TNumDefauts(CptDefauts))
                
                If NumDefautAAfficher >= DEFAUTS.NUM_MINI And NumDefautAAfficher <= DEFAUTS.NUM_MAXI Then
                    
                    '--- affectation du libell� compl�t� du d�faut (cas des num�ros de d�faut des variateurs) ---
                    LibelleCompleteDefaut = CompleteLibelleDefaut(NumDefautAAfficher, ComplementDefaut)
                    
                    '--- r�affectation du libell� complet ---
                    LibelleCompleteDefaut = "D�faut n� " & NumDefautAAfficher & " - " & LibelleCompleteDefaut
                    
                    '--- afffichage dans le champ concern� (�cran principal) ---
                    If OccFPrincipale.LMessages.Caption <> LibelleCompleteDefaut Then
                        OccFPrincipale.LMessages.Caption = LibelleCompleteDefaut
                    End If
                    
                    '--- affichage sur l'afficheur � condition qu'il n'y ai pas de priorit� pour les alertes---
                    If TDefauts(NumDefautAAfficher).AfficheurOuiNon = True And PrioriteAfficheurPourAlertes = False Then
                        Bidon = MessageAfficheur("B", TDefauts(NumDefautAAfficher).LibelleDefautAfficheur)
                    End If
                    
                End If
            
            End If
            Inc CptDefauts

        End If

    Else

        '--- passage en couleur verte / effacement du message sur l'afficheur � leds rouge ---
        If OccFPrincipale.LMessages.BackColor <> COULEURS.VERT_3 Then
            
            '--- passage en couleur verte ---
            OccFPrincipale.LMessages.BackColor = COULEURS.VERT_3
            OccFPrincipale.LMessages.ForeColor = COULEURS.NOIR
        
            '--- effacement du message sur l'afficheur � leds rouge ---
            Bidon = MessageAfficheur("B", "")
        
        End If

        '--- affectation ---
        CptAppels = 0
        CptDefauts = 0

    End If

    '--- contr�le du rafraichissement  ---
    Inc CptAppels
    If CptAppels > 5 Then CptAppels = 0

End Sub



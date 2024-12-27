Attribute VB_Name = "MAlarmes"
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Rôle                    : MODULE AIDANT A LA GESTION DES ALARMES
' Nom                    : MAlarmes.bas
' Date de création : 26/06/2002
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

'--- déclarations obligatoires ---
Option Explicit

'--- options générales ---
Option Base 1
DefVar A-Z

'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Rôle      : Acquitte les alarmes (coupe le gyrophare et le klaxon)
' Entrées :
' Retours :
' Détails  :
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Sub AcquittementAlarmes()

    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next

    '--- déclaration ---

    '--- analyse en fonction du PC ---
    'If TypePC <> TYPES_PC. Then Exit Sub
    
    '--- écriture dans l'API ---
    If PROGRAMME_AVEC_AUTOMATE = True Then
        Bidon = APIEcritureVariableNommee("DEFAUTS", "M_Acquittement_Defauts", True)                   'pour le gyrophare et le klaxon
    End If

End Sub

'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Rôle      : Attribution d'un numéro de défaut en fonction de chaque défaut
' Entrées :
' Retours :
' Détails  :
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function AttributionNumDefauts() As String
    
    '--- aiguillage en cas d'erreurs ---
    On Error GoTo GestionErreurs
    
    '--- déclaration ---
    Dim a As Integer
            
    '--- affectation ---
    AttributionNumDefauts = ""
    
    '--- affichage du type de tâche ---
    AfficheTypeTache ("Attribution des numéros de défauts")
    
    '*********************************************************************************************************************
    '*                                                                       ETATS DE LA LIGNE
    '*********************************************************************************************************************
    
    With TEtatsLigne.TNumDefauts
        
        '--- arrêt générale ---
        .NumDefautArretGeneral = 1
        
        '--- arrêt d'urgence ---
        .NumDefautArretUrgence = 2
        
        '--- portillons et ligne de vie ---
        .NumDefautPortillonsLigneVie = 3
        
        '--- sécurité du pont 1 ---
        .NumDefautSecuriteP1 = 4
        
        '--- sécurité du pont 2 ---
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
            
            '--- défaut du variateur de la translation ---
            .NumDefautDefautVariateurTrLPont = Choose(a, 101, 201)
                
            '--- défaut du variateur du levage ---
            .NumDefautDefautVariateurLevPont = Choose(a, 104, 204)
                
            '--- surcourse levage haut ---
            .NumDefautSurcourseLevHaut = Choose(a, 105, 205)
            
            '--- surcourse levage bas ---
            .NumDefautSurcourseLevBas = Choose(a, 106, 206)
                
            '--- axe non référencé de la translation ---
            .NumDefautAxeNonReferenceTrlPont = Choose(a, 110, 210)
            
            '--- axe non référencé du levage ---
            .NumDefautAxeNonReferenceLevPont = Choose(a, 111, 211)
                
            '--- anti-collision ---
            .NumDefautAntiCollision = Choose(a, 109, 209)
                
            '--- fin de zone ---
            .NumDefautFinDeZone = Choose(a, 110, 210)
                
            '--- présence pièce ---
            .NumDefautPresencePiece = Choose(a, 115, 215)
                
            '--- défaut laser ---
            .NumDefautDefautLaser = Choose(a, 116, 216)
        
            '--- délai trop long de descente des accroches de la charge ---
            .NumDefautDelaiTropLongDescenteAccroches = Choose(a, 120, 220)
            
            '--- délai trop long de montée des accroches de la charge ---
            .NumDefautDelaiTropLongMonteeAccroches = Choose(a, 121, 221)
        
        End With
    
    Next a
    
    '*********************************************************************************************************************
    '*                                                                        CUVE C00
    '*********************************************************************************************************************
    
    With TEtatsCuves(CUVES_REGULATION.C_C00).TNumDefauts

        '--- niveau très bas ---
        .NumDefautNiveauTresBas = 300
        
        '--- niveau très haut ---
        .NumDefautNiveauTresHaut = 301

        '--- défaut chauffage ---
        .NumDefautDefautChauffage = 302

        '--- défauts températures ---
        .NumDefautDefautPT100 = 303
        .NumDefautTemperatureTropBasse = 304
        .NumDefautTemperatureTropHaute = 305
        
        
      
'        '--- disjonction des électro-vannes ouverture/fermeture des couvercles ---
'        .NumDefautDisjonctionEVCouvercles = 301
'
'        '--- disjonction chauffage ---
        '.NumDefautDisjonctionChauffage = 302
'
'        '--- disjonction électro-vanne d'arrivée d'eau ---
'        .NumDefautDisjonctionEVEau = 304
'
'        '--- disjonction de la pompe ---
'        .NumDefautDisjonctionPompe = 305
'
'
'        '--- délai trop long d'ouverture des couvercles ---
'        .NumDefautDelaiTropLongOuvertureCouvercles = 312
'
'        '--- délai trop long de fermeture des couvercles ---
'        .NumDefautDelaiTropLongFermetureCouvercles = 313
'
'        '--- délai trop long de l'électro-vanne d'arrivée d'eau ---
'        .NumDefautDelaiTropLongEVEau = 314
'
    End With
    
    '*********************************************************************************************************************
    '*                                                                        CUVE C01
    '*********************************************************************************************************************
    If False Then
          'With TEtatsCuves(CUVES_REGULATION.C_SAT).TNumDefauts
           With TEtatsCuves(1).TNumDefauts
        
            '--- niveau très bas ---
            .NumDefautNiveauTresBas = 325
            
            '--- niveau très haut ---
            .NumDefautNiveauTresHaut = 326
    
            '--- défaut chauffage ---
            .NumDefautDefautChauffage = 327
    
            '--- défauts températures ---
            .NumDefautDefautPT100 = 328
            .NumDefautTemperatureTropBasse = 329
            .NumDefautTemperatureTropHaute = 330
    
            '        '--- disjonction des électro-vannes ouverture/fermeture des couvercles ---
            '        .NumDefautDisjonctionEVCouvercles = 331
            '
            '        '--- disjonction chauffage ---
            '        .NumDefautDisjonctionChauffage = 332
            '
            '        '--- disjonction électro-vanne d'arrivée d'eau ---
            '        .NumDefautDisjonctionEVEau = 334
            '
            '        '--- délai trop long d'ouverture des couvercles ---
            '        .NumDefautDelaiTropLongOuvertureCouvercles = 341
            '
            '        '--- délai trop long de fermeture des couvercles ---
            '        .NumDefautDelaiTropLongFermetureCouvercles = 342
            '
            '        '--- délai trop long de l'électro-vanne d'arrivée d'eau ---
            '        .NumDefautDelaiTropLongEVEau = 343
            '
        End With
    
    
    End If
  
        
    
 
   
    
    '*********************************************************************************************************************
    '*                                                                        CUVE C02
    '*********************************************************************************************************************
    
    With TEtatsCuves(CUVES_REGULATION.C_DEC).TNumDefauts

        '--- niveau très bas ---
        .NumDefautNiveauTresBas = 350
        
        '--- niveau très haut ---
        .NumDefautNiveauTresHaut = 351

        '--- défaut chauffage ---
        .NumDefautDefautChauffage = 352
        
        '--- défauts températures ---
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
         
        '--- niveau très bas ---
        .NumDefautNiveauTresBas = 375
        
        '--- niveau très haut ---
        .NumDefautNiveauTresHaut = 376
    
        '--- défaut chauffage ---
        .NumDefautDefautChauffage = 377
    
        '--- défauts températures ---
        .NumDefautDefautPT100 = 378
        .NumDefautTemperatureTropBasse = 379
        .NumDefautTemperatureTropHaute = 380
    
        End With
      
        
        '*********************************************************************************************************************
        '*                                                                        CUVE C05
        '*********************************************************************************************************************
        
        'With TEtatsCuves(CUVES_REGULATION.C_C05).TNumDefauts
         With TEtatsCuves(1).TNumDefauts
    
            '--- niveau très bas ---
            .NumDefautNiveauTresBas = 400
            
            '--- niveau très haut ---
            .NumDefautNiveauTresHaut = 401
    
            '--- défaut chauffage ---
            .NumDefautDefautChauffage = 402
            
            '--- défauts températures ---
            .NumDefautDefautPT100 = 403
            .NumDefautTemperatureTropBasse = 404
            .NumDefautTemperatureTropHaute = 405
        
        End With
        
        '*********************************************************************************************************************
        '*                                                                        CUVE C06
        '*********************************************************************************************************************
        
        'With TEtatsCuves(CUVES_REGULATION.C_C06).TNumDefauts
     With TEtatsCuves(1).TNumDefauts
            '--- niveau très bas ---
            .NumDefautNiveauTresBas = 425
            
            '--- niveau très haut ---
            .NumDefautNiveauTresHaut = 426
    
            '--- défaut chauffage ---
            .NumDefautDefautChauffage = 427
            
            '--- défauts températures ---
            .NumDefautDefautPT100 = 428
            .NumDefautTemperatureTropBasse = 429
            .NumDefautTemperatureTropHaute = 430
        
        End With
    
    End If
    
    '*********************************************************************************************************************
    '*                                                                        CUVE C07
    '*********************************************************************************************************************
    
    With TEtatsCuves(CUVES_REGULATION.C_C07).TNumDefauts

        '--- niveau très bas ---
        .NumDefautNiveauTresBas = 450
        
        '--- niveau très haut ---
        .NumDefautNiveauTresHaut = 451

        '--- défaut chauffage ---
        .NumDefautDefautChauffage = 452
    
        '--- défauts températures ---
        .NumDefautDefautPT100 = 453
        .NumDefautTemperatureTropBasse = 454
        .NumDefautTemperatureTropHaute = 455
    
    End With
    
    '*********************************************************************************************************************
    '*                                                                        CUVE C13
    '*********************************************************************************************************************
    
    With TEtatsCuves(CUVES_REGULATION.C_C13).TNumDefauts

        '--- niveau très bas ---
        .NumDefautNiveauTresBas = 475
        
        '--- niveau très haut ---
        .NumDefautNiveauTresHaut = 476

        '--- défaut chauffage ---
        .NumDefautDefautChauffage = 477
    
        '--- défauts températures ---
        .NumDefautDefautPT100 = 478
        .NumDefautTemperatureTropBasse = 479
        .NumDefautTemperatureTropHaute = 480
    
        '--- défaut refroidissement ---
        .NumDefautDefautRefroidissement = 481
    
    End With
    
    '*********************************************************************************************************************
    '*                                                                        CUVE C14
    '*********************************************************************************************************************
    
    With TEtatsCuves(CUVES_REGULATION.C_C14).TNumDefauts

        '--- niveau très bas ---
        .NumDefautNiveauTresBas = 500
        
        '--- niveau très haut ---
        .NumDefautNiveauTresHaut = 501

        '--- défaut chauffage ---
        .NumDefautDefautChauffage = 502
    
        '--- défauts températures ---
        .NumDefautDefautPT100 = 503
        .NumDefautTemperatureTropBasse = 504
        .NumDefautTemperatureTropHaute = 505
        
        '--- défaut refroidissement ---
        .NumDefautDefautRefroidissement = 506
    
    End With
    
    '*********************************************************************************************************************
    '*                                                                        CUVE C15
    '*********************************************************************************************************************
    
    With TEtatsCuves(CUVES_REGULATION.C_C15).TNumDefauts

        '--- niveau très bas ---
        .NumDefautNiveauTresBas = 525
        
        '--- niveau très haut ---
        .NumDefautNiveauTresHaut = 526

        '--- défaut chauffage ---
        .NumDefautDefautChauffage = 527
        
        '--- défauts températures ---
        .NumDefautDefautPT100 = 528
        .NumDefautTemperatureTropBasse = 529
        .NumDefautTemperatureTropHaute = 530
        
        '--- défaut refroidissement ---
        .NumDefautDefautRefroidissement = 531
    
    End With
    
    '*********************************************************************************************************************
    '*                                                                        CUVE C16
    '*********************************************************************************************************************
    If False Then
       'With TEtatsCuves(CUVES_REGULATION.C_C16).TNumDefauts
       With TEtatsCuves(1).TNumDefauts


        '--- niveau très bas ---
        .NumDefautNiveauTresBas = 550
        
        '--- niveau très haut ---
        .NumDefautNiveauTresHaut = 551
    
        '--- défaut chauffage ---
        .NumDefautDefautChauffage = 552
    
        '--- défauts températures ---
        .NumDefautDefautPT100 = 553
        .NumDefautTemperatureTropBasse = 554
        .NumDefautTemperatureTropHaute = 555
    
        '--- défaut refroidissement ---
        .NumDefautDefautRefroidissement = 556
    
        End With

    
        
        
        '*********************************************************************************************************************
        '*                                                                        CUVE C19
        '*********************************************************************************************************************
        'With TEtatsCuves(CUVES_REGULATION.C_C19).TNumDefauts
        With TEtatsCuves(1).TNumDefauts
            '--- niveau très bas ---
            .NumDefautNiveauTresBas = 575
            
            '--- niveau très haut ---
            .NumDefautNiveauTresHaut = 576
    
            '--- défaut chauffage ---
            .NumDefautDefautChauffage = 577
        
            '--- défauts températures ---
            .NumDefautDefautPT100 = 578
            .NumDefautTemperatureTropBasse = 579
            .NumDefautTemperatureTropHaute = 580
        
        End With
    End If
    '*********************************************************************************************************************
    '*                                                                        CUVE C22
    '*********************************************************************************************************************
    
    With TEtatsCuves(CUVES_REGULATION.C_C22).TNumDefauts

        '--- niveau très bas ---
        .NumDefautNiveauTresBas = 600
        
        '--- niveau très haut ---
        .NumDefautNiveauTresHaut = 601

        '--- défaut chauffage ---
        .NumDefautDefautChauffage = 602
        
        '--- défauts températures ---
        .NumDefautDefautPT100 = 603
        .NumDefautTemperatureTropBasse = 604
        .NumDefautTemperatureTropHaute = 605
    
    End With
    
    '*********************************************************************************************************************
    '*                                                                        CUVE C27
    '*********************************************************************************************************************
    
    With TEtatsCuves(CUVES_REGULATION.C_C27).TNumDefauts

        '--- niveau très bas ---
        .NumDefautNiveauTresBas = 625
        
        '--- niveau très haut ---
        .NumDefautNiveauTresHaut = 626

        '--- défaut chauffage ---
        .NumDefautDefautChauffage = 627
    
        '--- défauts températures ---
        .NumDefautDefautPT100 = 628
        .NumDefautTemperatureTropBasse = 629
        .NumDefautTemperatureTropHaute = 630
    
    End With
    
    '*********************************************************************************************************************
    '*                                                                        CUVE C28
    '*********************************************************************************************************************
    
    With TEtatsCuves(CUVES_REGULATION.C_C28).TNumDefauts

        '--- niveau très bas ---
        .NumDefautNiveauTresBas = 650
        
        '--- niveau très haut ---
        .NumDefautNiveauTresHaut = 651

        '--- défaut chauffage ---
        .NumDefautDefautChauffage = 652
        
        '--- défauts températures ---
        .NumDefautDefautPT100 = 653
        .NumDefautTemperatureTropBasse = 654
        .NumDefautTemperatureTropHaute = 655
    
    End With
    
    '*********************************************************************************************************************
    '*                                                                        CUVE C31
    '*********************************************************************************************************************
    
    With TEtatsCuves(CUVES_REGULATION.C_C31).TNumDefauts

        '--- niveau très bas ---
        .NumDefautNiveauTresBas = 675
        
        '--- niveau très haut ---
        .NumDefautNiveauTresHaut = 676

        '--- défaut chauffage ---
        .NumDefautDefautChauffage = 677
    
        '--- défauts températures ---
        .NumDefautDefautPT100 = 678
        .NumDefautTemperatureTropBasse = 679
        .NumDefautTemperatureTropHaute = 680
    
    End With
    
    '*********************************************************************************************************************
    '*                                                                        CUVE C32
    '*********************************************************************************************************************
    
    With TEtatsCuves(CUVES_REGULATION.C_C32).TNumDefauts

        '--- niveau très bas ---
        .NumDefautNiveauTresBas = 700
        
        '--- niveau très haut ---
        .NumDefautNiveauTresHaut = 701

        '--- défaut chauffage ---
        .NumDefautDefautChauffage = 702
        
        '--- défauts températures ---
        .NumDefautDefautPT100 = 702
        .NumDefautTemperatureTropBasse = 703
        .NumDefautTemperatureTropHaute = 704
    
    End With
    
    '*********************************************************************************************************************
    '*                                                                    CUVE C33 /C34
    '*********************************************************************************************************************
    
   'With TEtatsCuves(CUVES_REGULATION.C_C33).TNumDefauts
  '
        '--- défaut chauffage ---
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
    
    '--- affichage du type de tâche ---
    AfficheTypeTache ("")

    Exit Function

GestionErreurs:

    AttributionNumDefauts = CStr(Err.Number)

End Function


'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Rôle      : Complète le libellé d'un défaut en fonction de son numéro
'                 (cas des numéros d'erreurs dans les variateurs)
' Entrées :                    NumDefaut -> Représente un numéro de défaut
' Retours :       ComplementDefaut -> Contient le texte du complément ajouté au libellé du défaut
'                 CompleteLibelleDefaut -> Libellé du défaut complété
' Détails  :
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function CompleteLibelleDefaut(ByVal NumDefaut As Integer, _
                                                               ByRef ComplementDefaut As String) As String
                                                              
    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next

    '--- déclaration ---
    Dim LibelleDefaut As String                     'libellé d'un défaut
    
    '--- affectation par défaut du libellé du défaut ---
    If NumDefaut >= DEFAUTS.NUM_MINI And NumDefaut <= DEFAUTS.NUM_MAXI Then
        LibelleDefaut = TDefauts(NumDefaut).LibelleDefaut
    End If

    '--- complément du libellé si nécessaire ---
    Select Case NumDefaut

        Case 100
            '--- translation gauche du pont ---
            'ComplementDefaut = "Défaut variateur n° " & TEntreesSortiesVariateurs6DP(VARIATEURS_6DP.V_TRL_G_PONT).NumDefaut
        
        Case 101
            '--- translation droite du pont ---
            'ComplementDefaut = "Défaut variateur n° " & TEntreesSortiesVariateurs6DP(VARIATEURS_6DP.V_TRL_D_PONT).NumDefaut
        
        Case 102
            '--- levage du pont ---
            'ComplementDefaut = "Défaut variateur n° " & TEntreesSortiesVariateurs6DP(VARIATEURS_6DP.V_LEV_PONT).NumDefaut

        Case 103
            '--- défaut du variateur de la translation gauche du dégraissage ---
            'ComplementDefaut = "Défaut variateur n° " & TEntreesSortiesVariateurs3DP(VARIATEURS_3DP.V_TRL_G_DEGRAISSAGE).NumDefaut

        Case 104
            '--- défaut du variateur de la translation droite du dégraissage ---
            'ComplementDefaut = "Défaut variateur n° " & TEntreesSortiesVariateurs3DP(VARIATEURS_3DP.V_TRL_D_DEGRAISSAGE).NumDefaut
        
        Case 105
            '--- défaut du variateur rotation brosses avant ---
            'ComplementDefaut = "Défaut variateur n° " & TEntreesSortiesVariateurs3DP(VARIATEURS_3DP.V_ROT_BROSSES_AVANT).NumDefaut
        
        Case 106
            '--- défaut du variateur rotation brosses arrière ---
            'ComplementDefaut = "Défaut variateur n° " & TEntreesSortiesVariateurs3DP(VARIATEURS_3DP.V_ROT_BROSSES_ARRIERE).NumDefaut

        Case Else
            '--- valeur par défaut du complément du défaut ---
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
' Rôle      : Construit une liste des numéros de défauts en fonction de l'état d'un défaut
' Entrées :                                     EtatDefaut -> Etat d'un défaut FALSE = pas de défaut en cours
'                                                                                                      TRUE = défaut en cours
'                                                   NumDefaut -> Représente un numéro de défaut
' Retours :                         ListeNumDefauts -> Liste de tous les numéros de défauts
'                    ListeNumDefautsPourLaLigne -> Liste des numéros de défauts concernant
'                                                                          uniquement la ligne (sans les cuves)
'                  ListeNumDefautsPourUneCuve -> Liste des numéros de défauts concernant
'                                                                          uniquement les cuves
' Détails  :
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Sub ConstruitListeNumDefauts(ByVal EtatDefaut As Boolean, _
                                                              ByVal NumDefaut As Integer, _
                                                              ByRef ListeNumDefauts As String, _
                                                              Optional ByRef ListeNumDefautsPourLaLigne As String, _
                                                              Optional ByRef ListeNumDefautsPourUneCuve As String)
    
    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next

    '--- déclaration ---
    
    If NumDefaut > 0 Then
    
        If EtatDefaut = True And TDefauts(NumDefaut).SignalerOuiNon = True Then
            
            '--- affectation de tous les numéros de défaut ---
            ListeNumDefauts = ListeNumDefauts & CStr(NumDefaut) & SEPARATEUR_NUM_DEFAUTS
            
            '--- affectation des numéros de défauts uniquement pour la ligne ---
            If IsMissing(ListeNumDefautsPourLaLigne) = False Then
                ListeNumDefautsPourLaLigne = ListeNumDefautsPourLaLigne & CStr(NumDefaut) & SEPARATEUR_NUM_DEFAUTS
            End If
            
            '--- affectation des numéros de défauts uniquement pour une cuve ---
            If IsMissing(ListeNumDefautsPourUneCuve) = False Then
                ListeNumDefautsPourUneCuve = ListeNumDefautsPourUneCuve & CStr(NumDefaut) & SEPARATEUR_NUM_DEFAUTS
            End If
    
        End If

    End If

End Sub

'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Rôle      : Enregistre la date de détection et de disparition d'un défaut en fonction de son état
' Entrées :   EtatDefaut -> Etat d'un défaut FALSE = pas de défaut en cours
'                                                                    TRUE = défaut en cours
'                 NumDefaut -> Représente un numéro de défaut
' Retours :
' Détails  :
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Sub EnregistreDateDetectionDisparitionDefaut(ByVal EtatDefaut As Boolean, _
                                                                                       ByVal NumDefaut As Integer)
    
    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next

    '--- déclaration ---
        
    '--- enregistrement de la date de détection du défaut et la date de sa disparition ---
    If EtatDefaut = True Then
        SignalerDefaut NumDefaut, True
    Else
        SignalerDefaut NumDefaut, False
        InitialiserAntiRebondsDefaut NumDefaut
    End If
        
End Sub

'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Rôle      : Ajoute des numéros de défauts à une liste sans que le resultat contienne des doublons
' Entrées :            ListeNumDefautsDeBase -> Liste des numéros de défauts de base
'                           ListeNumDefautsAAjouter -> Liste des numéros de défauts à ajouter à la liste de base
' Retours : AjoutNumDefautsSansDoublons -> Contient la l'addition des 2 listes sans doublons
' Détails  :
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function AjoutNumDefautsSansDoublons(ByVal ListeNumDefautsDeBase As String, _
                                                                               ByVal ListeNumDefautsAAjouter As String) As String
    
    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next
    
    '--- déclaration ---
    Dim a As Integer
    Dim NumDefaut As String, _
            CopieListeNumDefautsDeBase  As String, _
            CopieListeNumDefautsAAjouter As String
    Dim TListeNumDefautsAAjouter As Variant
    Dim TListeNumDefautsASupprimer As Variant
           
    If ListeNumDefautsAAjouter <> "" Then
    
        If ListeNumDefautsDeBase = "" Then
    
            '--- si la liste des n° de défauts de base est vide alors affecté directement la liste des ajouts ---
            ListeNumDefautsDeBase = ListeNumDefautsAAjouter
        
        Else
    
            '--- avant toutes opérations il faut effectuer un contrôle sur les chaines afin d'éviter que le
            '    dernier caractère soit un séparateur de n° de défauts ---
            If Right(ListeNumDefautsDeBase, 1) = SEPARATEUR_NUM_DEFAUTS Then
                ListeNumDefautsDeBase = Left(ListeNumDefautsDeBase, Pred(Len(ListeNumDefautsDeBase))) 'suppression du dernier séparateur
            End If
            If Right(ListeNumDefautsAAjouter, 1) = SEPARATEUR_NUM_DEFAUTS Then
                ListeNumDefautsAAjouter = Left(ListeNumDefautsAAjouter, Pred(Len(ListeNumDefautsAAjouter))) 'suppression du dernier séparateur
            End If
            
            '--- construction du tableau des numéros de défauts à ajouter ---
            TListeNumDefautsAAjouter = Split(ListeNumDefautsAAjouter, SEPARATEUR_NUM_DEFAUTS)
        
            '--- recherche des données dans le tableau ---
            For a = LBound(TListeNumDefautsAAjouter) To UBound(TListeNumDefautsAAjouter)
        
                '--- affectation sur la liste de base afin de pouvoir détecter
                    'un numéro de défaut avec 2 séparateurs (ex : -123-,-41-,-523-) pour éviter une mauvaise comparaison ---
                CopieListeNumDefautsDeBase = SEPARATEUR_NUM_DEFAUTS & ListeNumDefautsDeBase & SEPARATEUR_NUM_DEFAUTS
        
                '--- affectation ---
                NumDefaut = TListeNumDefautsAAjouter(a)
        
                '--- recherche du défaut ---
                If InStr(CopieListeNumDefautsDeBase, SEPARATEUR_NUM_DEFAUTS & NumDefaut & SEPARATEUR_NUM_DEFAUTS) = 0 Then
        
                    '--- le défaut n'est pas dans la liste alors le rajouter ---
                    ListeNumDefautsDeBase = ListeNumDefautsDeBase & SEPARATEUR_NUM_DEFAUTS & NumDefaut
        
                End If
        
            Next a
            
        End If
    
    End If
        
    '--- affectation ---
    AjoutNumDefautsSansDoublons = ListeNumDefautsDeBase

End Function

'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Rôle      : Décode une suite de n° d'alarmes de la ligne en une suite de libellés correspondant
' Entrées :              AlarmesLigne -> Suite de n° d'alarmes avec séparateur
' Retours : DecodeAlarmesLigne -> Suite des libellés des alarmes avec divers renseignements
' Détails  :
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function DecodeAlarmesLigne(ByVal AlarmesLigne As String) As String
    
    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next
    
    '--- déclaration ---
    Dim a As Integer
    Dim NumDefaut As Integer
    Dim TAlarmesLigne As Variant
           
    '--- affectation ---
    DecodeAlarmesLigne = ""
    
    If AlarmesLigne <> "" Then
                    
        '--- construction du tableau contenant les numéros d'alarmes ---
        TAlarmesLigne = Split(AlarmesLigne, SEPARATEUR_NUM_DEFAUTS)
                        
        '--- construction de la chaine des libellés ---
        For a = LBound(TAlarmesLigne) To UBound(TAlarmesLigne)
            If IsNumeric(TAlarmesLigne(a)) = True Then
                
                '--- affectation du n° du défaut ---
                NumDefaut = TAlarmesLigne(a)
                
                If NumDefaut >= DEFAUTS.NUM_MINI And NumDefaut <= DEFAUTS.NUM_MAXI Then
                    DecodeAlarmesLigne = DecodeAlarmesLigne & _
                                                          "Défaut n° " & TAlarmesLigne(a) & " - " & _
                                                          TDefauts(TAlarmesLigne(a)).LibelleDefaut & vbCrLf
                End If
            
            End If
        Next a
    
    End If
                                                                      
End Function

'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Rôle      : Décode une suite de n° d'alarmes d'un poste en une suite de libellés correspondant
' Entrées :             AlarmesPoste -> Suite de n° d'alarmes avec séparateur
' Retours : DecodeAlarmesPoste -> Suite des libellés des alarmes avec divers renseignements
' Détails  :
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function DecodeAlarmesPoste(ByVal AlarmesPoste As String) As String
    
    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next
    
    '--- déclaration ---
    Dim a As Integer
    Dim NumDefaut As Integer
    Dim TAlarmesPoste As Variant
           
    '--- affectation ---
    DecodeAlarmesPoste = ""
    
    If AlarmesPoste <> "" Then
                    
        '--- construction du tableau contenant les numéros d'alarmes ---
        TAlarmesPoste = Split(AlarmesPoste, SEPARATEUR_NUM_DEFAUTS)
                        
        '--- construction de la chaine des libellés ---
        For a = LBound(TAlarmesPoste) To UBound(TAlarmesPoste)
            If IsNumeric(TAlarmesPoste(a)) = True Then
                
                '--- affectation du n° du défaut ---
                NumDefaut = TAlarmesPoste(a)
                
                If NumDefaut >= DEFAUTS.NUM_MINI And NumDefaut <= DEFAUTS.NUM_MAXI Then
                    DecodeAlarmesPoste = DecodeAlarmesPoste & _
                                                          "Défaut n° " & TAlarmesPoste(a) & " - " & _
                                                          TDefauts(TAlarmesPoste(a)).LibelleDefaut & vbCrLf
                End If
            
            End If
        Next a
    
    End If
                                                                      
End Function

'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Rôle      : Initialise les bits d'anti-rebonds
' Entrées : NumDefaut -> n° du défaut faisant l'objet de la demande
' Retours :
' Détails  :
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Sub InitialiserAntiRebondsDefaut(ByVal NumDefaut As Long)
    
    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next
    
    '--- déclaration ---
           
    If NumDefaut >= DEFAUTS.NUM_MINI And NumDefaut <= DEFAUTS.NUM_MAXI Then
    
         With TDefauts(NumDefaut)
            
            '--- initialisation des anti-rebonds ---
            .AntiRebondGyrophare = False
            .AntiRebondKlaxon = False
            .AntiRebondTraçabiliteAlarmes = False
        
        End With
    
    End If

End Sub

'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Rôle      : Signale un défaut sur le gyrophare et le klaxon en fonction des valeurs de la table des défauts
' Entrées : NumDefaut -> n° du défaut faisant l'objet de la demande
' Retours :
' Détails  :
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Sub SignalerDefaut(ByVal NumDefaut As Long, EtatsDefaut As Boolean)
    
    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next
    
    '--- déclaration ---
           
    If NumDefaut >= DEFAUTS.NUM_MINI And NumDefaut <= DEFAUTS.NUM_MAXI Then
    
         With TDefauts(NumDefaut)
            
            '--- signalisation dans le fichier de traçabilité des alarmes ---
            If .AntiRebondTraçabiliteAlarmes = False And EtatsDefaut = True Then
                Bidon = EnregistrementDefautDansTraçabiliteAlarmes(NumDefaut, EtatsDefaut)
                .AntiRebondTraçabiliteAlarmes = True
            End If
            
            '--- signalisation dans le fichier de traçabilité des alarmes ---
            If .AntiRebondTraçabiliteAlarmes = True And EtatsDefaut = False Then
                Bidon = EnregistrementDefautDansTraçabiliteAlarmes(NumDefaut, EtatsDefaut)
                .AntiRebondTraçabiliteAlarmes = False
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
' Rôle      :  Analyse de la totalité des alarmes et visualisation sur la ligne en bas à gauche de l'écran principal
'                 des alarmes en cours
' Détails  :
' Retours :
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Sub VisualisationLigneAlarmes()
    
    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next

    '--- constantes privées ---
    
    '--- déclaration ---
    Dim UnDefautAuMoinsSignale As Boolean       'indique un défaut au moins de signaler par section (pont, cuves, etc...)
    Dim EtatDefaut As Boolean                               'représente l'état d'un défaut (FALSE = pas de défaut, TRUE = défaut en cours)
    
    Dim a As Integer                                                'pour les boucles FOR...NEXT
    Dim NumDefaut As Integer, _
            NumDefautAAfficher As Integer
    Static CptAppels As Integer                                'compteur d'appels de la routine
    Static CptDefauts As Integer                              'compteur de défauts
    Dim NumPoste1 As Integer                                'représente le poste 1 d'une cuve à postes multiples
    Dim NumPoste2 As Integer                                'représente le poste 2 d'une cuve à postes multiples
    Dim NumChargeACePoste As Integer                'représente le numéro de charge dans un poste
    
    Dim NumCuve As Long                                      'représente un numéro de cuve quelconque
    
    Dim ComplementDefaut As String                     'contient le texte du complément ajouté au libellé du défaut
    Dim LibelleCompleteDefaut As String               'représente un libellé complété d'un défaut (pour les numéros de défaut des variateurs, etc ...)
    Dim LibelleDefautAfficheur As String                 'libellé du défaut destiné à l'afficheur
    
    Dim ListeNumDefauts As String, _
            ListeNumDefautsPourUneCuve As String, _
            ListeNumDefautsPourLaLigne As String
    Dim TNumDefauts As Variant, _
            TNumDefautsPourUneCuve As Variant
   
    '*********************************************************************************************************************
    '*                                                                   ETATS DE LA LIGNE
    '*********************************************************************************************************************

    With TEtatsLigne

        '--- arrêt générale ---
        NumDefaut = .TNumDefauts.NumDefautArretGeneral
        EtatDefaut = .ArretGeneral
        ConstruitListeNumDefauts EtatDefaut, NumDefaut, ListeNumDefauts, ListeNumDefautsPourLaLigne   'contruction des listes de défauts
        EnregistreDateDetectionDisparitionDefaut EtatDefaut, NumDefaut 'enregistrement de la date de détection et disparition du défaut

        '-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

        '--- arrêt d'urgence général ---
        NumDefaut = .TNumDefauts.NumDefautArretUrgence
        EtatDefaut = .ArretUrgence
        ConstruitListeNumDefauts EtatDefaut, NumDefaut, ListeNumDefauts, ListeNumDefautsPourLaLigne    'contruction des listes de défauts
        EnregistreDateDetectionDisparitionDefaut EtatDefaut, NumDefaut 'enregistrement de la date de détection et disparition du défaut

        '-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
        
        '--- portillons et ligne de vie ---
        NumDefaut = .TNumDefauts.NumDefautPortillonsLigneVie
        EtatDefaut = .PortillonsLigneVie
        ConstruitListeNumDefauts EtatDefaut, NumDefaut, ListeNumDefauts, ListeNumDefautsPourLaLigne    'contruction des listes de défauts
        EnregistreDateDetectionDisparitionDefaut EtatDefaut, NumDefaut 'enregistrement de la date de détection et disparition du défaut

        '-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

        '--- chaine de sécurité du pont 1 ---
        NumDefaut = .TNumDefauts.NumDefautSecuriteP1
        EtatDefaut = .SecuriteP1
        ConstruitListeNumDefauts EtatDefaut, NumDefaut, ListeNumDefauts, ListeNumDefautsPourLaLigne    'contruction des listes de défauts
        EnregistreDateDetectionDisparitionDefaut EtatDefaut, NumDefaut 'enregistrement de la date de détection et disparition du défaut

        '-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

        '--- chaine de sécurité du pont 2 ---
        NumDefaut = .TNumDefauts.NumDefautSecuriteP2
        EtatDefaut = .SecuriteP2
        ConstruitListeNumDefauts EtatDefaut, NumDefaut, ListeNumDefauts, ListeNumDefautsPourLaLigne    'contruction des listes de défauts
        EnregistreDateDetectionDisparitionDefaut EtatDefaut, NumDefaut 'enregistrement de la date de détection et disparition du défaut

        '-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

        '--- manque de tension ---
        NumDefaut = .TNumDefauts.NumDefautManqueTension
        EtatDefaut = .ManqueTension
        ConstruitListeNumDefauts EtatDefaut, NumDefaut, ListeNumDefauts, ListeNumDefautsPourLaLigne    'contruction des listes de défauts
        EnregistreDateDetectionDisparitionDefaut EtatDefaut, NumDefaut     'enregistrement de la date de détection et disparition du défaut

        '-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

        '--- manque d'air ---
        NumDefaut = .TNumDefauts.NumDefautManqueAir
        EtatDefaut = .ManqueAir
        ConstruitListeNumDefauts EtatDefaut, NumDefaut, ListeNumDefauts, ListeNumDefautsPourLaLigne    'contruction des listes de défauts
        EnregistreDateDetectionDisparitionDefaut EtatDefaut, NumDefaut     'enregistrement de la date de détection et disparition du défaut

        '-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

        '--- arrêt par stop ligne ---
        NumDefaut = .TNumDefauts.NumDefautStopLigne
        EtatDefaut = .StopLigne
        ConstruitListeNumDefauts EtatDefaut, NumDefaut, ListeNumDefauts, ListeNumDefautsPourLaLigne    'contruction des listes de défauts
        EnregistreDateDetectionDisparitionDefaut EtatDefaut, NumDefaut     'enregistrement de la date de détection et disparition du défaut

        '-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

'        '--- anti collision sur les défauts des ponts ---
'        NumDefaut = .TNumDefauts.NumDefautAntiCollisionDefautsPonts
'        EtatDefaut = .AntiCollisionDefautsPonts
'        ConstruitListeNumDefauts EtatDefaut, NumDefaut, ListeNumDefauts, ListeNumDefautsPourLaLigne    'contruction des listes de défauts
'        EnregistreDateDetectionDisparitionDefaut EtatDefaut, NumDefaut     'enregistrement de la date de détection et disparition du défaut
'
'        '-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'
'        '--- anti collision sur la somme des lasers des 2 ponts ---
'        NumDefaut = .TNumDefauts.NumDefautAntiCollisionLasersPonts
'        EtatDefaut = .AntiCollisionLasersPonts
'        ConstruitListeNumDefauts EtatDefaut, NumDefaut, ListeNumDefauts, ListeNumDefautsPourLaLigne    'contruction des listes de défauts
'        EnregistreDateDetectionDisparitionDefaut EtatDefaut, NumDefaut     'enregistrement de la date de détection et disparition du défaut
'
'        '-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'
'        '--- poste occupé ---
'        NumDefaut = .TNumDefauts.NumDefautPosteOccupe
'        EtatDefaut = .PosteOccupe
'        ConstruitListeNumDefauts EtatDefaut, NumDefaut, ListeNumDefauts, ListeNumDefautsPourLaLigne    'contruction des listes de défauts
'        EnregistreDateDetectionDisparitionDefaut EtatDefaut, NumDefaut     'enregistrement de la date de détection et disparition du défaut

    End With
        
    '*********************************************************************************************************************
    '*                                                                   LES PONTS
    '*********************************************************************************************************************
    
    For a = LBound(TEtatsPonts()) To UBound(TEtatsPonts())

        With TEtatsPonts(a)

            '--- initialisation à FALSE par défaut ---
            UnDefautAuMoinsSignale = False

            '-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

            '--- Poste OQP cellule du PONT ---
            NumDefaut = .TNumDefauts.NumDefautPresencePiece
            EtatDefaut = .TEntreesAPI.M_DefautPresencePicece
            UnDefautAuMoinsSignale = UnDefautAuMoinsSignale Or EtatDefaut
            ConstruitListeNumDefauts EtatDefaut, NumDefaut, ListeNumDefauts, ListeNumDefautsPourLaLigne    'contruction des listes de défauts
            EnregistreDateDetectionDisparitionDefaut EtatDefaut, NumDefaut     'enregistrement de la date de détection et disparition du défaut

            '-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

            '--- défaut variateur de la translation du pont ---
            NumDefaut = .TNumDefauts.NumDefautDefautVariateurTrLPont
            EtatDefaut = .TEntreesAPI.M_DefautVariateurTrlPont
            UnDefautAuMoinsSignale = UnDefautAuMoinsSignale Or EtatDefaut
            ConstruitListeNumDefauts EtatDefaut, NumDefaut, ListeNumDefauts, ListeNumDefautsPourLaLigne    'contruction des listes de défauts
            EnregistreDateDetectionDisparitionDefaut EtatDefaut, NumDefaut     'enregistrement de la date de détection et disparition du défaut

            '-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

            '--- axe non référencé de la translation du pont ---
            NumDefaut = .TNumDefauts.NumDefautAxeNonReferenceTrlPont
            EtatDefaut = .TEntreesAPI.M_AxeNonReferenceTrlPont
            UnDefautAuMoinsSignale = UnDefautAuMoinsSignale Or EtatDefaut
            ConstruitListeNumDefauts EtatDefaut, NumDefaut, ListeNumDefauts, ListeNumDefautsPourLaLigne    'contruction des listes de défauts
            EnregistreDateDetectionDisparitionDefaut EtatDefaut, NumDefaut     'enregistrement de la date de détection et disparition du défaut

            '-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

            '--- défaut variateur du levage du pont ---
            NumDefaut = .TNumDefauts.NumDefautDefautVariateurLevPont
            EtatDefaut = .TEntreesAPI.M_DefautVariateurLevPont
            UnDefautAuMoinsSignale = UnDefautAuMoinsSignale Or EtatDefaut
            ConstruitListeNumDefauts EtatDefaut, NumDefaut, ListeNumDefauts, ListeNumDefautsPourLaLigne    'contruction des listes de défauts
            EnregistreDateDetectionDisparitionDefaut EtatDefaut, NumDefaut     'enregistrement de la date de détection et disparition du défaut
            
            '-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

            '--- surcourse haut du levage ---
            NumDefaut = .TNumDefauts.NumDefautSurcourseLevHaut
            EtatDefaut = .TEntreesAPI.M_SurcourseLevHaut
            UnDefautAuMoinsSignale = UnDefautAuMoinsSignale Or EtatDefaut
            ConstruitListeNumDefauts EtatDefaut, NumDefaut, ListeNumDefauts, ListeNumDefautsPourLaLigne    'contruction des listes de défauts
            EnregistreDateDetectionDisparitionDefaut EtatDefaut, NumDefaut     'enregistrement de la date de détection et disparition du défaut
            
            '-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

            '--- surcourse bas du levage ---
            NumDefaut = .TNumDefauts.NumDefautSurcourseLevBas
            EtatDefaut = .TEntreesAPI.M_SurcourseLevBas
            UnDefautAuMoinsSignale = UnDefautAuMoinsSignale Or EtatDefaut
            ConstruitListeNumDefauts EtatDefaut, NumDefaut, ListeNumDefauts, ListeNumDefautsPourLaLigne    'contruction des listes de défauts
            EnregistreDateDetectionDisparitionDefaut EtatDefaut, NumDefaut     'enregistrement de la date de détection et disparition du défaut

            '-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
            
            '--- axe non référencé du levage du pont ---
            NumDefaut = .TNumDefauts.NumDefautAxeNonReferenceLevPont
            EtatDefaut = .TEntreesAPI.M_AxeNonReferenceLevPont
            UnDefautAuMoinsSignale = UnDefautAuMoinsSignale Or EtatDefaut
            ConstruitListeNumDefauts EtatDefaut, NumDefaut, ListeNumDefauts, ListeNumDefautsPourLaLigne    'contruction des listes de défauts
            EnregistreDateDetectionDisparitionDefaut EtatDefaut, NumDefaut     'enregistrement de la date de détection et disparition du défaut

            '-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
            
            '--- délai trop long de descente des accroches ---
            NumDefaut = .TNumDefauts.NumDefautDelaiTropLongDescenteAccroches
            EtatDefaut = .TEntreesAPI.M_DelaiTropLongDescenteAccroches
            UnDefautAuMoinsSignale = UnDefautAuMoinsSignale Or EtatDefaut
            ConstruitListeNumDefauts EtatDefaut, NumDefaut, ListeNumDefauts, ListeNumDefautsPourLaLigne    'contruction des listes de défauts
            EnregistreDateDetectionDisparitionDefaut EtatDefaut, NumDefaut     'enregistrement de la date de détection et disparition du défaut

            '-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
            
            '--- délai trop long de montée des accroches ---
            NumDefaut = .TNumDefauts.NumDefautDelaiTropLongMonteeAccroches
            EtatDefaut = .TEntreesAPI.M_DelaiTropLongMonteeAccroches
            UnDefautAuMoinsSignale = UnDefautAuMoinsSignale Or EtatDefaut
            ConstruitListeNumDefauts EtatDefaut, NumDefaut, ListeNumDefauts, ListeNumDefautsPourLaLigne    'contruction des listes de défauts
            EnregistreDateDetectionDisparitionDefaut EtatDefaut, NumDefaut     'enregistrement de la date de détection et disparition du défaut

            '-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

            '--- affectation définitive ---
            .UnDefautAuMoinsSignale = UnDefautAuMoinsSignale

        End With

    Next a
    
    '*********************************************************************************************************************
    '*                                                            CUVE DE DEGRAISSAGE - C00
    '*********************************************************************************************************************
    
    '--- affectation ---
    NumCuve = CUVES_REGULATION.C_C00
    ListeNumDefautsPourUneCuve = ""
    
    '--- recherche de la présence d'une charge à ce poste ---
    NumChargeACePoste = TEtatsPostes(CorrespondanceCuvesAPIPostes(NumCuve)).NumCharge

    '--- recherche de la présence d'une charge à ce poste ---
    With TEtatsCuves(NumCuve)

        '--- initialisation à FALSE par défaut ---
        UnDefautAuMoinsSignale = False

        If TEtatsLigne.MarcheGenerale = True Then

            '-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

            '--- niveau trop bas ---
            NumDefaut = .TNumDefauts.NumDefautNiveauTresBas
            EtatDefaut = .TEntreesAPI.E_NiveauTresBas
            UnDefautAuMoinsSignale = UnDefautAuMoinsSignale Or EtatDefaut
            ConstruitListeNumDefauts EtatDefaut, NumDefaut, ListeNumDefauts, , ListeNumDefautsPourUneCuve  'contruction des listes de défauts
            EnregistreDateDetectionDisparitionDefaut EtatDefaut, NumDefaut 'enregistrement de la date de détection et disparition du défaut

            '--- niveau trop haut ---
            NumDefaut = .TNumDefauts.NumDefautNiveauTresHaut
            EtatDefaut = .TEntreesAPI.E_NiveauTresHaut
            UnDefautAuMoinsSignale = UnDefautAuMoinsSignale Or EtatDefaut
            ConstruitListeNumDefauts EtatDefaut, NumDefaut, ListeNumDefauts, , ListeNumDefautsPourUneCuve  'contruction des listes de défauts
            EnregistreDateDetectionDisparitionDefaut EtatDefaut, NumDefaut 'enregistrement de la date de détection et disparition du défaut

            '-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

            '--- défaut du chauffage ---
            NumDefaut = .TNumDefauts.NumDefautDefautChauffage
            EtatDefaut = .TEntreesAPI.E_DefautChauffage
            UnDefautAuMoinsSignale = UnDefautAuMoinsSignale Or EtatDefaut
            ConstruitListeNumDefauts EtatDefaut, NumDefaut, ListeNumDefauts, , ListeNumDefautsPourUneCuve  'contruction des listes de défauts
            EnregistreDateDetectionDisparitionDefaut EtatDefaut, NumDefaut 'enregistrement de la date de détection et disparition du défaut
            
            '-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
            
            '--- défaut PT100 ---
            NumDefaut = .TNumDefauts.NumDefautDefautPT100
            EtatDefaut = .TEntreesAPI.DefautPT100
            UnDefautAuMoinsSignale = UnDefautAuMoinsSignale Or EtatDefaut
            ConstruitListeNumDefauts EtatDefaut, NumDefaut, ListeNumDefauts, , ListeNumDefautsPourUneCuve  'contruction des listes de défauts
            EnregistreDateDetectionDisparitionDefaut EtatDefaut, NumDefaut 'enregistrement de la date de détection et disparition du défaut
            
            '--- température trop basse ---
            NumDefaut = .TNumDefauts.NumDefautTemperatureTropBasse
            EtatDefaut = .TEntreesAPI.TemperatureTropBasse
            UnDefautAuMoinsSignale = UnDefautAuMoinsSignale Or EtatDefaut
            ConstruitListeNumDefauts EtatDefaut, NumDefaut, ListeNumDefauts, , ListeNumDefautsPourUneCuve  'contruction des listes de défauts
            EnregistreDateDetectionDisparitionDefaut EtatDefaut, NumDefaut 'enregistrement de la date de détection et disparition du défaut
            
            '--- température trop haute ---
            NumDefaut = .TNumDefauts.NumDefautTemperatureTropHaute
            EtatDefaut = .TEntreesAPI.TemperatureTropHaute
            UnDefautAuMoinsSignale = UnDefautAuMoinsSignale Or EtatDefaut
            ConstruitListeNumDefauts EtatDefaut, NumDefaut, ListeNumDefauts, , ListeNumDefautsPourUneCuve  'contruction des listes de défauts
            EnregistreDateDetectionDisparitionDefaut EtatDefaut, NumDefaut 'enregistrement de la date de détection et disparition du défaut

            '-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

        End If

        '--- construction de la liste des numéros de défauts pour une cuve ayant une charge ---
        If NumChargeACePoste = 0 Then
            .ListeNumDefautsSiCharge = ""    'vider la liste si pas de charge dans le poste
        Else
            If ListeNumDefautsPourUneCuve <> "" Then  'remplir la liste en évitant les doublons
                                                                                     'suppression du dernier séparateur inutile car dans fonction
                .ListeNumDefautsSiCharge = AjoutNumDefautsSansDoublons(.ListeNumDefautsSiCharge, _
                                                                                                                     ListeNumDefautsPourUneCuve) 'remplir la liste en éliminant les doublons
            End If
        End If

        '--- affectation définitive ---
        .UnDefautAuMoinsSignale = UnDefautAuMoinsSignale

    End With
    
    '*********************************************************************************************************************
    '*                                                                CUVE C01
    '*********************************************************************************************************************
    If False Then
       '--- affectation ---
        'NumCuve = CUVES_REGULATION.C_SAT
        ListeNumDefautsPourUneCuve = ""
        
        '--- recherche de la présence d'une charge à ce poste ---
        NumChargeACePoste = TEtatsPostes(CorrespondanceCuvesAPIPostes(NumCuve)).NumCharge
    
        With TEtatsCuves(NumCuve)
    
            '--- initialisation à FALSE par défaut ---
            UnDefautAuMoinsSignale = False
    
            If TEtatsLigne.MarcheGenerale = True Then
    
                '-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    
                '--- niveau trop bas ---
                NumDefaut = .TNumDefauts.NumDefautNiveauTresBas
                EtatDefaut = .TEntreesAPI.E_NiveauTresBas
                UnDefautAuMoinsSignale = UnDefautAuMoinsSignale Or EtatDefaut
                ConstruitListeNumDefauts EtatDefaut, NumDefaut, ListeNumDefauts, , ListeNumDefautsPourUneCuve  'contruction des listes de défauts
                EnregistreDateDetectionDisparitionDefaut EtatDefaut, NumDefaut 'enregistrement de la date de détection et disparition du défaut
    
                '--- niveau trop haut ---
                NumDefaut = .TNumDefauts.NumDefautNiveauTresHaut
                EtatDefaut = .TEntreesAPI.E_NiveauTresHaut
                UnDefautAuMoinsSignale = UnDefautAuMoinsSignale Or EtatDefaut
                ConstruitListeNumDefauts EtatDefaut, NumDefaut, ListeNumDefauts, , ListeNumDefautsPourUneCuve  'contruction des listes de défauts
                EnregistreDateDetectionDisparitionDefaut EtatDefaut, NumDefaut 'enregistrement de la date de détection et disparition du défaut
    
                '-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    
                '--- défaut du chauffage ---
                NumDefaut = .TNumDefauts.NumDefautDefautChauffage
                EtatDefaut = .TEntreesAPI.E_DefautChauffage
                UnDefautAuMoinsSignale = UnDefautAuMoinsSignale Or EtatDefaut
                ConstruitListeNumDefauts EtatDefaut, NumDefaut, ListeNumDefauts, , ListeNumDefautsPourUneCuve  'contruction des listes de défauts
                EnregistreDateDetectionDisparitionDefaut EtatDefaut, NumDefaut 'enregistrement de la date de détection et disparition du défaut
                
                '-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
                
                '--- défaut PT100 ---
                NumDefaut = .TNumDefauts.NumDefautDefautPT100
                EtatDefaut = .TEntreesAPI.DefautPT100
                UnDefautAuMoinsSignale = UnDefautAuMoinsSignale Or EtatDefaut
                ConstruitListeNumDefauts EtatDefaut, NumDefaut, ListeNumDefauts, , ListeNumDefautsPourUneCuve  'contruction des listes de défauts
                EnregistreDateDetectionDisparitionDefaut EtatDefaut, NumDefaut 'enregistrement de la date de détection et disparition du défaut
                
                '--- température trop basse ---
                NumDefaut = .TNumDefauts.NumDefautTemperatureTropBasse
                EtatDefaut = .TEntreesAPI.TemperatureTropBasse
                UnDefautAuMoinsSignale = UnDefautAuMoinsSignale Or EtatDefaut
                ConstruitListeNumDefauts EtatDefaut, NumDefaut, ListeNumDefauts, , ListeNumDefautsPourUneCuve  'contruction des listes de défauts
                EnregistreDateDetectionDisparitionDefaut EtatDefaut, NumDefaut 'enregistrement de la date de détection et disparition du défaut
                
                '--- température trop haute ---
                NumDefaut = .TNumDefauts.NumDefautTemperatureTropHaute
                EtatDefaut = .TEntreesAPI.TemperatureTropHaute
                UnDefautAuMoinsSignale = UnDefautAuMoinsSignale Or EtatDefaut
                ConstruitListeNumDefauts EtatDefaut, NumDefaut, ListeNumDefauts, , ListeNumDefautsPourUneCuve  'contruction des listes de défauts
                EnregistreDateDetectionDisparitionDefaut EtatDefaut, NumDefaut 'enregistrement de la date de détection et disparition du défaut
    
                '-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    
            End If
    
            '--- construction de la liste des numéros de défauts pour une cuve ayant une charge ---
            If NumChargeACePoste = 0 Then
                .ListeNumDefautsSiCharge = ""    'vider la liste si pas de charge dans le poste
            Else
                If ListeNumDefautsPourUneCuve <> "" Then  'remplir la liste en évitant les doublons
                                                                                         'suppression du dernier séparateur inutile car dans fonction
                    .ListeNumDefautsSiCharge = AjoutNumDefautsSansDoublons(.ListeNumDefautsSiCharge, _
                                                                                                                         ListeNumDefautsPourUneCuve) 'remplir la liste en éliminant les doublons
                End If
            End If
    
            '--- affectation définitive ---
            .UnDefautAuMoinsSignale = UnDefautAuMoinsSignale
    
        End With
    
    End If
     
        
    
    '*********************************************************************************************************************
    '*                                                                CUVE C02
    '*********************************************************************************************************************
    
    '--- affectation ---
    NumCuve = CUVES_REGULATION.C_DEC
    ListeNumDefautsPourUneCuve = ""
    
    '--- recherche de la présence d'une charge à ce poste ---
    NumChargeACePoste = TEtatsPostes(CorrespondanceCuvesAPIPostes(NumCuve)).NumCharge

    With TEtatsCuves(NumCuve)

        '--- initialisation à FALSE par défaut ---
        UnDefautAuMoinsSignale = False

        If TEtatsLigne.MarcheGenerale = True Then

            '-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
            
            '--- niveau trop bas ---
            NumDefaut = .TNumDefauts.NumDefautNiveauTresBas
            EtatDefaut = .TEntreesAPI.E_NiveauTresBas
            UnDefautAuMoinsSignale = UnDefautAuMoinsSignale Or EtatDefaut
            ConstruitListeNumDefauts EtatDefaut, NumDefaut, ListeNumDefauts, , ListeNumDefautsPourUneCuve  'contruction des listes de défauts
            EnregistreDateDetectionDisparitionDefaut EtatDefaut, NumDefaut 'enregistrement de la date de détection et disparition du défaut

            '--- niveau trop haut ---
            NumDefaut = .TNumDefauts.NumDefautNiveauTresHaut
            EtatDefaut = .TEntreesAPI.E_NiveauTresHaut
            UnDefautAuMoinsSignale = UnDefautAuMoinsSignale Or EtatDefaut
            ConstruitListeNumDefauts EtatDefaut, NumDefaut, ListeNumDefauts, , ListeNumDefautsPourUneCuve  'contruction des listes de défauts
            EnregistreDateDetectionDisparitionDefaut EtatDefaut, NumDefaut 'enregistrement de la date de détection et disparition du défaut

            '-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

            '--- défaut du chauffage ---
            NumDefaut = .TNumDefauts.NumDefautDefautChauffage
            EtatDefaut = .TEntreesAPI.E_DefautChauffage
            UnDefautAuMoinsSignale = UnDefautAuMoinsSignale Or EtatDefaut
            ConstruitListeNumDefauts EtatDefaut, NumDefaut, ListeNumDefauts, , ListeNumDefautsPourUneCuve  'contruction des listes de défauts
            EnregistreDateDetectionDisparitionDefaut EtatDefaut, NumDefaut 'enregistrement de la date de détection et disparition du défaut
            
            '-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
            
            '--- défaut PT100 ---
            NumDefaut = .TNumDefauts.NumDefautDefautPT100
            EtatDefaut = .TEntreesAPI.DefautPT100
            UnDefautAuMoinsSignale = UnDefautAuMoinsSignale Or EtatDefaut
            ConstruitListeNumDefauts EtatDefaut, NumDefaut, ListeNumDefauts, , ListeNumDefautsPourUneCuve  'contruction des listes de défauts
            EnregistreDateDetectionDisparitionDefaut EtatDefaut, NumDefaut 'enregistrement de la date de détection et disparition du défaut
            
            '--- température trop basse ---
            NumDefaut = .TNumDefauts.NumDefautTemperatureTropBasse
            EtatDefaut = .TEntreesAPI.TemperatureTropBasse
            UnDefautAuMoinsSignale = UnDefautAuMoinsSignale Or EtatDefaut
            ConstruitListeNumDefauts EtatDefaut, NumDefaut, ListeNumDefauts, , ListeNumDefautsPourUneCuve  'contruction des listes de défauts
            EnregistreDateDetectionDisparitionDefaut EtatDefaut, NumDefaut 'enregistrement de la date de détection et disparition du défaut
            
            '--- température trop haute ---
            NumDefaut = .TNumDefauts.NumDefautTemperatureTropHaute
            EtatDefaut = .TEntreesAPI.TemperatureTropHaute
            UnDefautAuMoinsSignale = UnDefautAuMoinsSignale Or EtatDefaut
            ConstruitListeNumDefauts EtatDefaut, NumDefaut, ListeNumDefauts, , ListeNumDefautsPourUneCuve  'contruction des listes de défauts
            EnregistreDateDetectionDisparitionDefaut EtatDefaut, NumDefaut 'enregistrement de la date de détection et disparition du défaut

            '-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

        End If

        '--- construction de la liste des numéros de défauts pour une cuve ayant une charge ---
        If NumChargeACePoste = 0 Then
            .ListeNumDefautsSiCharge = ""    'vider la liste si pas de charge dans le poste
        Else
            If ListeNumDefautsPourUneCuve <> "" Then  'remplir la liste en évitant les doublons
                                                                                     'suppression du dernier séparateur inutile car dans fonction
                .ListeNumDefautsSiCharge = AjoutNumDefautsSansDoublons(.ListeNumDefautsSiCharge, _
                                                                                                                     ListeNumDefautsPourUneCuve) 'remplir la liste en éliminant les doublons
            End If
        End If

        '--- affectation définitive ---
        .UnDefautAuMoinsSignale = UnDefautAuMoinsSignale

    End With
    
    '*********************************************************************************************************************
    '*                                                                CUVE C03
    '*********************************************************************************************************************
    
    '--- affectation ---
    If (False) Then
    
    'NumCuve = CUVES_REGULATION.C_C03
    ListeNumDefautsPourUneCuve = ""

    '--- recherche de la présence d'une charge à ce poste ---
    NumChargeACePoste = TEtatsPostes(CorrespondanceCuvesAPIPostes(NumCuve)).NumCharge

    With TEtatsCuves(NumCuve)

        '--- initialisation à FALSE par défaut ---
        UnDefautAuMoinsSignale = False

        If TEtatsLigne.MarcheGenerale = True Then

            '-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
            
            '--- niveau trop bas ---
            NumDefaut = .TNumDefauts.NumDefautNiveauTresBas
            EtatDefaut = .TEntreesAPI.E_NiveauTresBas
            UnDefautAuMoinsSignale = UnDefautAuMoinsSignale Or EtatDefaut
            ConstruitListeNumDefauts EtatDefaut, NumDefaut, ListeNumDefauts, , ListeNumDefautsPourUneCuve  'contruction des listes de défauts
            EnregistreDateDetectionDisparitionDefaut EtatDefaut, NumDefaut 'enregistrement de la date de détection et disparition du défaut

            '--- niveau trop haut ---
            NumDefaut = .TNumDefauts.NumDefautNiveauTresHaut
            EtatDefaut = .TEntreesAPI.E_NiveauTresHaut
            UnDefautAuMoinsSignale = UnDefautAuMoinsSignale Or EtatDefaut
            ConstruitListeNumDefauts EtatDefaut, NumDefaut, ListeNumDefauts, , ListeNumDefautsPourUneCuve  'contruction des listes de défauts
            EnregistreDateDetectionDisparitionDefaut EtatDefaut, NumDefaut 'enregistrement de la date de détection et disparition du défaut

            '-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

            '--- défaut du chauffage ---
            NumDefaut = .TNumDefauts.NumDefautDefautChauffage
            EtatDefaut = .TEntreesAPI.E_DefautChauffage
            UnDefautAuMoinsSignale = UnDefautAuMoinsSignale Or EtatDefaut
            ConstruitListeNumDefauts EtatDefaut, NumDefaut, ListeNumDefauts, , ListeNumDefautsPourUneCuve  'contruction des listes de défauts
            EnregistreDateDetectionDisparitionDefaut EtatDefaut, NumDefaut 'enregistrement de la date de détection et disparition du défaut
            
            '-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
            
            '--- défaut PT100 ---
            NumDefaut = .TNumDefauts.NumDefautDefautPT100
            EtatDefaut = .TEntreesAPI.DefautPT100
            UnDefautAuMoinsSignale = UnDefautAuMoinsSignale Or EtatDefaut
            ConstruitListeNumDefauts EtatDefaut, NumDefaut, ListeNumDefauts, , ListeNumDefautsPourUneCuve  'contruction des listes de défauts
            EnregistreDateDetectionDisparitionDefaut EtatDefaut, NumDefaut 'enregistrement de la date de détection et disparition du défaut
            
            '--- température trop basse ---
            NumDefaut = .TNumDefauts.NumDefautTemperatureTropBasse
            EtatDefaut = .TEntreesAPI.TemperatureTropBasse
            UnDefautAuMoinsSignale = UnDefautAuMoinsSignale Or EtatDefaut
            ConstruitListeNumDefauts EtatDefaut, NumDefaut, ListeNumDefauts, , ListeNumDefautsPourUneCuve  'contruction des listes de défauts
            EnregistreDateDetectionDisparitionDefaut EtatDefaut, NumDefaut 'enregistrement de la date de détection et disparition du défaut
            
            '--- température trop haute ---
            NumDefaut = .TNumDefauts.NumDefautTemperatureTropHaute
            EtatDefaut = .TEntreesAPI.TemperatureTropHaute
            UnDefautAuMoinsSignale = UnDefautAuMoinsSignale Or EtatDefaut
            ConstruitListeNumDefauts EtatDefaut, NumDefaut, ListeNumDefauts, , ListeNumDefautsPourUneCuve  'contruction des listes de défauts
            EnregistreDateDetectionDisparitionDefaut EtatDefaut, NumDefaut 'enregistrement de la date de détection et disparition du défaut

            '-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

        End If

        '--- construction de la liste des numéros de défauts pour une cuve ayant une charge ---
        If NumChargeACePoste = 0 Then
            .ListeNumDefautsSiCharge = ""    'vider la liste si pas de charge dans le poste
        Else
            If ListeNumDefautsPourUneCuve <> "" Then  'remplir la liste en évitant les doublons
                                                                                     'suppression du dernier séparateur inutile car dans fonction
                .ListeNumDefautsSiCharge = AjoutNumDefautsSansDoublons(.ListeNumDefautsSiCharge, _
                                                                                                                     ListeNumDefautsPourUneCuve) 'remplir la liste en éliminant les doublons
            End If
        End If

        '--- affectation définitive ---
        .UnDefautAuMoinsSignale = UnDefautAuMoinsSignale

    End With

    End If
    
 
    
   
   
    
    '*********************************************************************************************************************
    '*                                                                CUVE C07
    '*********************************************************************************************************************
    
    '--- affectation ---
    NumCuve = CUVES_REGULATION.C_C07
    ListeNumDefautsPourUneCuve = ""
    
    '--- recherche de la présence d'une charge à ce poste ---
    NumChargeACePoste = TEtatsPostes(CorrespondanceCuvesAPIPostes(NumCuve)).NumCharge

    With TEtatsCuves(NumCuve)

        '--- initialisation à FALSE par défaut ---
        UnDefautAuMoinsSignale = False

        If TEtatsLigne.MarcheGenerale = True Then

            '-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
            
            '--- niveau trop bas ---
            NumDefaut = .TNumDefauts.NumDefautNiveauTresBas
            EtatDefaut = .TEntreesAPI.E_NiveauTresBas
            UnDefautAuMoinsSignale = UnDefautAuMoinsSignale Or EtatDefaut
            ConstruitListeNumDefauts EtatDefaut, NumDefaut, ListeNumDefauts, , ListeNumDefautsPourUneCuve  'contruction des listes de défauts
            EnregistreDateDetectionDisparitionDefaut EtatDefaut, NumDefaut 'enregistrement de la date de détection et disparition du défaut

            '--- niveau trop haut ---
            NumDefaut = .TNumDefauts.NumDefautNiveauTresHaut
            EtatDefaut = .TEntreesAPI.E_NiveauTresHaut
            UnDefautAuMoinsSignale = UnDefautAuMoinsSignale Or EtatDefaut
            ConstruitListeNumDefauts EtatDefaut, NumDefaut, ListeNumDefauts, , ListeNumDefautsPourUneCuve  'contruction des listes de défauts
            EnregistreDateDetectionDisparitionDefaut EtatDefaut, NumDefaut 'enregistrement de la date de détection et disparition du défaut

            '-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

            '--- défaut du chauffage ---
            NumDefaut = .TNumDefauts.NumDefautDefautChauffage
            EtatDefaut = .TEntreesAPI.E_DefautChauffage
            UnDefautAuMoinsSignale = UnDefautAuMoinsSignale Or EtatDefaut
            ConstruitListeNumDefauts EtatDefaut, NumDefaut, ListeNumDefauts, , ListeNumDefautsPourUneCuve  'contruction des listes de défauts
            EnregistreDateDetectionDisparitionDefaut EtatDefaut, NumDefaut 'enregistrement de la date de détection et disparition du défaut
            
            '-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
            
            '--- défaut PT100 ---
            NumDefaut = .TNumDefauts.NumDefautDefautPT100
            EtatDefaut = .TEntreesAPI.DefautPT100
            UnDefautAuMoinsSignale = UnDefautAuMoinsSignale Or EtatDefaut
            ConstruitListeNumDefauts EtatDefaut, NumDefaut, ListeNumDefauts, , ListeNumDefautsPourUneCuve  'contruction des listes de défauts
            EnregistreDateDetectionDisparitionDefaut EtatDefaut, NumDefaut 'enregistrement de la date de détection et disparition du défaut
            
            '--- température trop basse ---
            NumDefaut = .TNumDefauts.NumDefautTemperatureTropBasse
            EtatDefaut = .TEntreesAPI.TemperatureTropBasse
            UnDefautAuMoinsSignale = UnDefautAuMoinsSignale Or EtatDefaut
            ConstruitListeNumDefauts EtatDefaut, NumDefaut, ListeNumDefauts, , ListeNumDefautsPourUneCuve  'contruction des listes de défauts
            EnregistreDateDetectionDisparitionDefaut EtatDefaut, NumDefaut 'enregistrement de la date de détection et disparition du défaut
            
            '--- température trop haute ---
            NumDefaut = .TNumDefauts.NumDefautTemperatureTropHaute
            EtatDefaut = .TEntreesAPI.TemperatureTropHaute
            UnDefautAuMoinsSignale = UnDefautAuMoinsSignale Or EtatDefaut
            ConstruitListeNumDefauts EtatDefaut, NumDefaut, ListeNumDefauts, , ListeNumDefautsPourUneCuve  'contruction des listes de défauts
            EnregistreDateDetectionDisparitionDefaut EtatDefaut, NumDefaut 'enregistrement de la date de détection et disparition du défaut

            '-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

        End If

        '--- construction de la liste des numéros de défauts pour une cuve ayant une charge ---
        If NumChargeACePoste = 0 Then
            .ListeNumDefautsSiCharge = ""    'vider la liste si pas de charge dans le poste
        Else
            If ListeNumDefautsPourUneCuve <> "" Then  'remplir la liste en évitant les doublons
                                                                                     'suppression du dernier séparateur inutile car dans fonction
                .ListeNumDefautsSiCharge = AjoutNumDefautsSansDoublons(.ListeNumDefautsSiCharge, _
                                                                                                                     ListeNumDefautsPourUneCuve) 'remplir la liste en éliminant les doublons
            End If
        End If

        '--- affectation définitive ---
        .UnDefautAuMoinsSignale = UnDefautAuMoinsSignale

    End With
    
    '*********************************************************************************************************************
    '*                                                                CUVE C13
    '*********************************************************************************************************************
    
    '--- affectation ---
    NumCuve = CUVES_REGULATION.C_C13
    ListeNumDefautsPourUneCuve = ""
    
    '--- recherche de la présence d'une charge à ce poste ---
    NumChargeACePoste = TEtatsPostes(CorrespondanceCuvesAPIPostes(NumCuve)).NumCharge

    With TEtatsCuves(NumCuve)

        '--- initialisation à FALSE par défaut ---
        UnDefautAuMoinsSignale = False

        If TEtatsLigne.MarcheGenerale = True Then

            '-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
            
            '--- niveau trop bas ---
            NumDefaut = .TNumDefauts.NumDefautNiveauTresBas
            EtatDefaut = .TEntreesAPI.E_NiveauTresBas
            UnDefautAuMoinsSignale = UnDefautAuMoinsSignale Or EtatDefaut
            ConstruitListeNumDefauts EtatDefaut, NumDefaut, ListeNumDefauts, , ListeNumDefautsPourUneCuve  'contruction des listes de défauts
            EnregistreDateDetectionDisparitionDefaut EtatDefaut, NumDefaut 'enregistrement de la date de détection et disparition du défaut

            '--- niveau trop haut ---
            NumDefaut = .TNumDefauts.NumDefautNiveauTresHaut
            EtatDefaut = .TEntreesAPI.E_NiveauTresHaut
            UnDefautAuMoinsSignale = UnDefautAuMoinsSignale Or EtatDefaut
            ConstruitListeNumDefauts EtatDefaut, NumDefaut, ListeNumDefauts, , ListeNumDefautsPourUneCuve  'contruction des listes de défauts
            EnregistreDateDetectionDisparitionDefaut EtatDefaut, NumDefaut 'enregistrement de la date de détection et disparition du défaut

            '-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

            '--- défaut du chauffage ---
            NumDefaut = .TNumDefauts.NumDefautDefautChauffage
            EtatDefaut = .TEntreesAPI.E_DefautChauffage
            UnDefautAuMoinsSignale = UnDefautAuMoinsSignale Or EtatDefaut
            ConstruitListeNumDefauts EtatDefaut, NumDefaut, ListeNumDefauts, , ListeNumDefautsPourUneCuve  'contruction des listes de défauts
            EnregistreDateDetectionDisparitionDefaut EtatDefaut, NumDefaut 'enregistrement de la date de détection et disparition du défaut
            
            '-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
            
            '--- défaut PT100 ---
            NumDefaut = .TNumDefauts.NumDefautDefautPT100
            EtatDefaut = .TEntreesAPI.DefautPT100
            UnDefautAuMoinsSignale = UnDefautAuMoinsSignale Or EtatDefaut
            ConstruitListeNumDefauts EtatDefaut, NumDefaut, ListeNumDefauts, , ListeNumDefautsPourUneCuve  'contruction des listes de défauts
            EnregistreDateDetectionDisparitionDefaut EtatDefaut, NumDefaut 'enregistrement de la date de détection et disparition du défaut
            
            '--- température trop basse ---
            NumDefaut = .TNumDefauts.NumDefautTemperatureTropBasse
            EtatDefaut = .TEntreesAPI.TemperatureTropBasse
            UnDefautAuMoinsSignale = UnDefautAuMoinsSignale Or EtatDefaut
            ConstruitListeNumDefauts EtatDefaut, NumDefaut, ListeNumDefauts, , ListeNumDefautsPourUneCuve  'contruction des listes de défauts
            EnregistreDateDetectionDisparitionDefaut EtatDefaut, NumDefaut 'enregistrement de la date de détection et disparition du défaut
            
            '--- température trop haute ---
            NumDefaut = .TNumDefauts.NumDefautTemperatureTropHaute
            EtatDefaut = .TEntreesAPI.TemperatureTropHaute
            UnDefautAuMoinsSignale = UnDefautAuMoinsSignale Or EtatDefaut
            ConstruitListeNumDefauts EtatDefaut, NumDefaut, ListeNumDefauts, , ListeNumDefautsPourUneCuve  'contruction des listes de défauts
            EnregistreDateDetectionDisparitionDefaut EtatDefaut, NumDefaut 'enregistrement de la date de détection et disparition du défaut

            '-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

            '--- défaut refroidissement ---
            NumDefaut = .TNumDefauts.NumDefautDefautRefroidissement
            EtatDefaut = .TEntreesAPI.E_DefautRefroidissement
            UnDefautAuMoinsSignale = UnDefautAuMoinsSignale Or EtatDefaut
            ConstruitListeNumDefauts EtatDefaut, NumDefaut, ListeNumDefauts, , ListeNumDefautsPourUneCuve  'contruction des listes de défauts
            EnregistreDateDetectionDisparitionDefaut EtatDefaut, NumDefaut 'enregistrement de la date de détection et disparition du défaut
        
            '-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
        
        End If

        '--- construction de la liste des numéros de défauts pour une cuve ayant une charge ---
        If NumChargeACePoste = 0 Then
            .ListeNumDefautsSiCharge = ""    'vider la liste si pas de charge dans le poste
        Else
            If ListeNumDefautsPourUneCuve <> "" Then  'remplir la liste en évitant les doublons
                                                                                     'suppression du dernier séparateur inutile car dans fonction
                .ListeNumDefautsSiCharge = AjoutNumDefautsSansDoublons(.ListeNumDefautsSiCharge, _
                                                                                                                     ListeNumDefautsPourUneCuve) 'remplir la liste en éliminant les doublons
            End If
        End If

        '--- affectation définitive ---
        .UnDefautAuMoinsSignale = UnDefautAuMoinsSignale

    End With
    
    '*********************************************************************************************************************
    '*                                                                CUVE C14
    '*********************************************************************************************************************
    
    '--- affectation ---
    NumCuve = CUVES_REGULATION.C_C14
    ListeNumDefautsPourUneCuve = ""
    
    '--- recherche de la présence d'une charge à ce poste ---
    NumChargeACePoste = TEtatsPostes(CorrespondanceCuvesAPIPostes(NumCuve)).NumCharge

    With TEtatsCuves(NumCuve)

        '--- initialisation à FALSE par défaut ---
        UnDefautAuMoinsSignale = False

        If TEtatsLigne.MarcheGenerale = True Then

            '-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
            
            '--- niveau trop bas ---
            NumDefaut = .TNumDefauts.NumDefautNiveauTresBas
            EtatDefaut = .TEntreesAPI.E_NiveauTresBas
            UnDefautAuMoinsSignale = UnDefautAuMoinsSignale Or EtatDefaut
            ConstruitListeNumDefauts EtatDefaut, NumDefaut, ListeNumDefauts, , ListeNumDefautsPourUneCuve  'contruction des listes de défauts
            EnregistreDateDetectionDisparitionDefaut EtatDefaut, NumDefaut 'enregistrement de la date de détection et disparition du défaut

            '--- niveau trop haut ---
            NumDefaut = .TNumDefauts.NumDefautNiveauTresHaut
            EtatDefaut = .TEntreesAPI.E_NiveauTresHaut
            UnDefautAuMoinsSignale = UnDefautAuMoinsSignale Or EtatDefaut
            ConstruitListeNumDefauts EtatDefaut, NumDefaut, ListeNumDefauts, , ListeNumDefautsPourUneCuve  'contruction des listes de défauts
            EnregistreDateDetectionDisparitionDefaut EtatDefaut, NumDefaut 'enregistrement de la date de détection et disparition du défaut

            '-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

            '--- défaut du chauffage ---
            NumDefaut = .TNumDefauts.NumDefautDefautChauffage
            EtatDefaut = .TEntreesAPI.E_DefautChauffage
            UnDefautAuMoinsSignale = UnDefautAuMoinsSignale Or EtatDefaut
            ConstruitListeNumDefauts EtatDefaut, NumDefaut, ListeNumDefauts, , ListeNumDefautsPourUneCuve  'contruction des listes de défauts
            EnregistreDateDetectionDisparitionDefaut EtatDefaut, NumDefaut 'enregistrement de la date de détection et disparition du défaut
            
            '-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
            
            '--- défaut PT100 ---
            NumDefaut = .TNumDefauts.NumDefautDefautPT100
            EtatDefaut = .TEntreesAPI.DefautPT100
            UnDefautAuMoinsSignale = UnDefautAuMoinsSignale Or EtatDefaut
            ConstruitListeNumDefauts EtatDefaut, NumDefaut, ListeNumDefauts, , ListeNumDefautsPourUneCuve  'contruction des listes de défauts
            EnregistreDateDetectionDisparitionDefaut EtatDefaut, NumDefaut 'enregistrement de la date de détection et disparition du défaut
            
            '--- température trop basse ---
            NumDefaut = .TNumDefauts.NumDefautTemperatureTropBasse
            EtatDefaut = .TEntreesAPI.TemperatureTropBasse
            UnDefautAuMoinsSignale = UnDefautAuMoinsSignale Or EtatDefaut
            ConstruitListeNumDefauts EtatDefaut, NumDefaut, ListeNumDefauts, , ListeNumDefautsPourUneCuve  'contruction des listes de défauts
            EnregistreDateDetectionDisparitionDefaut EtatDefaut, NumDefaut 'enregistrement de la date de détection et disparition du défaut
            
            '--- température trop haute ---
            NumDefaut = .TNumDefauts.NumDefautTemperatureTropHaute
            EtatDefaut = .TEntreesAPI.TemperatureTropHaute
            UnDefautAuMoinsSignale = UnDefautAuMoinsSignale Or EtatDefaut
            ConstruitListeNumDefauts EtatDefaut, NumDefaut, ListeNumDefauts, , ListeNumDefautsPourUneCuve  'contruction des listes de défauts
            EnregistreDateDetectionDisparitionDefaut EtatDefaut, NumDefaut 'enregistrement de la date de détection et disparition du défaut

            '-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

            '--- défaut refroidissement ---
            NumDefaut = .TNumDefauts.NumDefautDefautRefroidissement
            EtatDefaut = .TEntreesAPI.E_DefautRefroidissement
            UnDefautAuMoinsSignale = UnDefautAuMoinsSignale Or EtatDefaut
            ConstruitListeNumDefauts EtatDefaut, NumDefaut, ListeNumDefauts, , ListeNumDefautsPourUneCuve  'contruction des listes de défauts
            EnregistreDateDetectionDisparitionDefaut EtatDefaut, NumDefaut 'enregistrement de la date de détection et disparition du défaut
        
            '-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
        
        End If

        '--- construction de la liste des numéros de défauts pour une cuve ayant une charge ---
        If NumChargeACePoste = 0 Then
            .ListeNumDefautsSiCharge = ""    'vider la liste si pas de charge dans le poste
        Else
            If ListeNumDefautsPourUneCuve <> "" Then  'remplir la liste en évitant les doublons
                                                                                     'suppression du dernier séparateur inutile car dans fonction
                .ListeNumDefautsSiCharge = AjoutNumDefautsSansDoublons(.ListeNumDefautsSiCharge, _
                                                                                                                     ListeNumDefautsPourUneCuve) 'remplir la liste en éliminant les doublons
            End If
        End If

        '--- affectation définitive ---
        .UnDefautAuMoinsSignale = UnDefautAuMoinsSignale

    End With
    
    '*********************************************************************************************************************
    '*                                                                CUVE C15
    '*********************************************************************************************************************
    
    '--- affectation ---
    NumCuve = CUVES_REGULATION.C_C15
    ListeNumDefautsPourUneCuve = ""
    
    '--- recherche de la présence d'une charge à ce poste ---
    NumChargeACePoste = TEtatsPostes(CorrespondanceCuvesAPIPostes(NumCuve)).NumCharge

    With TEtatsCuves(NumCuve)

        '--- initialisation à FALSE par défaut ---
        UnDefautAuMoinsSignale = False

        If TEtatsLigne.MarcheGenerale = True Then

            '-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

            '--- niveau trop bas ---
            NumDefaut = .TNumDefauts.NumDefautNiveauTresBas
            EtatDefaut = .TEntreesAPI.E_NiveauTresBas
            UnDefautAuMoinsSignale = UnDefautAuMoinsSignale Or EtatDefaut
            ConstruitListeNumDefauts EtatDefaut, NumDefaut, ListeNumDefauts, , ListeNumDefautsPourUneCuve  'contruction des listes de défauts
            EnregistreDateDetectionDisparitionDefaut EtatDefaut, NumDefaut 'enregistrement de la date de détection et disparition du défaut

            '--- niveau trop haut ---
            NumDefaut = .TNumDefauts.NumDefautNiveauTresHaut
            EtatDefaut = .TEntreesAPI.E_NiveauTresHaut
            UnDefautAuMoinsSignale = UnDefautAuMoinsSignale Or EtatDefaut
            ConstruitListeNumDefauts EtatDefaut, NumDefaut, ListeNumDefauts, , ListeNumDefautsPourUneCuve  'contruction des listes de défauts
            EnregistreDateDetectionDisparitionDefaut EtatDefaut, NumDefaut 'enregistrement de la date de détection et disparition du défaut

            '-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

            '--- défaut du chauffage ---
            NumDefaut = .TNumDefauts.NumDefautDefautChauffage
            EtatDefaut = .TEntreesAPI.E_DefautChauffage
            UnDefautAuMoinsSignale = UnDefautAuMoinsSignale Or EtatDefaut
            ConstruitListeNumDefauts EtatDefaut, NumDefaut, ListeNumDefauts, , ListeNumDefautsPourUneCuve  'contruction des listes de défauts
            EnregistreDateDetectionDisparitionDefaut EtatDefaut, NumDefaut 'enregistrement de la date de détection et disparition du défaut
            
            '-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
            
            '--- défaut PT100 ---
            NumDefaut = .TNumDefauts.NumDefautDefautPT100
            EtatDefaut = .TEntreesAPI.DefautPT100
            UnDefautAuMoinsSignale = UnDefautAuMoinsSignale Or EtatDefaut
            ConstruitListeNumDefauts EtatDefaut, NumDefaut, ListeNumDefauts, , ListeNumDefautsPourUneCuve  'contruction des listes de défauts
            EnregistreDateDetectionDisparitionDefaut EtatDefaut, NumDefaut 'enregistrement de la date de détection et disparition du défaut
            
            '--- température trop basse ---
            NumDefaut = .TNumDefauts.NumDefautTemperatureTropBasse
            EtatDefaut = .TEntreesAPI.TemperatureTropBasse
            UnDefautAuMoinsSignale = UnDefautAuMoinsSignale Or EtatDefaut
            ConstruitListeNumDefauts EtatDefaut, NumDefaut, ListeNumDefauts, , ListeNumDefautsPourUneCuve  'contruction des listes de défauts
            EnregistreDateDetectionDisparitionDefaut EtatDefaut, NumDefaut 'enregistrement de la date de détection et disparition du défaut
            
            '--- température trop haute ---
            NumDefaut = .TNumDefauts.NumDefautTemperatureTropHaute
            EtatDefaut = .TEntreesAPI.TemperatureTropHaute
            UnDefautAuMoinsSignale = UnDefautAuMoinsSignale Or EtatDefaut
            ConstruitListeNumDefauts EtatDefaut, NumDefaut, ListeNumDefauts, , ListeNumDefautsPourUneCuve  'contruction des listes de défauts
            EnregistreDateDetectionDisparitionDefaut EtatDefaut, NumDefaut 'enregistrement de la date de détection et disparition du défaut

            '-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

            '--- défaut refroidissement ---
            NumDefaut = .TNumDefauts.NumDefautDefautRefroidissement
            EtatDefaut = .TEntreesAPI.E_DefautRefroidissement
            UnDefautAuMoinsSignale = UnDefautAuMoinsSignale Or EtatDefaut
            ConstruitListeNumDefauts EtatDefaut, NumDefaut, ListeNumDefauts, , ListeNumDefautsPourUneCuve  'contruction des listes de défauts
            EnregistreDateDetectionDisparitionDefaut EtatDefaut, NumDefaut 'enregistrement de la date de détection et disparition du défaut
        
            '-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
        
        End If

        '--- construction de la liste des numéros de défauts pour une cuve ayant une charge ---
        If NumChargeACePoste = 0 Then
            .ListeNumDefautsSiCharge = ""    'vider la liste si pas de charge dans le poste
        Else
            If ListeNumDefautsPourUneCuve <> "" Then  'remplir la liste en évitant les doublons
                                                                                     'suppression du dernier séparateur inutile car dans fonction
                .ListeNumDefautsSiCharge = AjoutNumDefautsSansDoublons(.ListeNumDefautsSiCharge, _
                                                                                                                     ListeNumDefautsPourUneCuve) 'remplir la liste en éliminant les doublons
            End If
        End If

        '--- affectation définitive ---
        .UnDefautAuMoinsSignale = UnDefautAuMoinsSignale

    End With
    
    '*********************************************************************************************************************
    '*                                                                CUVE C16
    '*********************************************************************************************************************
    If False Then
    '--- affectation ---
    'NumCuve = CUVES_REGULATION.C_C16
    ListeNumDefautsPourUneCuve = ""
    
    '--- recherche de la présence d'une charge à ce poste ---
    NumChargeACePoste = TEtatsPostes(CorrespondanceCuvesAPIPostes(NumCuve)).NumCharge

    With TEtatsCuves(NumCuve)

        '--- initialisation à FALSE par défaut ---
        UnDefautAuMoinsSignale = False

        If TEtatsLigne.MarcheGenerale = True Then

            '-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

            '--- niveau trop bas ---
            NumDefaut = .TNumDefauts.NumDefautNiveauTresBas
            EtatDefaut = .TEntreesAPI.E_NiveauTresBas
            UnDefautAuMoinsSignale = UnDefautAuMoinsSignale Or EtatDefaut
            ConstruitListeNumDefauts EtatDefaut, NumDefaut, ListeNumDefauts, , ListeNumDefautsPourUneCuve  'contruction des listes de défauts
            EnregistreDateDetectionDisparitionDefaut EtatDefaut, NumDefaut 'enregistrement de la date de détection et disparition du défaut

            '--- niveau trop haut ---
            NumDefaut = .TNumDefauts.NumDefautNiveauTresHaut
            EtatDefaut = .TEntreesAPI.E_NiveauTresHaut
            UnDefautAuMoinsSignale = UnDefautAuMoinsSignale Or EtatDefaut
            ConstruitListeNumDefauts EtatDefaut, NumDefaut, ListeNumDefauts, , ListeNumDefautsPourUneCuve  'contruction des listes de défauts
            EnregistreDateDetectionDisparitionDefaut EtatDefaut, NumDefaut 'enregistrement de la date de détection et disparition du défaut

            '-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

            '--- défaut du chauffage ---
            NumDefaut = .TNumDefauts.NumDefautDefautChauffage
            EtatDefaut = .TEntreesAPI.E_DefautChauffage
            UnDefautAuMoinsSignale = UnDefautAuMoinsSignale Or EtatDefaut
            ConstruitListeNumDefauts EtatDefaut, NumDefaut, ListeNumDefauts, , ListeNumDefautsPourUneCuve  'contruction des listes de défauts
            EnregistreDateDetectionDisparitionDefaut EtatDefaut, NumDefaut 'enregistrement de la date de détection et disparition du défaut
            
            '-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
            
            '--- défaut PT100 ---
            NumDefaut = .TNumDefauts.NumDefautDefautPT100
            EtatDefaut = .TEntreesAPI.DefautPT100
            UnDefautAuMoinsSignale = UnDefautAuMoinsSignale Or EtatDefaut
            ConstruitListeNumDefauts EtatDefaut, NumDefaut, ListeNumDefauts, , ListeNumDefautsPourUneCuve  'contruction des listes de défauts
            EnregistreDateDetectionDisparitionDefaut EtatDefaut, NumDefaut 'enregistrement de la date de détection et disparition du défaut
            
            '--- température trop basse ---
            NumDefaut = .TNumDefauts.NumDefautTemperatureTropBasse
            EtatDefaut = .TEntreesAPI.TemperatureTropBasse
            UnDefautAuMoinsSignale = UnDefautAuMoinsSignale Or EtatDefaut
            ConstruitListeNumDefauts EtatDefaut, NumDefaut, ListeNumDefauts, , ListeNumDefautsPourUneCuve  'contruction des listes de défauts
            EnregistreDateDetectionDisparitionDefaut EtatDefaut, NumDefaut 'enregistrement de la date de détection et disparition du défaut
            
            '--- température trop haute ---
            NumDefaut = .TNumDefauts.NumDefautTemperatureTropHaute
            EtatDefaut = .TEntreesAPI.TemperatureTropHaute
            UnDefautAuMoinsSignale = UnDefautAuMoinsSignale Or EtatDefaut
            ConstruitListeNumDefauts EtatDefaut, NumDefaut, ListeNumDefauts, , ListeNumDefautsPourUneCuve  'contruction des listes de défauts
            EnregistreDateDetectionDisparitionDefaut EtatDefaut, NumDefaut 'enregistrement de la date de détection et disparition du défaut

            '-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

            '--- défaut refroidissement ---
            NumDefaut = .TNumDefauts.NumDefautDefautRefroidissement
            EtatDefaut = .TEntreesAPI.E_DefautRefroidissement
            UnDefautAuMoinsSignale = UnDefautAuMoinsSignale Or EtatDefaut
            ConstruitListeNumDefauts EtatDefaut, NumDefaut, ListeNumDefauts, , ListeNumDefautsPourUneCuve  'contruction des listes de défauts
            EnregistreDateDetectionDisparitionDefaut EtatDefaut, NumDefaut 'enregistrement de la date de détection et disparition du défaut
        
            '-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
        
        End If

        '--- construction de la liste des numéros de défauts pour une cuve ayant une charge ---
        If NumChargeACePoste = 0 Then
            .ListeNumDefautsSiCharge = ""    'vider la liste si pas de charge dans le poste
        Else
            If ListeNumDefautsPourUneCuve <> "" Then  'remplir la liste en évitant les doublons
                                                                                     'suppression du dernier séparateur inutile car dans fonction
                .ListeNumDefautsSiCharge = AjoutNumDefautsSansDoublons(.ListeNumDefautsSiCharge, _
                                                                                                                     ListeNumDefautsPourUneCuve) 'remplir la liste en éliminant les doublons
            End If
        End If

        '--- affectation définitive ---
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
    
    '--- recherche de la présence d'une charge à ce poste ---
    NumChargeACePoste = TEtatsPostes(CorrespondanceCuvesAPIPostes(NumCuve)).NumCharge

    With TEtatsCuves(NumCuve)

        '--- initialisation à FALSE par défaut ---
        UnDefautAuMoinsSignale = False

        If TEtatsLigne.MarcheGenerale = True Then
            
            '-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
            
            '--- niveau trop bas ---
            NumDefaut = .TNumDefauts.NumDefautNiveauTresBas
            EtatDefaut = .TEntreesAPI.E_NiveauTresBas
            UnDefautAuMoinsSignale = UnDefautAuMoinsSignale Or EtatDefaut
            ConstruitListeNumDefauts EtatDefaut, NumDefaut, ListeNumDefauts, , ListeNumDefautsPourUneCuve  'contruction des listes de défauts
            EnregistreDateDetectionDisparitionDefaut EtatDefaut, NumDefaut 'enregistrement de la date de détection et disparition du défaut

            '--- niveau trop haut ---
            NumDefaut = .TNumDefauts.NumDefautNiveauTresHaut
            EtatDefaut = .TEntreesAPI.E_NiveauTresHaut
            UnDefautAuMoinsSignale = UnDefautAuMoinsSignale Or EtatDefaut
            ConstruitListeNumDefauts EtatDefaut, NumDefaut, ListeNumDefauts, , ListeNumDefautsPourUneCuve  'contruction des listes de défauts
            EnregistreDateDetectionDisparitionDefaut EtatDefaut, NumDefaut 'enregistrement de la date de détection et disparition du défaut

            '-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

            '--- défaut du chauffage ---
            NumDefaut = .TNumDefauts.NumDefautDefautChauffage
            EtatDefaut = .TEntreesAPI.E_DefautChauffage
            UnDefautAuMoinsSignale = UnDefautAuMoinsSignale Or EtatDefaut
            ConstruitListeNumDefauts EtatDefaut, NumDefaut, ListeNumDefauts, , ListeNumDefautsPourUneCuve  'contruction des listes de défauts
            EnregistreDateDetectionDisparitionDefaut EtatDefaut, NumDefaut 'enregistrement de la date de détection et disparition du défaut
            
            '-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
            
            '--- défaut PT100 ---
            NumDefaut = .TNumDefauts.NumDefautDefautPT100
            EtatDefaut = .TEntreesAPI.DefautPT100
            UnDefautAuMoinsSignale = UnDefautAuMoinsSignale Or EtatDefaut
            ConstruitListeNumDefauts EtatDefaut, NumDefaut, ListeNumDefauts, , ListeNumDefautsPourUneCuve  'contruction des listes de défauts
            EnregistreDateDetectionDisparitionDefaut EtatDefaut, NumDefaut 'enregistrement de la date de détection et disparition du défaut
            
            '--- température trop basse ---
            NumDefaut = .TNumDefauts.NumDefautTemperatureTropBasse
            EtatDefaut = .TEntreesAPI.TemperatureTropBasse
            UnDefautAuMoinsSignale = UnDefautAuMoinsSignale Or EtatDefaut
            ConstruitListeNumDefauts EtatDefaut, NumDefaut, ListeNumDefauts, , ListeNumDefautsPourUneCuve  'contruction des listes de défauts
            EnregistreDateDetectionDisparitionDefaut EtatDefaut, NumDefaut 'enregistrement de la date de détection et disparition du défaut
            
            '--- température trop haute ---
            NumDefaut = .TNumDefauts.NumDefautTemperatureTropHaute
            EtatDefaut = .TEntreesAPI.TemperatureTropHaute
            UnDefautAuMoinsSignale = UnDefautAuMoinsSignale Or EtatDefaut
            ConstruitListeNumDefauts EtatDefaut, NumDefaut, ListeNumDefauts, , ListeNumDefautsPourUneCuve  'contruction des listes de défauts
            EnregistreDateDetectionDisparitionDefaut EtatDefaut, NumDefaut 'enregistrement de la date de détection et disparition du défaut

            '-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

        End If

        '--- construction de la liste des numéros de défauts pour une cuve ayant une charge ---
        If NumChargeACePoste = 0 Then
            .ListeNumDefautsSiCharge = ""    'vider la liste si pas de charge dans le poste
        Else
            If ListeNumDefautsPourUneCuve <> "" Then  'remplir la liste en évitant les doublons
                                                                                     'suppression du dernier séparateur inutile car dans fonction
                .ListeNumDefautsSiCharge = AjoutNumDefautsSansDoublons(.ListeNumDefautsSiCharge, _
                                                                                                                     ListeNumDefautsPourUneCuve) 'remplir la liste en éliminant les doublons
            End If
        End If

        '--- affectation définitive ---
        .UnDefautAuMoinsSignale = UnDefautAuMoinsSignale

    End With
    End If
   
    
    '*********************************************************************************************************************
    '*                                                                CUVE C22
    '*********************************************************************************************************************
    
    '--- affectation ---
    NumCuve = CUVES_REGULATION.C_C22
    ListeNumDefautsPourUneCuve = ""
    
    '--- recherche de la présence d'une charge à ce poste ---
    NumChargeACePoste = TEtatsPostes(CorrespondanceCuvesAPIPostes(NumCuve)).NumCharge

    With TEtatsCuves(NumCuve)

        '--- initialisation à FALSE par défaut ---
        UnDefautAuMoinsSignale = False

        If TEtatsLigne.MarcheGenerale = True Then
            
            '-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
            
            '--- niveau trop bas ---
            NumDefaut = .TNumDefauts.NumDefautNiveauTresBas
            EtatDefaut = .TEntreesAPI.E_NiveauTresBas
            UnDefautAuMoinsSignale = UnDefautAuMoinsSignale Or EtatDefaut
            ConstruitListeNumDefauts EtatDefaut, NumDefaut, ListeNumDefauts, , ListeNumDefautsPourUneCuve  'contruction des listes de défauts
            EnregistreDateDetectionDisparitionDefaut EtatDefaut, NumDefaut 'enregistrement de la date de détection et disparition du défaut

            '--- niveau trop haut ---
            NumDefaut = .TNumDefauts.NumDefautNiveauTresHaut
            EtatDefaut = .TEntreesAPI.E_NiveauTresHaut
            UnDefautAuMoinsSignale = UnDefautAuMoinsSignale Or EtatDefaut
            ConstruitListeNumDefauts EtatDefaut, NumDefaut, ListeNumDefauts, , ListeNumDefautsPourUneCuve  'contruction des listes de défauts
            EnregistreDateDetectionDisparitionDefaut EtatDefaut, NumDefaut 'enregistrement de la date de détection et disparition du défaut

            '-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

            '--- défaut du chauffage ---
            NumDefaut = .TNumDefauts.NumDefautDefautChauffage
            EtatDefaut = .TEntreesAPI.E_DefautChauffage
            UnDefautAuMoinsSignale = UnDefautAuMoinsSignale Or EtatDefaut
            ConstruitListeNumDefauts EtatDefaut, NumDefaut, ListeNumDefauts, , ListeNumDefautsPourUneCuve  'contruction des listes de défauts
            EnregistreDateDetectionDisparitionDefaut EtatDefaut, NumDefaut 'enregistrement de la date de détection et disparition du défaut
            
            '-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
            
            '--- défaut PT100 ---
            NumDefaut = .TNumDefauts.NumDefautDefautPT100
            EtatDefaut = .TEntreesAPI.DefautPT100
            UnDefautAuMoinsSignale = UnDefautAuMoinsSignale Or EtatDefaut
            ConstruitListeNumDefauts EtatDefaut, NumDefaut, ListeNumDefauts, , ListeNumDefautsPourUneCuve  'contruction des listes de défauts
            EnregistreDateDetectionDisparitionDefaut EtatDefaut, NumDefaut 'enregistrement de la date de détection et disparition du défaut
            
            '--- température trop basse ---
            NumDefaut = .TNumDefauts.NumDefautTemperatureTropBasse
            EtatDefaut = .TEntreesAPI.TemperatureTropBasse
            UnDefautAuMoinsSignale = UnDefautAuMoinsSignale Or EtatDefaut
            ConstruitListeNumDefauts EtatDefaut, NumDefaut, ListeNumDefauts, , ListeNumDefautsPourUneCuve  'contruction des listes de défauts
            EnregistreDateDetectionDisparitionDefaut EtatDefaut, NumDefaut 'enregistrement de la date de détection et disparition du défaut
            
            '--- température trop haute ---
            NumDefaut = .TNumDefauts.NumDefautTemperatureTropHaute
            EtatDefaut = .TEntreesAPI.TemperatureTropHaute
            UnDefautAuMoinsSignale = UnDefautAuMoinsSignale Or EtatDefaut
            ConstruitListeNumDefauts EtatDefaut, NumDefaut, ListeNumDefauts, , ListeNumDefautsPourUneCuve  'contruction des listes de défauts
            EnregistreDateDetectionDisparitionDefaut EtatDefaut, NumDefaut 'enregistrement de la date de détection et disparition du défaut

            '-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

        End If

        '--- construction de la liste des numéros de défauts pour une cuve ayant une charge ---
        If NumChargeACePoste = 0 Then
            .ListeNumDefautsSiCharge = ""    'vider la liste si pas de charge dans le poste
        Else
            If ListeNumDefautsPourUneCuve <> "" Then  'remplir la liste en évitant les doublons
                                                                                     'suppression du dernier séparateur inutile car dans fonction
                .ListeNumDefautsSiCharge = AjoutNumDefautsSansDoublons(.ListeNumDefautsSiCharge, _
                                                                                                                     ListeNumDefautsPourUneCuve) 'remplir la liste en éliminant les doublons
            End If
        End If

        '--- affectation définitive ---
        .UnDefautAuMoinsSignale = UnDefautAuMoinsSignale

    End With
    
    '*********************************************************************************************************************
    '*                                                                CUVE C27
    '*********************************************************************************************************************
    
    '--- affectation ---
    NumCuve = CUVES_REGULATION.C_C27
    ListeNumDefautsPourUneCuve = ""
    
    '--- recherche de la présence d'une charge à ce poste ---
    NumChargeACePoste = TEtatsPostes(CorrespondanceCuvesAPIPostes(NumCuve)).NumCharge

    With TEtatsCuves(NumCuve)

        '--- initialisation à FALSE par défaut ---
        UnDefautAuMoinsSignale = False

        If TEtatsLigne.MarcheGenerale = True Then

            '-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

            '--- niveau trop bas ---
            NumDefaut = .TNumDefauts.NumDefautNiveauTresBas
            EtatDefaut = .TEntreesAPI.E_NiveauTresBas
            UnDefautAuMoinsSignale = UnDefautAuMoinsSignale Or EtatDefaut
            ConstruitListeNumDefauts EtatDefaut, NumDefaut, ListeNumDefauts, , ListeNumDefautsPourUneCuve  'contruction des listes de défauts
            EnregistreDateDetectionDisparitionDefaut EtatDefaut, NumDefaut 'enregistrement de la date de détection et disparition du défaut

            '--- niveau trop haut ---
            NumDefaut = .TNumDefauts.NumDefautNiveauTresHaut
            EtatDefaut = .TEntreesAPI.E_NiveauTresHaut
            UnDefautAuMoinsSignale = UnDefautAuMoinsSignale Or EtatDefaut
            ConstruitListeNumDefauts EtatDefaut, NumDefaut, ListeNumDefauts, , ListeNumDefautsPourUneCuve  'contruction des listes de défauts
            EnregistreDateDetectionDisparitionDefaut EtatDefaut, NumDefaut 'enregistrement de la date de détection et disparition du défaut

            '-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

            '--- défaut du chauffage ---
            NumDefaut = .TNumDefauts.NumDefautDefautChauffage
            EtatDefaut = .TEntreesAPI.E_DefautChauffage
            UnDefautAuMoinsSignale = UnDefautAuMoinsSignale Or EtatDefaut
            ConstruitListeNumDefauts EtatDefaut, NumDefaut, ListeNumDefauts, , ListeNumDefautsPourUneCuve  'contruction des listes de défauts
            EnregistreDateDetectionDisparitionDefaut EtatDefaut, NumDefaut 'enregistrement de la date de détection et disparition du défaut
            
            '-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
            
            '--- défaut PT100 ---
            NumDefaut = .TNumDefauts.NumDefautDefautPT100
            EtatDefaut = .TEntreesAPI.DefautPT100
            UnDefautAuMoinsSignale = UnDefautAuMoinsSignale Or EtatDefaut
            ConstruitListeNumDefauts EtatDefaut, NumDefaut, ListeNumDefauts, , ListeNumDefautsPourUneCuve  'contruction des listes de défauts
            EnregistreDateDetectionDisparitionDefaut EtatDefaut, NumDefaut 'enregistrement de la date de détection et disparition du défaut
            
            '--- température trop basse ---
            NumDefaut = .TNumDefauts.NumDefautTemperatureTropBasse
            EtatDefaut = .TEntreesAPI.TemperatureTropBasse
            UnDefautAuMoinsSignale = UnDefautAuMoinsSignale Or EtatDefaut
            ConstruitListeNumDefauts EtatDefaut, NumDefaut, ListeNumDefauts, , ListeNumDefautsPourUneCuve  'contruction des listes de défauts
            EnregistreDateDetectionDisparitionDefaut EtatDefaut, NumDefaut 'enregistrement de la date de détection et disparition du défaut
            
            '--- température trop haute ---
            NumDefaut = .TNumDefauts.NumDefautTemperatureTropHaute
            EtatDefaut = .TEntreesAPI.TemperatureTropHaute
            UnDefautAuMoinsSignale = UnDefautAuMoinsSignale Or EtatDefaut
            ConstruitListeNumDefauts EtatDefaut, NumDefaut, ListeNumDefauts, , ListeNumDefautsPourUneCuve  'contruction des listes de défauts
            EnregistreDateDetectionDisparitionDefaut EtatDefaut, NumDefaut 'enregistrement de la date de détection et disparition du défaut

            '-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

        End If

        '--- construction de la liste des numéros de défauts pour une cuve ayant une charge ---
        If NumChargeACePoste = 0 Then
            .ListeNumDefautsSiCharge = ""    'vider la liste si pas de charge dans le poste
        Else
            If ListeNumDefautsPourUneCuve <> "" Then  'remplir la liste en évitant les doublons
                                                                                     'suppression du dernier séparateur inutile car dans fonction
                .ListeNumDefautsSiCharge = AjoutNumDefautsSansDoublons(.ListeNumDefautsSiCharge, _
                                                                                                                     ListeNumDefautsPourUneCuve) 'remplir la liste en éliminant les doublons
            End If
        End If

        '--- affectation définitive ---
        .UnDefautAuMoinsSignale = UnDefautAuMoinsSignale

    End With
    
    '*********************************************************************************************************************
    '*                                                                CUVE C28
    '*********************************************************************************************************************
    
    '--- affectation ---
    NumCuve = CUVES_REGULATION.C_C28
    ListeNumDefautsPourUneCuve = ""
    
    '--- recherche de la présence d'une charge à ce poste ---
    NumChargeACePoste = TEtatsPostes(CorrespondanceCuvesAPIPostes(NumCuve)).NumCharge

    With TEtatsCuves(NumCuve)

        '--- initialisation à FALSE par défaut ---
        UnDefautAuMoinsSignale = False

        If TEtatsLigne.MarcheGenerale = True Then

            '-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

            '--- niveau trop bas ---
            NumDefaut = .TNumDefauts.NumDefautNiveauTresBas
            EtatDefaut = .TEntreesAPI.E_NiveauTresBas
            UnDefautAuMoinsSignale = UnDefautAuMoinsSignale Or EtatDefaut
            ConstruitListeNumDefauts EtatDefaut, NumDefaut, ListeNumDefauts, , ListeNumDefautsPourUneCuve  'contruction des listes de défauts
            EnregistreDateDetectionDisparitionDefaut EtatDefaut, NumDefaut 'enregistrement de la date de détection et disparition du défaut

            '--- niveau trop haut ---
            NumDefaut = .TNumDefauts.NumDefautNiveauTresHaut
            EtatDefaut = .TEntreesAPI.E_NiveauTresHaut
            UnDefautAuMoinsSignale = UnDefautAuMoinsSignale Or EtatDefaut
            ConstruitListeNumDefauts EtatDefaut, NumDefaut, ListeNumDefauts, , ListeNumDefautsPourUneCuve  'contruction des listes de défauts
            EnregistreDateDetectionDisparitionDefaut EtatDefaut, NumDefaut 'enregistrement de la date de détection et disparition du défaut

            '-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

            '--- défaut du chauffage ---
            NumDefaut = .TNumDefauts.NumDefautDefautChauffage
            EtatDefaut = .TEntreesAPI.E_DefautChauffage
            UnDefautAuMoinsSignale = UnDefautAuMoinsSignale Or EtatDefaut
            ConstruitListeNumDefauts EtatDefaut, NumDefaut, ListeNumDefauts, , ListeNumDefautsPourUneCuve  'contruction des listes de défauts
            EnregistreDateDetectionDisparitionDefaut EtatDefaut, NumDefaut 'enregistrement de la date de détection et disparition du défaut
            
            '-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
            
            '--- défaut PT100 ---
            NumDefaut = .TNumDefauts.NumDefautDefautPT100
            EtatDefaut = .TEntreesAPI.DefautPT100
            UnDefautAuMoinsSignale = UnDefautAuMoinsSignale Or EtatDefaut
            ConstruitListeNumDefauts EtatDefaut, NumDefaut, ListeNumDefauts, , ListeNumDefautsPourUneCuve  'contruction des listes de défauts
            EnregistreDateDetectionDisparitionDefaut EtatDefaut, NumDefaut 'enregistrement de la date de détection et disparition du défaut
            
            '--- température trop basse ---
            NumDefaut = .TNumDefauts.NumDefautTemperatureTropBasse
            EtatDefaut = .TEntreesAPI.TemperatureTropBasse
            UnDefautAuMoinsSignale = UnDefautAuMoinsSignale Or EtatDefaut
            ConstruitListeNumDefauts EtatDefaut, NumDefaut, ListeNumDefauts, , ListeNumDefautsPourUneCuve  'contruction des listes de défauts
            EnregistreDateDetectionDisparitionDefaut EtatDefaut, NumDefaut 'enregistrement de la date de détection et disparition du défaut
            
            '--- température trop haute ---
            NumDefaut = .TNumDefauts.NumDefautTemperatureTropHaute
            EtatDefaut = .TEntreesAPI.TemperatureTropHaute
            UnDefautAuMoinsSignale = UnDefautAuMoinsSignale Or EtatDefaut
            ConstruitListeNumDefauts EtatDefaut, NumDefaut, ListeNumDefauts, , ListeNumDefautsPourUneCuve  'contruction des listes de défauts
            EnregistreDateDetectionDisparitionDefaut EtatDefaut, NumDefaut 'enregistrement de la date de détection et disparition du défaut

            '-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

        End If

        '--- construction de la liste des numéros de défauts pour une cuve ayant une charge ---
        If NumChargeACePoste = 0 Then
            .ListeNumDefautsSiCharge = ""    'vider la liste si pas de charge dans le poste
        Else
            If ListeNumDefautsPourUneCuve <> "" Then  'remplir la liste en évitant les doublons
                                                                                     'suppression du dernier séparateur inutile car dans fonction
                .ListeNumDefautsSiCharge = AjoutNumDefautsSansDoublons(.ListeNumDefautsSiCharge, _
                                                                                                                     ListeNumDefautsPourUneCuve) 'remplir la liste en éliminant les doublons
            End If
        End If

        '--- affectation définitive ---
        .UnDefautAuMoinsSignale = UnDefautAuMoinsSignale

    End With
    
    '*********************************************************************************************************************
    '*                                                                CUVE C31
    '*********************************************************************************************************************
    
    '--- affectation ---
    NumCuve = CUVES_REGULATION.C_C31
    ListeNumDefautsPourUneCuve = ""
    
    '--- recherche de la présence d'une charge à ce poste ---
    NumChargeACePoste = TEtatsPostes(CorrespondanceCuvesAPIPostes(NumCuve)).NumCharge

    With TEtatsCuves(NumCuve)

        '--- initialisation à FALSE par défaut ---
        UnDefautAuMoinsSignale = False

        If TEtatsLigne.MarcheGenerale = True Then

            '-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
            
            '--- niveau trop bas ---
            NumDefaut = .TNumDefauts.NumDefautNiveauTresBas
            EtatDefaut = .TEntreesAPI.E_NiveauTresBas
            UnDefautAuMoinsSignale = UnDefautAuMoinsSignale Or EtatDefaut
            ConstruitListeNumDefauts EtatDefaut, NumDefaut, ListeNumDefauts, , ListeNumDefautsPourUneCuve  'contruction des listes de défauts
            EnregistreDateDetectionDisparitionDefaut EtatDefaut, NumDefaut 'enregistrement de la date de détection et disparition du défaut

            '--- niveau trop haut ---
            NumDefaut = .TNumDefauts.NumDefautNiveauTresHaut
            EtatDefaut = .TEntreesAPI.E_NiveauTresHaut
            UnDefautAuMoinsSignale = UnDefautAuMoinsSignale Or EtatDefaut
            ConstruitListeNumDefauts EtatDefaut, NumDefaut, ListeNumDefauts, , ListeNumDefautsPourUneCuve  'contruction des listes de défauts
            EnregistreDateDetectionDisparitionDefaut EtatDefaut, NumDefaut 'enregistrement de la date de détection et disparition du défaut

            '-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

            '--- défaut du chauffage ---
            NumDefaut = .TNumDefauts.NumDefautDefautChauffage
            EtatDefaut = .TEntreesAPI.E_DefautChauffage
            UnDefautAuMoinsSignale = UnDefautAuMoinsSignale Or EtatDefaut
            ConstruitListeNumDefauts EtatDefaut, NumDefaut, ListeNumDefauts, , ListeNumDefautsPourUneCuve  'contruction des listes de défauts
            EnregistreDateDetectionDisparitionDefaut EtatDefaut, NumDefaut 'enregistrement de la date de détection et disparition du défaut
            
            '-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
            
            '--- défaut PT100 ---
            NumDefaut = .TNumDefauts.NumDefautDefautPT100
            EtatDefaut = .TEntreesAPI.DefautPT100
            UnDefautAuMoinsSignale = UnDefautAuMoinsSignale Or EtatDefaut
            ConstruitListeNumDefauts EtatDefaut, NumDefaut, ListeNumDefauts, , ListeNumDefautsPourUneCuve  'contruction des listes de défauts
            EnregistreDateDetectionDisparitionDefaut EtatDefaut, NumDefaut 'enregistrement de la date de détection et disparition du défaut
            
            '--- température trop basse ---
            NumDefaut = .TNumDefauts.NumDefautTemperatureTropBasse
            EtatDefaut = .TEntreesAPI.TemperatureTropBasse
            UnDefautAuMoinsSignale = UnDefautAuMoinsSignale Or EtatDefaut
            ConstruitListeNumDefauts EtatDefaut, NumDefaut, ListeNumDefauts, , ListeNumDefautsPourUneCuve  'contruction des listes de défauts
            EnregistreDateDetectionDisparitionDefaut EtatDefaut, NumDefaut 'enregistrement de la date de détection et disparition du défaut
            
            '--- température trop haute ---
            NumDefaut = .TNumDefauts.NumDefautTemperatureTropHaute
            EtatDefaut = .TEntreesAPI.TemperatureTropHaute
            UnDefautAuMoinsSignale = UnDefautAuMoinsSignale Or EtatDefaut
            ConstruitListeNumDefauts EtatDefaut, NumDefaut, ListeNumDefauts, , ListeNumDefautsPourUneCuve  'contruction des listes de défauts
            EnregistreDateDetectionDisparitionDefaut EtatDefaut, NumDefaut 'enregistrement de la date de détection et disparition du défaut

            '-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

        End If

        '--- construction de la liste des numéros de défauts pour une cuve ayant une charge ---
        If NumChargeACePoste = 0 Then
            .ListeNumDefautsSiCharge = ""    'vider la liste si pas de charge dans le poste
        Else
            If ListeNumDefautsPourUneCuve <> "" Then  'remplir la liste en évitant les doublons
                                                                                     'suppression du dernier séparateur inutile car dans fonction
                .ListeNumDefautsSiCharge = AjoutNumDefautsSansDoublons(.ListeNumDefautsSiCharge, _
                                                                                                                     ListeNumDefautsPourUneCuve) 'remplir la liste en éliminant les doublons
            End If
        End If

        '--- affectation définitive ---
        .UnDefautAuMoinsSignale = UnDefautAuMoinsSignale

    End With
    
    '*********************************************************************************************************************
    '*                                                                CUVE C32
    '*********************************************************************************************************************
    
    '--- affectation ---
    NumCuve = CUVES_REGULATION.C_C32
    ListeNumDefautsPourUneCuve = ""
    
    '--- recherche de la présence d'une charge à ce poste ---
    NumChargeACePoste = TEtatsPostes(CorrespondanceCuvesAPIPostes(NumCuve)).NumCharge

    With TEtatsCuves(NumCuve)

        '--- initialisation à FALSE par défaut ---
        UnDefautAuMoinsSignale = False

        If TEtatsLigne.MarcheGenerale = True Then

            '-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

            '--- niveau trop bas ---
            NumDefaut = .TNumDefauts.NumDefautNiveauTresBas
            EtatDefaut = .TEntreesAPI.E_NiveauTresBas
            UnDefautAuMoinsSignale = UnDefautAuMoinsSignale Or EtatDefaut
            ConstruitListeNumDefauts EtatDefaut, NumDefaut, ListeNumDefauts, , ListeNumDefautsPourUneCuve  'contruction des listes de défauts
            EnregistreDateDetectionDisparitionDefaut EtatDefaut, NumDefaut 'enregistrement de la date de détection et disparition du défaut

            '--- niveau trop haut ---
            NumDefaut = .TNumDefauts.NumDefautNiveauTresHaut
            EtatDefaut = .TEntreesAPI.E_NiveauTresHaut
            UnDefautAuMoinsSignale = UnDefautAuMoinsSignale Or EtatDefaut
            ConstruitListeNumDefauts EtatDefaut, NumDefaut, ListeNumDefauts, , ListeNumDefautsPourUneCuve  'contruction des listes de défauts
            EnregistreDateDetectionDisparitionDefaut EtatDefaut, NumDefaut 'enregistrement de la date de détection et disparition du défaut

            '-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

            '--- défaut du chauffage ---
            NumDefaut = .TNumDefauts.NumDefautDefautChauffage
            EtatDefaut = .TEntreesAPI.E_DefautChauffage
            UnDefautAuMoinsSignale = UnDefautAuMoinsSignale Or EtatDefaut
            ConstruitListeNumDefauts EtatDefaut, NumDefaut, ListeNumDefauts, , ListeNumDefautsPourUneCuve  'contruction des listes de défauts
            EnregistreDateDetectionDisparitionDefaut EtatDefaut, NumDefaut 'enregistrement de la date de détection et disparition du défaut
            
            '-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
            
            '--- défaut PT100 ---
            NumDefaut = .TNumDefauts.NumDefautDefautPT100
            EtatDefaut = .TEntreesAPI.DefautPT100
            UnDefautAuMoinsSignale = UnDefautAuMoinsSignale Or EtatDefaut
            ConstruitListeNumDefauts EtatDefaut, NumDefaut, ListeNumDefauts, , ListeNumDefautsPourUneCuve  'contruction des listes de défauts
            EnregistreDateDetectionDisparitionDefaut EtatDefaut, NumDefaut 'enregistrement de la date de détection et disparition du défaut
            
            '--- température trop basse ---
            NumDefaut = .TNumDefauts.NumDefautTemperatureTropBasse
            EtatDefaut = .TEntreesAPI.TemperatureTropBasse
            UnDefautAuMoinsSignale = UnDefautAuMoinsSignale Or EtatDefaut
            ConstruitListeNumDefauts EtatDefaut, NumDefaut, ListeNumDefauts, , ListeNumDefautsPourUneCuve  'contruction des listes de défauts
            EnregistreDateDetectionDisparitionDefaut EtatDefaut, NumDefaut 'enregistrement de la date de détection et disparition du défaut
            
            '--- température trop haute ---
            NumDefaut = .TNumDefauts.NumDefautTemperatureTropHaute
            EtatDefaut = .TEntreesAPI.TemperatureTropHaute
            UnDefautAuMoinsSignale = UnDefautAuMoinsSignale Or EtatDefaut
            ConstruitListeNumDefauts EtatDefaut, NumDefaut, ListeNumDefauts, , ListeNumDefautsPourUneCuve  'contruction des listes de défauts
            EnregistreDateDetectionDisparitionDefaut EtatDefaut, NumDefaut 'enregistrement de la date de détection et disparition du défaut

            '-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

        End If

        '--- construction de la liste des numéros de défauts pour une cuve ayant une charge ---
        If NumChargeACePoste = 0 Then
            .ListeNumDefautsSiCharge = ""    'vider la liste si pas de charge dans le poste
        Else
            If ListeNumDefautsPourUneCuve <> "" Then  'remplir la liste en évitant les doublons
                                                                                     'suppression du dernier séparateur inutile car dans fonction
                .ListeNumDefautsSiCharge = AjoutNumDefautsSansDoublons(.ListeNumDefautsSiCharge, _
                                                                                                                     ListeNumDefautsPourUneCuve) 'remplir la liste en éliminant les doublons
            End If
        End If

        '--- affectation définitive ---
        .UnDefautAuMoinsSignale = UnDefautAuMoinsSignale

    End With

    '*********************************************************************************************************************
    '*                                                                  REDRESSEURS
    '*********************************************************************************************************************
    
    For a = LBound(TEtatsRedresseurs()) To UBound(TEtatsRedresseurs())
        
        With TEtatsRedresseurs(a)

            '--- défaut général ---
            NumDefaut = .TNumDefauts.NumDefautDefautGeneral
            EtatDefaut = .TEntreesAPI.M_DefautGeneral
            UnDefautAuMoinsSignale = UnDefautAuMoinsSignale Or EtatDefaut
            ConstruitListeNumDefauts EtatDefaut, NumDefaut, ListeNumDefauts, , ListeNumDefautsPourUneCuve  'contruction des listes de défauts
            EnregistreDateDetectionDisparitionDefaut EtatDefaut, NumDefaut 'enregistrement de la date de détection et disparition du défaut
            
            '--- délai trop long de mise en marche ---
            NumDefaut = .TNumDefauts.NumDefautDelaiTropLongMarcheRedresseur
            EtatDefaut = .TEntreesAPI.M_DelaiTropLongMarcheRedresseur
            UnDefautAuMoinsSignale = UnDefautAuMoinsSignale Or EtatDefaut
            ConstruitListeNumDefauts EtatDefaut, NumDefaut, ListeNumDefauts, , ListeNumDefautsPourUneCuve  'contruction des listes de défauts
            EnregistreDateDetectionDisparitionDefaut EtatDefaut, NumDefaut 'enregistrement de la date de détection et disparition du défaut
                            
            '--- intensité non atteinte ---
            NumDefaut = .TNumDefauts.NumDefautIntensiteNonAtteinte
            EtatDefaut = .TEntreesAPI.M_IntensiteNonAtteinte
            UnDefautAuMoinsSignale = UnDefautAuMoinsSignale Or EtatDefaut
            ConstruitListeNumDefauts EtatDefaut, NumDefaut, ListeNumDefauts, , ListeNumDefautsPourUneCuve  'contruction des listes de défauts
            EnregistreDateDetectionDisparitionDefaut EtatDefaut, NumDefaut 'enregistrement de la date de détection et disparition du défaut
                            
            '--- intensité instable ---
            NumDefaut = .TNumDefauts.NumDefautIntensiteInstable
            EtatDefaut = .TEntreesAPI.M_IntensiteInstable
            UnDefautAuMoinsSignale = UnDefautAuMoinsSignale Or EtatDefaut
            ConstruitListeNumDefauts EtatDefaut, NumDefaut, ListeNumDefauts, , ListeNumDefautsPourUneCuve  'contruction des listes de défauts
            EnregistreDateDetectionDisparitionDefaut EtatDefaut, NumDefaut 'enregistrement de la date de détection et disparition du défaut
    
        End With
    
    Next a
    
    '*********************************************************************************************************************
    '*                                                ANALYSE DES ALARMES LIGNE  EN COURS
    '*********************************************************************************************************************

    '--- affectation ---
    If ListeNumDefautsPourLaLigne = "" Then
        AlarmesLigneEnCours = ""
    Else
        AlarmesLigneEnCours = Left(ListeNumDefautsPourLaLigne, Pred(Len(ListeNumDefautsPourLaLigne))) 'suppression du dernier séparateur
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

            '--- construction du tableau contenant les numéros de défauts ---
            TNumDefauts = Split(ListeNumDefauts, SEPARATEUR_NUM_DEFAUTS)
            If CptDefauts >= UBound(TNumDefauts) Then 'ATTENTION la dernière valeur du tableau est
                                                                                      'automatiquement une chaine vide à cause
                                                                                      'du dernier séparateur ajouté
                CptDefauts = 0
                CptAppels = 0
            End If

            '--- affichage du défaut ---
            If IsNumeric(TNumDefauts(CptDefauts)) = True Then
                
                NumDefautAAfficher = CLng(TNumDefauts(CptDefauts))
                
                If NumDefautAAfficher >= DEFAUTS.NUM_MINI And NumDefautAAfficher <= DEFAUTS.NUM_MAXI Then
                    
                    '--- affectation du libellé complété du défaut (cas des numéros de défaut des variateurs) ---
                    LibelleCompleteDefaut = CompleteLibelleDefaut(NumDefautAAfficher, ComplementDefaut)
                    
                    '--- réaffectation du libellé complet ---
                    LibelleCompleteDefaut = "Défaut n° " & NumDefautAAfficher & " - " & LibelleCompleteDefaut
                    
                    '--- afffichage dans le champ concerné (écran principal) ---
                    If OccFPrincipale.LMessages.Caption <> LibelleCompleteDefaut Then
                        OccFPrincipale.LMessages.Caption = LibelleCompleteDefaut
                    End If
                    
                    '--- affichage sur l'afficheur à condition qu'il n'y ai pas de priorité pour les alertes---
                    If TDefauts(NumDefautAAfficher).AfficheurOuiNon = True And PrioriteAfficheurPourAlertes = False Then
                        Bidon = MessageAfficheur("B", TDefauts(NumDefautAAfficher).LibelleDefautAfficheur)
                    End If
                    
                End If
            
            End If
            Inc CptDefauts

        End If

    Else

        '--- passage en couleur verte / effacement du message sur l'afficheur à leds rouge ---
        If OccFPrincipale.LMessages.BackColor <> COULEURS.VERT_3 Then
            
            '--- passage en couleur verte ---
            OccFPrincipale.LMessages.BackColor = COULEURS.VERT_3
            OccFPrincipale.LMessages.ForeColor = COULEURS.NOIR
        
            '--- effacement du message sur l'afficheur à leds rouge ---
            Bidon = MessageAfficheur("B", "")
        
        End If

        '--- affectation ---
        CptAppels = 0
        CptDefauts = 0

    End If

    '--- contrôle du rafraichissement  ---
    Inc CptAppels
    If CptAppels > 5 Then CptAppels = 0

End Sub



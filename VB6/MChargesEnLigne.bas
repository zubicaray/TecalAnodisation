Attribute VB_Name = "MChargesEnLigne"
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Rôle                    : MODULE AIDANT A LA GESTION DES CHARGES EN LIGNE
' Nom                    : MChargesEnLigne.bas
' Date de création : 08/03/2011
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

'--- déclarations obligatoires ---
Option Explicit

'--- options générales ---
Option Base 1
DefVar A-Z

'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Rôle      : Vide une charge dans le tableau des états des charges
' Entrées : NumCharge -> Numéro de la charge à initialiser
' Retours :
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Sub InitialisationCharge(ByVal NumCharge As Integer)
    
    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next
    
    '--- déclaration ---
    Dim a As Integer
    Dim FicheVideGammesAnodisation As EnrGammesAnodisation, _
            FicheVideDetailsCharges As DetailsCharges, _
            FicheVideDetailsGammesAnodisation As EnrDetailsGammesAnodisation, _
            FicheVideDetailsFichesProduction As DetailsFichesProduction
    
    '--- analyse en fonction du PC ---
    'If TypePC <> TYPES_PC. Then Exit Sub

    '--- contrôle avant d'initialiser la charge ---
    If NumCharge >= CHARGES.C_NUM_MINI And NumCharge <= CHARGES.C_NUM_MAXI Then

        With TEtatsCharges(NumCharge)
            
            '--- date d'entrée en ligne ---
            .DateEntreeEnLigne = Empty
            
            '--- date d'arrivée au déchargement ---
            .DateArriveeAuDechargement = Empty
            
            '--- numéro de barre ---
            .NumBarre = 0
            
            '--- charge prioritaire ---
            .ChargePrioritaire = False                                'indique qu'il sagit  d'une charge prioritaire
                                                                                      'cette option est validé au chargement
            
            '--- délai supplémentaire de stabilisation de la charge ---
            .DelaiSupStabilisationChargeSecondes = 0
            
            '--- options 1 et 2 de la charge (vitesse de montée-descente, etc ...) ---
            .Options1 = 0
            .Options2 = 0
            
            '--- détails des charges ---
            For a = LBound(.TDetailsCharges()) To UBound(.TDetailsCharges())
                .TDetailsCharges(a) = FicheVideDetailsCharges
            Next a
        
            '--- gamme d'anodisation ---
            .TGammesAnodisation = FicheVideGammesAnodisation
            
            '--- pointeur de la zone de la gamme d'anodisation ---
            .PtrZoneGammeAnodisation = 0
    
            '--- U et I des phases ---
            For a = LBound(.TDetailsPhasesProduction()) To UBound(.TDetailsPhasesProduction())
                With .TDetailsPhasesProduction(a)
                    .UPhase = 0                               'pour renseigner
                    .IPhase = 0                                'le redresseur en entrant dans le bain
                End With
            Next a
            
            '--- nombre de postes traités (pour la fiche de production) ---
            .NbrPostesTraites = 0
            
            '--- fiche de production ---
            For a = LBound(.TDetailsFichesProduction()) To UBound(.TDetailsFichesProduction())
                .TDetailsFichesProduction(a) = FicheVideDetailsFichesProduction
            Next a
            
            '--- alarmes de la ligne (hors postes qui se trouvent dans les fiches de production) ---
            .AlarmesLigne = ""
        
        End With

    End If
    
End Sub

'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Rôle      : Transmet la valeur du mot des options pour une charge
' Entrées :                             NumCharge -> N° de la charge ou l'on veut envoyer les options
'                                                  Options1 -> Valeur de toutes les options 1
'                                                  Options2 -> Valeur de toutes les options 2
' Retours : EnvoiOptionsPourUneCharge -> OK = Transmission correcte
'                                                                       "" = Incident de transmission
' Détails  :
'                           Poids FORT du mot transmis OPTIONS 1
'                           ---------------------------------------------------------------------------------------
'                           |  Bit 7 |  Bit 6 | Bit 5 | Bit 4 | Bit 3 | Bit 2 | Bit 1 | Bit 0 |
'                           ---------------------------------------------------------------------------------------
'                           |  128   |   64   |   32  |   16   |    8   |    4    |    2   |     1   |
'                           ---------------------------------------------------------------------------------------
'                                 |           |          |         |         |          |          |         |_____  forcer la montée en très petite vitesse
'                                 |           |          |         |         |          |          |__________  forcer la montée en petite vitesse
'                                 |           |          |         |         |          |________________ forcer la descente en très petite vitesse
'                                 |           |          |         |         |_____________________  forcer la descente en petite vitesse
'                                 |           |          |         |__________________________
'                                 |           |          |_______________________________
'                                 |           |_____________________________________
'                                 |___________________________________________
'
'                           Poids FORT du mot transmis OPTIONS 2
'                           ---------------------------------------------------------------------------------------
'                           |  Bit 7 |  Bit 6 | Bit 5 | Bit 4 | Bit 3 | Bit 2 | Bit 1 | Bit 0 |
'                           ---------------------------------------------------------------------------------------
'                           |  128   |   64   |   32  |   16   |    8   |    4    |    2   |     1   |
'                           ---------------------------------------------------------------------------------------
'                                 |           |          |         |         |          |          |         |_____  gestion de l'électro-vanne du brillantage avec les gammes
'                                 |           |          |         |         |          |          |__________
'                                 |           |          |         |         |          |________________
'                                 |           |          |         |         |_____________________
'                                 |           |          |         |__________________________
'                                 |           |          |_______________________________
'                                 |           |_____________________________________
'                                 |___________________________________________
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function EnvoiOptionsPourUneCharge(ByVal NumCharge As Integer, _
                                                                          ByVal Options1 As Integer, _
                                                                          ByVal Options2 As Integer) As String

    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next
    
    '--- constantes privées ---
    Const NOM_GROUPE As String = "SUIVI_LIGNE"
    Const OPTIONS_GAMME_1_POSTE As String = "OptionsGamme1Poste"     'variable options de la gamme partie 1 pour un poste
    Const OPTIONS_GAMME_2_POSTE As String = "OptionsGamme2Poste"     'variable options de la gamme partie 1 pour un poste
    Const OPTIONS_GAMME_1_PONT As String = "OptionsGamme1P"              'variable options de la gamme partie 1 pour un pont
    Const OPTIONS_GAMME_2_PONT As String = "OptionsGamme2P"              'variable options de la gamme partie 1 pour un pont
    
    '--- déclaration ---
    Dim NumPoste As Integer, _
           NumPont As Integer
    Dim ValeurRetourneeAPI As Long                  'valeur retournée par une fonction concernant le dialogue avec l'automate
    Dim NomVariableOptions1 As String              'nom de la variable pour les options 1
    Dim NomVariableOptions2 As String              'nom de la variable pour les options 2
    
    '--- affectation ---
    EnvoiOptionsPourUneCharge = ""

    If NumCharge >= CHARGES.C_NUM_MINI And NumCharge <= CHARGES.C_NUM_MAXI Then
    
        '--- recherche du poste ou se trouve la charge ---
        NumPoste = RechercheNumPostePourUneCharge(NumCharge)
    
        If NumPoste <> 0 Then
        
            If NumPoste < 0 Then
    
                '--- le n° de poste est négatif alors la charge est sur un des ponts ---
                NumPont = Abs(NumPoste)
                
                '--- calcul de l'adresse des options pour UN PONT ---
                NomVariableOptions1 = OPTIONS_GAMME_1_PONT & NumPont
                NomVariableOptions2 = OPTIONS_GAMME_2_PONT & NumPont
                
            ElseIf NumPoste >= POSTES.P_CHGT_1 And NumPoste <= DERNIER_POSTE Then
                        
                '--- calcul de l'adresse des options pour UN POSTE ---
                NomVariableOptions1 = OPTIONS_GAMME_1_POSTE & Right("00" & NumPoste, 2)
                NomVariableOptions2 = OPTIONS_GAMME_2_POSTE & Right("00" & NumPoste, 2)
            
            End If
    
            '--- écriture des options ---
            If PROGRAMME_AVEC_AUTOMATE = True Then
                
                '--- transfert de l'option 1 ---
                ValeurRetourneeAPI = APIEcritureVariableNommee(NOM_GROUPE, NomVariableOptions1, Options1)
                If ValeurRetourneeAPI = 0 Then
                
                    '--- transfert de l'option 2 ---
                    ValeurRetourneeAPI = APIEcritureVariableNommee(NOM_GROUPE, NomVariableOptions2, Options2)
                    
                    If ValeurRetourneeAPI = 0 Then
                        EnvoiOptionsPourUneCharge = OK
                    End If
            
                End If
            Else
                EnvoiOptionsPourUneCharge = OK
            End If
        
        End If

    End If
    
End Function

'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Rôle      : Recherche le numéro de poste ou se trouve une charge
' Entrées :                                          NumCharge -> N° de la charge ou l'on recherche le poste actuel
' Retours : RechercheNumPostePourUneCharge -> 0 si pas de charge dans la ligne
'                                                                                 moins x, une valeur négative représente le numéro du pont si
'                                                                                 la charge se trouve sur un des ponts
'                                                                                 plus x,  une valeur positive représente le numéro du poste ou
'                                                                                 se trouve la charge
' Détails  :
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function RechercheNumPostePourUneCharge(ByVal NumCharge As Integer) As Integer

    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next

    '--- déclaration ---
    Dim a As Integer

    '--- affectation ---
    RechercheNumPostePourUneCharge = 0
                                                                      
    If NumCharge >= CHARGES.C_NUM_MINI And NumCharge <= CHARGES.C_NUM_MAXI Then
                                                                      
        '--- recherche dans les postes ---
        For a = LBound(TEtatsPostes()) To UBound(TEtatsPostes())
            With TEtatsPostes(a)
                If NumCharge = .NumCharge Then
                    RechercheNumPostePourUneCharge = a
                End If
            End With
        Next a
                                                                      
        '--- recherche sur les ponts ---
        If RechercheNumPostePourUneCharge = 0 Then
            For a = LBound(TEtatsPonts()) To UBound(TEtatsPonts())
                With TEtatsPonts(a)
                    If NumCharge = .NumCharge Then
                        RechercheNumPostePourUneCharge = -a
                    End If
                End With
            Next a
        End If

    End If

End Function

'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Rôle      : Recherche le temps passé au poste pour une charge
' Entrées :                                                 NumCharge -> N° de la charge
'                                                                   NumPoste -> N° du poste ou l'on recherche le temps passé
' Retours : RechercheTempsAuPostePourUneCharge -> 0 si pas de temps au poste ou pas de passage dans ce
'                                                                                         poste
'                                                                                         Sinon le temps passé au poste en secondes
'                                            DateEntreeDansLePoste -> Date complète d'entrée dans le poste
' Détails  :
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function RechercheTempsAuPostePourUneCharge(ByVal NumCharge As Integer, _
                                                                                              ByVal NumPoste As Integer, _
                                                                                              ByRef DateEntreeDansLePoste As Date) As Long

    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next

    '--- déclaration ---
    Dim a As Integer
    
    '--- affectation ---
    RechercheTempsAuPostePourUneCharge = 0
                                                                      
    If NumCharge >= CHARGES.C_NUM_MINI And NumCharge <= CHARGES.C_NUM_MAXI Then
                                                                      
        If NumPoste >= POSTES.P_CHGT_1 And NumPoste <= DERNIER_POSTE Then
        
            With TEtatsCharges(NumCharge)
        
                For a = LBound(.TDetailsFichesProduction()) To UBound(.TDetailsFichesProduction())
            
                    With .TDetailsFichesProduction(a)
        
                        '--- recherche du temps si le poste a été trouvé ---
                        If .NumPoste = NumPoste Then
            
                            '--- temps réel au poste ---
                            If .DateEntreePoste <> Empty And .DateSortiePoste <> Empty Then
                                DateEntreeDansLePoste = .DateEntreePoste
                                RechercheTempsAuPostePourUneCharge = DateDiff("s", .DateEntreePoste, .DateSortiePoste)
                                Exit For
                            End If
                    
                        End If
                    
                    End With
            
                Next a
            
            End With
        
        End If
        
    End If
        
End Function

'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Rôle      : Recherche le décompte du temps passé pour une charge dans un poste
' Entrées :                                          NumPoste -> N° du poste ou l'on recherche le décompte du temps passé
' Retours : RechercheDecompteTempsAuPoste -> décompte du temps passé au poste en secondes
'                                 ChargePresenteAuPostee -> TRUE = il y a une charge dans ce poste
'                                                                               FALSE = pas de charge présente dans ce poste
' Détails  :
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function RechercheDecompteTempsAuPoste(ByVal NumPoste As Integer, _
                                                                                     ByRef ChargePresenteAuPoste As Boolean) As Long

    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next

    '--- déclaration ---
    Dim NumCharge As Integer                                                       'numéro de charge

    '--- affectation par défaut ---
    RechercheDecompteTempsAuPoste = 0
    ChargePresenteAuPoste = False
    
    If NumPoste >= POSTES.P_CHGT_1 And NumPoste <= DERNIER_POSTE Then
    
        If TEtatsPostes(NumPoste).DefinitionPoste.AvecTemps = True Then
        
            '--- affectation du numéro de charge ---
            NumCharge = TEtatsPostes(NumPoste).NumCharge
    
            If NumCharge >= CHARGES.C_NUM_MINI And NumCharge <= CHARGES.C_NUM_MAXI Then
                                                                      
                '--- affectation du décompte du temps ---
                With TEtatsCharges(NumCharge)
                    
                    If .PtrZoneGammeAnodisation > 0 And .NbrPostesTraites > 0 Then
                        
                        If .TGammesAnodisation.TDetailsGammesAnodisation(.PtrZoneGammeAnodisation).NumPosteReel = NumPoste Then
                                        
                            '--- affectation du décompte du temps au poste ---
                            RechercheDecompteTempsAuPoste = .TGammesAnodisation.TDetailsGammesAnodisation(.PtrZoneGammeAnodisation).DecompteDuTempsAuPosteReelSecondes
                                
                            '--- affectation indiquant qu'une charge est bien dans le poste ---
                            ChargePresenteAuPoste = True
                        
                        End If
                            
                    End If
                        
                End With
            
            End If
                    
        End If
                
    End If

End Function

'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Rôle      : Recherche le temps passé au poste d'anodisation
' Entrées :                                      NumCharge -> N° de la charge
'                                              NumPosteAnodisation -> N° du poste d'anodisation
' Retours : RechercheTempsAuPosteDeAnodisation -> 0 si pas de temps au poste ou pas de passage dans au poste
'                                                                             d'anodisation
'                                                                             Sinon le temps passé au poste en secondes
'                               DateEntreeAuPosteAnodisation -> Date complète d'entrée dans le poste d'anodisation
' Détails  :
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function RechercheTempsAuPosteAnodisation(ByVal NumCharge As Integer, _
                                                                                       ByRef NumPosteAnodisation As Integer, _
                                                                                       ByRef DateEntreeAuPosteAnodisation As Date) As Long

    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next

    '--- déclaration ---
    Dim a As Integer
    
    '--- affectation ---
    RechercheTempsAuPosteAnodisation = 0
                                                                      
    If NumCharge >= CHARGES.C_NUM_MINI And NumCharge <= CHARGES.C_NUM_MAXI Then
        
        With TEtatsCharges(NumCharge)
    
            For a = LBound(.TDetailsFichesProduction()) To UBound(.TDetailsFichesProduction())
        
                With .TDetailsFichesProduction(a)
    
                    '--- recherche du temps si le poste a été trouvé ---
                    If .NumPoste = POSTES.P_C13 Or .NumPoste = POSTES.P_C14 Or .NumPoste = POSTES.P_C15 Or .NumPoste = POSTES.P_C16 Then
        
                        '--- temps réel au poste ---
                        If .DateEntreePoste <> Empty And .DateSortiePoste <> Empty Then
                            NumPosteAnodisation = .NumPoste
                            DateEntreeAuPosteAnodisation = .DateEntreePoste
                            RechercheTempsAuPosteAnodisation = DateDiff("s", .DateEntreePoste, .DateSortiePoste)
                            Exit For
                        End If
                
                    End If
                
                End With
        
            Next a
        
        End With
        
    End If
        
End Function

'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Rôle      : Recherche le numéro de charge le PLUS PETIT dans la ligne
' Entrées :
' Retours : RechercheNumeroChargeLePlusPetit -> 0 si pas de charge dans la ligne sinon le numéro le plus petit
' Détails  :
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function RechercheNumeroChargeLePlusPetit() As Integer
    
    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next
    
    '--- déclaration ---
    Dim a As Integer, _
           NumCharge As Integer
                                                                      
    '--- affectation ---
    RechercheNumeroChargeLePlusPetit = 0
    NumCharge = CHARGES.C_NUM_MAXI + 1 'forcer à la valeur la plus élevée par rapport au n° de charge maxi
    
    '--- recherche sur les ponts ---
    For a = LBound(TEtatsPonts()) To UBound(TEtatsPonts())
        With TEtatsPonts(a)
            If .NumCharge > 0 And .NumCharge < NumCharge Then
                NumCharge = .NumCharge
            End If
        End With
    Next a
                                                                      
    '--- recherche dans les postes ---
    For a = LBound(TEtatsPostes()) To UBound(TEtatsPostes())
        With TEtatsPostes(a)
            If .NumCharge > 0 And .NumCharge < NumCharge Then
                NumCharge = .NumCharge
            End If
        End With
    Next a
                                                                      
    '--- valeur de retour ---
    If NumCharge <> Succ(CHARGES.C_NUM_MAXI) Then
        RechercheNumeroChargeLePlusPetit = NumCharge
    End If
                                                                      
End Function

'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Rôle      : Vérifie l'existence d'un numéro de charge dans la ligne
' Entrées :                      NumCharge -> Numéro de charge faisant l'objet du contrôle
' Retours : ExistenceNumeroCharge -> FALSE = La charge n'existe pas dans la ligne
'                                                               TRUE = La charge existe déjà dans la ligne
' Détails  :
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function ExistenceNumeroCharge(ByVal NumCharge As Integer) As Boolean

    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next
    
    '--- déclaration ---
    Dim a As Integer

    '--- affectation ---
    ExistenceNumeroCharge = False
    
    '--- analyse pour les postes ---
    For a = POSTES.P_CHGT_1 To DERNIER_POSTE
    
        If (a <> POSTES.P_D1 And a <> POSTES.P_D2) Then
            If TEtatsPostes(a).NumCharge = NumCharge Then
                ExistenceNumeroCharge = True
                Exit Function
            End If
        End If
        
        
    Next a

    '--- analyse pour les ponts ---
    For a = LBound(TEtatsPonts()) To UBound(TEtatsPonts())
        If TEtatsPonts(a).NumCharge = NumCharge Then
            ExistenceNumeroCharge = True
            Exit Function
        End If
    Next a
    
End Function

'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Rôle      : Vérifie l'existence d'au moins une charge dans la ligne sans compter le déchargement
' Entrées :
' Retours : ExistenceNumeroChargeHorsDechargement -> FALSE = Aucune charge dans la ligne
'                                                                                               TRUE = Une charge au moins existe dans la ligne
' Détails  :
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function ExistenceChargeEnLigneHorsDechargement() As Boolean

    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next
    
    '--- déclaration ---
    Dim a As Integer

    '--- affectation ---
    ExistenceChargeEnLigneHorsDechargement = False
    
    '--- analyse pour les postes ---
    For a = POSTES.P_CHGT_1 To DERNIER_POSTE
    
        If (a <> POSTES.P_D1 And a <> POSTES.P_D2) Then
            If TEtatsPostes(a).NumCharge >= CHARGES.C_NUM_MINI And a <> POSTES.P_D1 And a <> POSTES.P_D2 Then
                ExistenceChargeEnLigneHorsDechargement = True
                Exit Function
            End If
        End If
        
        
    Next a

    '--- analyse pour les ponts ---
    For a = LBound(TEtatsPonts()) To UBound(TEtatsPonts())
        If TEtatsPonts(a).NumCharge >= CHARGES.C_NUM_MINI Then
            ExistenceChargeEnLigneHorsDechargement = True
            Exit Function
        End If
    Next a
    
End Function

'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Rôle      : Vérifie l'existence d'au moins une charge dans la ligne sans compter le chargement et le
'                 déchargement
' Entrées :
' Retours : ExistenceNumeroChargeHorsChargementDechargement -> FALSE = Aucune charge dans la ligne
'                                                                                                                 TRUE = Une charge au moins existe dans
'                                                                                                                               la ligne
' Détails  :
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function ExistenceChargeEnLigneHorsChargementDechargement() As Boolean

    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next
    
    '--- déclaration ---
    Dim a As Integer

    '--- affectation ---
    ExistenceChargeEnLigneHorsChargementDechargement = False
    
    '--- analyse pour les postes ---
    For a = PREMIER_BAIN To DERNIER_POSTE
    

        If (a <> POSTES.P_D1 And a <> POSTES.P_D2) Then
            If TEtatsPostes(a).NumCharge >= CHARGES.C_NUM_MINI Then
                ExistenceChargeEnLigneHorsChargementDechargement = True
                Exit Function
            End If
        End If
        
    Next a

    '--- analyse pour les ponts ---
    For a = LBound(TEtatsPonts()) To UBound(TEtatsPonts())
        If TEtatsPonts(a).NumCharge >= CHARGES.C_NUM_MINI Then
            ExistenceChargeEnLigneHorsChargementDechargement = True
            Exit Function
        End If
    Next a
    
End Function

'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Rôle      : Enregistre le numéro du poste réel (cas des postes multiples notamment) dans la gamme d'anodisation
'                 d'une charge se trouvant à un poste
' Entrées : NumPoste -> Numéro du poste de contrôle
' Retours :
' Détails  :
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Sub EnregistreNumPosteReelGamme(ByVal NumPoste As Integer)
    
    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next
    
    '--- déclaration ---
    Dim NumCharge As Integer, _
           NumZone As Integer, _
           NumPremierPosteZone As Integer, _
           NumDernierPosteZone As Integer

    '--- enregistrement ---
    If NumPoste >= POSTES.P_CHGT_1 And NumPoste <= DERNIER_POSTE Then
    
        '--- affectation du numéro de charge ---
        NumCharge = TEtatsPostes(NumPoste).NumCharge

        If NumCharge >= CHARGES.C_NUM_MINI And NumCharge <= CHARGES.C_NUM_MAXI Then
        
            With TEtatsCharges(NumCharge)
                        
                If .PtrZoneGammeAnodisation > 0 Then
                        
                    '--- recherche de la zone dans la gamme ---
                    NumZone = .TGammesAnodisation.TDetailsGammesAnodisation(.PtrZoneGammeAnodisation).NumZone
                        
                    If NumZone >= LIMITE_BASSE_ZONES And NumZone <= LIMITE_HAUTE_ZONES Then
        
                        '--- affectation ---
                        NumPremierPosteZone = TZones(NumZone).NumPremierPoste
                        NumDernierPosteZone = TZones(NumZone).NumDernierPoste
    
                        '--- affectation du numéro ---
                        If NumPoste >= NumPremierPosteZone And NumPoste <= NumDernierPosteZone Then
                            If .TGammesAnodisation.TDetailsGammesAnodisation(.PtrZoneGammeAnodisation).NumPosteReel = 0 Then
                                .TGammesAnodisation.TDetailsGammesAnodisation(.PtrZoneGammeAnodisation).NumPosteReel = NumPoste
                            End If
                        End If
                        
                    End If

                End If
        
            End With
        
        End If

    End If

End Sub

'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Rôle      : Incrémente le pointeur de la zone d'anodisation une fois la charge arrivée dans le poste
' Entrées : NumPoste -> Numéro du poste de contrôle
' Retours :
' Détails  :
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Sub IncrementationPtrZoneGammeAnodisation(ByVal NumPoste As Integer)
    
    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next
    
    '--- déclaration ---
    Dim NumCharge As Integer, _
           NumProchaineZone As Integer, _
           NumPremierPosteProchaineZone As Integer, _
           NumDernierPosteProchaineZone As Integer

    '--- enregistrement ---
    If NumPoste >= POSTES.P_CHGT_1 And NumPoste <= DERNIER_POSTE Then
    
        '--- affectation du numéro de charge ---
        NumCharge = TEtatsPostes(NumPoste).NumCharge

        If NumCharge >= CHARGES.C_NUM_MINI And NumCharge <= CHARGES.C_NUM_MAXI Then
        
            With TEtatsCharges(NumCharge)
                        
                If .PtrZoneGammeAnodisation > 0 Then
                        
                    '--- recherche de la zone dans la gamme ---
                    NumProchaineZone = .TGammesAnodisation.TDetailsGammesAnodisation(Succ(.PtrZoneGammeAnodisation)).NumZone
                        
                    If NumProchaineZone >= LIMITE_BASSE_ZONES And NumProchaineZone <= LIMITE_HAUTE_ZONES Then
        
                        '--- affectation ---
                        NumPremierPosteProchaineZone = TZones(NumProchaineZone).NumPremierPoste
                        NumDernierPosteProchaineZone = TZones(NumProchaineZone).NumDernierPoste
    
                        '--- affectation du numéro ---
                        If NumPoste >= NumPremierPosteProchaineZone And NumPoste <= NumDernierPosteProchaineZone Then
                            .PtrZoneGammeAnodisation = .PtrZoneGammeAnodisation + 1
                        End If
                        
                    End If

                End If
        
            End With
        
        End If

    End If

End Sub

'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Rôle      : Recherche le numéro de charge le PLUS GRAND dans la ligne
' Entrées :
' Retours : RechercheNumeroChargeLePlusGrand -> 0 si pas de charge dans la ligne sinon le numéro le plus
'                 grand
' Détails  :
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function RechercheNumeroChargeLePlusGrand() As Integer
    
    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next
    
    '--- déclaration ---
    Dim a As Integer, _
           NumCharge As Integer
                                                                      
    '--- affectation ---
    RechercheNumeroChargeLePlusGrand = 0
    NumCharge = 0
    
    '--- recherche sur les ponts ---
    For a = LBound(TEtatsPonts()) To UBound(TEtatsPonts())
        With TEtatsPonts(a)
            If .NumCharge > NumCharge Then
                NumCharge = .NumCharge
            End If
        End With
    Next a
                                                                      
    '--- recherche dans les postes ---
    For a = LBound(TEtatsPostes()) To UBound(TEtatsPostes())
        With TEtatsPostes(a)
            If .NumCharge > NumCharge Then
                NumCharge = .NumCharge
            End If
        End With
    Next a
                                                                      
    '--- valeur de retour ---
    RechercheNumeroChargeLePlusGrand = NumCharge
                                                                      
End Function

'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Rôle      : Envoi dans l'automate d'un numéro de charge à un POSTE
' Entrées :    NumPoste -> Numéro du poste concerné
'                 NumCharge -> Numéro de charge
' Retours :
' Détails  :
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function EnvoiNumeroChargePoste(ByVal NumPoste As Integer, _
                                                                      ByVal NumCharge As Integer) As String
    
    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next
    
    '--- constantes privées ---
    Const NOM_GROUPE As String = "SUIVI_LIGNE"
    
    '--- déclaration ---
    Dim ValeurRetourneeAPI As Long                  'valeur retournée par une fonction concernant le dialogue avec l'automate
    Dim NomVariable As String                            'nom de la variable
    
    '--- affectation ---
    EnvoiNumeroChargePoste = ""

    If NumPoste >= POSTES.P_CHGT_1 And NumPoste <= DERNIER_POSTE Then
    
        '--- affectation du nom de la variable ---
        NomVariable = "NumChargePoste" & Right("00" & NumPoste, 2)
                
        '--- écriture ---
        If PROGRAMME_AVEC_AUTOMATE = True Then
            ValeurRetourneeAPI = APIEcritureVariableNommee(NOM_GROUPE, NomVariable, NumCharge)
            If ValeurRetourneeAPI = 0 Then
                EnvoiNumeroChargePoste = OK
            End If
        Else
            EnvoiNumeroChargePoste = OK
        End If

    End If
    
End Function

'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Rôle      : Envoi dans l'automate d'un numéro de charge à un POSTE avec les options
' Entrées :                                               NumPoste -> Numéro du poste concerné
'                                                            NumCharge -> Numéro de charge
'                                                                Options1 -> Valeur de toutes les options 1
'                                                                Options2 -> Valeur de toutes les options 2
' Retours : EnvoiNumeroChargePosteAvecOptions -> OK = Transmission correcte
'                                                                                      "" = Incident de transmission
' Détails  :
'                           Poids FORT du mot transmis
'                           ---------------------------------------------------------------------------------------
'                           |  Bit 7 |  Bit 6 | Bit 5 | Bit 4 | Bit 3 | Bit 2 | Bit 1 | Bit 0 |
'                           ---------------------------------------------------------------------------------------
'                           |  128   |   64   |   32  |   16   |    8   |    4    |    2   |     1   |
'                           ---------------------------------------------------------------------------------------
'                           |           |          |         |         |          |          |         |_____  forcer la montée en très petite vitesse
'                           |           |          |         |         |          |          |__________  forcer la montée en petite vitesse
'                           |           |          |         |         |          |________________ forcer la descente en très petite vitesse
'                           |           |          |         |         |_____________________  forcer la descente en petite vitesse
'                           |           |          |         |__________________________
'                           |           |          |_______________________________
'                           |           |_____________________________________
'                           |___________________________________________
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function EnvoiNumeroChargePosteAvecOptions(ByVal NumPoste As Integer, _
                                                                                          ByVal NumCharge As Integer, _
                                                                                          ByVal Options1 As Integer, _
                                                                                          ByVal Options2 As Integer) As String
    
    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next
    
    '--- constantes privées ---
    Const NOM_GROUPE As String = "SUIVI_LIGNE"
    Const OPTIONS_GAMME_1_POSTE As String = "OptionsGamme1Poste"     'variable options de la gamme partie 1 pour un poste
    Const OPTIONS_GAMME_2_POSTE As String = "OptionsGamme2Poste"     'variable options de la gamme partie 1 pour un poste
    
    '--- déclaration ---
    Dim ValeurRetourneeAPI As Long                                 'valeur retournée par une fonction concernant le dialogue avec l'automate
    Dim NomVariableNumChargePoste As String               'nom de la variable pour le numéro de charge au poste
    Dim NomVariableOption1 As String                               'nom de la variable pour l'option 1
    Dim NomVariableOption2 As String                               'nom de la variable pour l'option 2
    
    '--- affectation ---
    EnvoiNumeroChargePosteAvecOptions = ""
    
    If NumPoste >= POSTES.P_CHGT_1 And NumPoste <= DERNIER_POSTE Then
                    
        If NumCharge = 0 Or (NumCharge >= CHARGES.C_NUM_MINI And NumCharge <= CHARGES.C_NUM_MAXI) Then

            '--- affectation de toutes les variables ---
            NomVariableNumChargePoste = "NumChargePoste" & Right("00" & NumPoste, 2)
            NomVariableOption1 = OPTIONS_GAMME_1_POSTE & Right("00" & NumPoste, 2)
            NomVariableOption2 = OPTIONS_GAMME_2_POSTE & Right("00" & NumPoste, 2)
            
            If PROGRAMME_AVEC_AUTOMATE = True Then
            
                '--- écriture du numéro de charge ---
                ValeurRetourneeAPI = APIEcritureVariableNommee(NOM_GROUPE, NomVariableNumChargePoste, NumCharge)
                If ValeurRetourneeAPI = 0 Then
                
                    '--- écriture des options 1 ---
                    ValeurRetourneeAPI = APIEcritureVariableNommee(NOM_GROUPE, NomVariableOption1, Options1)
                    
                    If ValeurRetourneeAPI = 0 Then
                    
                        '--- écriture des options 2 ---
                        ValeurRetourneeAPI = APIEcritureVariableNommee(NOM_GROUPE, NomVariableOption2, Options2)

                        If ValeurRetourneeAPI = 0 Then
                            EnvoiNumeroChargePosteAvecOptions = OK
                        End If

                    End If
                
                End If
            Else
                EnvoiNumeroChargePosteAvecOptions = OK
            End If
        
        End If

    End If
    
End Function

'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Rôle      : Envoi dans l'automate d'un numéro de charge à un PONT
' Entrées :      NumPont -> Numéro du pont concerné
'                 NumCharge -> Numéro de charge
' Retours :
' Détails  :
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function EnvoiNumeroChargePont(ByVal NumPont As Integer, _
                                                                    ByVal NumCharge As Integer) As String
    
    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next
    
    '--- constantes privées ---
    Const NOM_GROUPE As String = "SUIVI_LIGNE"
    
    '--- déclaration ---
    Dim ValeurRetourneeAPI As Long                  'valeur retournée par une fonction concernant le dialogue avec l'automate
    Dim NomVariable As String                            'nom de la variable
    

    
    '--- affectation ---
    EnvoiNumeroChargePont = ""
    
    If NumPont = PONTS.P_1 Or NumPont = PONTS.P_2 Then
                    
        If NumCharge = 0 Or (NumCharge >= CHARGES.C_NUM_MINI And NumCharge <= CHARGES.C_NUM_MAXI) Then

            '--- affectation du nom de la variable ---
            NomVariable = "NumChargeP" & NumPont
                
            '--- écriture ---
            If PROGRAMME_AVEC_AUTOMATE = True Then
                ValeurRetourneeAPI = APIEcritureVariableNommee(NOM_GROUPE, NomVariable, NumCharge)
                If ValeurRetourneeAPI = 0 Then
                    EnvoiNumeroChargePont = OK
                End If
            Else
                EnvoiNumeroChargePont = OK
            End If

        End If
    
    End If
                    
End Function

'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Rôle      : Recherche le prochain numéro de charge valide pour une entrée dans la ligne
' Entrées :
' Retours :
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function ProchainNumeroCharge() As Integer

    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next

    '--- déclaration ---
    Dim TNumChargesUtilisees(CHARGES.C_NUM_MINI To CHARGES.C_NUM_MAXI) As Boolean
    Dim a As Integer, _
           LeProchainNumeroCharge As Integer

    '--- affectation ---
    LeProchainNumeroCharge = 0
    
    '--- recherche du prochain numéro de charge pour les ponts ---
    For a = LBound(TEtatsPonts) To UBound(TEtatsPonts())
        With TEtatsPonts(a)
            If .NumCharge >= CHARGES.C_NUM_MINI And .NumCharge <= CHARGES.C_NUM_MAXI Then
                
                '--- indique que le n° de charge est déjà utilisé ---
                TNumChargesUtilisees(.NumCharge) = True
                
                '--- prendre le n° de charge le plus élevé ---
                If .NumCharge > LeProchainNumeroCharge Then
                    LeProchainNumeroCharge = .NumCharge
                End If
            
            End If
        End With
    Next a
    
    '--- recherche du prochain numéro de charge pour les postes ---
    For a = LBound(TEtatsPostes()) To UBound(TEtatsPostes())
        With TEtatsPostes(a)
            If .NumCharge >= CHARGES.C_NUM_MINI And .NumCharge <= CHARGES.C_NUM_MAXI Then
                
                '--- indique que le n° de charge est déjà utilisé ---
                TNumChargesUtilisees(.NumCharge) = True
                
                '--- prendre le n° de charge le plus élevé ---
                If .NumCharge > LeProchainNumeroCharge Then
                    LeProchainNumeroCharge = .NumCharge
                End If
            
            End If
        End With
    Next a

    '--- incrémentation de la variable ---
    Inc LeProchainNumeroCharge

    '--- analyse des limites ---
    If LeProchainNumeroCharge > CHARGES.C_NUM_MAXI Then
        
        '--- rechercher le premier n° de charge non utilisé ---
        For a = LBound(TNumChargesUtilisees()) To UBound(TNumChargesUtilisees())
            If TNumChargesUtilisees(a) = False Then
                LeProchainNumeroCharge = a
                Exit For
            End If
        Next a
        If LeProchainNumeroCharge >= CHARGES.C_NUM_MAXI Then
            LeProchainNumeroCharge = CHARGES.C_NUM_MINI
        End If
    
    End If

    '--- valeur de retour ---
    ProchainNumeroCharge = LeProchainNumeroCharge

End Function

'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Rôle      : Recherche le prochain numéro de poste au chargement pour entrée la charge dans la ligne pour une
'                 gamme de production ou C13 est imposé
' Entrées :
' Retours :                                                                      ChargePrioritaire -> Indique que la charge est prioritaire
'                 ProchainNumeroPosteChargementSiAnodisationC13Impose ->            0 = Pas de charge en entrée de ligne
'                                                                                                                       C1 à C6 = Numéro de poste ou se trouve la
'                                                                                                                                        charge à rentrer
' Détails  :
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function ProchainNumeroPosteChargementSiAnodisationC13Impose(ByRef ChargePrioritaire As Boolean) As Integer

    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next

    '--- déclaration ---
    Dim a As Integer
    Dim DateEntreeEnLigne As Date, _
            MemDateEntreeEnLigne As Date
    
    '--- affectation ---
    ProchainNumeroPosteChargementSiAnodisationC13Impose = 0
    ChargePrioritaire = False
    
    '--- analyse du chargement ---
    For a = POSTES.P_CHGT_1 To POSTES.P_CHGT_2
        
        With TEtatsPostes(a)
            
            If .Condamnation = False Then
                
                If .EtatsChariots = E_PRESENT_VERROUILLE Then
                    
                    If .NumCharge >= CHARGES.C_NUM_MINI And .NumCharge <= CHARGES.C_NUM_MAXI Then
                        
                        If TEtatsCharges(.NumCharge).TGammesAnodisation.ChoixPosteAnodisation = C_C13_IMPOSE Then
                            
                            If TEtatsPostes(POSTES.P_C13).Condamnation = False Then     'une charge ne peut entrer si le
                                                                                                                                   'poste est condamné
                                If TEtatsCharges(.NumCharge).ChargePrioritaire = True Then
                                    
                                    '--- la charge est prioritaire ---
                                    ChargePrioritaire = True
                                    ProchainNumeroPosteChargementSiAnodisationC13Impose = a
                                    Exit For
                                
                                Else
                                    
                                    '--- la charge n'est pas prioritaire donc contrôler la date d'entrée ---
                                    DateEntreeEnLigne = TEtatsCharges(.NumCharge).DateEntreeEnLigne
                                    If DateEntreeEnLigne <> Empty Then
                                        If MemDateEntreeEnLigne = Empty Then
                                            MemDateEntreeEnLigne = DateEntreeEnLigne
                                            ProchainNumeroPosteChargementSiAnodisationC13Impose = a
                                        Else
                                            If DateEntreeEnLigne < MemDateEntreeEnLigne Then
                                                MemDateEntreeEnLigne = DateEntreeEnLigne
                                                ProchainNumeroPosteChargementSiAnodisationC13Impose = a
                                            End If
                                        End If
                                    End If
                                
                                End If
                        
                            End If
                        
                        End If
                    
                    End If
                
                End If
            
            End If
        
        End With
    
    Next a
    
End Function

'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Rôle      : Recherche le prochain numéro de poste au chargement pour entrée la charge dans la ligne pour une
'                 gamme d'anodisation ou C15 est imposé
' Entrées :
' Retours :                                                             ChargePrioritaire -> Indique que la charge est prioritaire
'                 ProchainNumeroPosteChargementSiAnodisationC15Impose ->            0 = Pas de charge en entrée de ligne
'                                                                                                             C1 à C6 = Numéro de poste ou se trouve la
'                                                                                                                              charge à rentrer
' Détails  :
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function ProchainNumeroPosteChargementSiAnodisationC16Impose(ByRef ChargePrioritaire As Boolean) As Integer

    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next

    '--- déclaration ---
    Dim a As Integer
    Dim DateEntreeEnLigne As Date, _
            MemDateEntreeEnLigne As Date
    
    '--- affectation ---
    ProchainNumeroPosteChargementSiAnodisationC16Impose = 0
    ChargePrioritaire = False
    
    '--- analyse du chargement ---
    For a = POSTES.P_CHGT_1 To POSTES.P_CHGT_2
        
        With TEtatsPostes(a)
            
            If .Condamnation = False Then
                
                If .EtatsChariots = E_PRESENT_VERROUILLE Then
                    
                    If .NumCharge >= CHARGES.C_NUM_MINI And .NumCharge <= CHARGES.C_NUM_MAXI Then
                        
                        If TEtatsCharges(.NumCharge).TGammesAnodisation.ChoixPosteAnodisation = C_C16_IMPOSE Then
                            
                            If TEtatsPostes(POSTES.P_C16).Condamnation = False Then     'une charge ne peut entrer si le
                                                                                                                                   'poste est condamné
                            
                                If TEtatsCharges(.NumCharge).ChargePrioritaire = True Then
                                    
                                    '--- la charge est prioritaire ---
                                    ChargePrioritaire = True
                                    ProchainNumeroPosteChargementSiAnodisationC16Impose = a
                                    Exit For
                                
                                Else
                                    
                                    '--- la charge n'est pas prioritaire donc contrôler la date d'entrée ---
                                    DateEntreeEnLigne = TEtatsCharges(.NumCharge).DateEntreeEnLigne
                                    If DateEntreeEnLigne <> Empty Then
                                        If MemDateEntreeEnLigne = Empty Then
                                            MemDateEntreeEnLigne = DateEntreeEnLigne
                                            ProchainNumeroPosteChargementSiAnodisationC16Impose = a
                                        Else
                                            If DateEntreeEnLigne < MemDateEntreeEnLigne Then
                                                MemDateEntreeEnLigne = DateEntreeEnLigne
                                                ProchainNumeroPosteChargementSiAnodisationC16Impose = a
                                            End If
                                        End If
                                    End If
                                
                                End If
                        
                            End If
                        
                        End If
                    
                    End If
                
                End If
            
            End If
        
        End With
    
    Next a

End Function

'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Rôle      : Recherche le prochain numéro de poste au chargement pour entrée la charge dans la ligne pour une
'                 gamme de production ou C14 est imposé
' Entrées :
' Retours :                                                                      ChargePrioritaire -> Indique que la charge est prioritaire
'                 ProchainNumeroPosteChargementSiAnodisationC14Impose ->            0 = Pas de charge en entrée de ligne
'                                                                                                                       C1 à C6 = Numéro de poste ou se trouve la
'                                                                                                                                        charge à rentrer
' Détails  :
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function ProchainNumeroPosteChargementSiAnodisationC14Impose(ByRef ChargePrioritaire As Boolean) As Integer

    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next

    '--- déclaration ---
    Dim a As Integer
    Dim DateEntreeEnLigne As Date, _
            MemDateEntreeEnLigne As Date
    
    '--- affectation ---
    ProchainNumeroPosteChargementSiAnodisationC14Impose = 0
    ChargePrioritaire = False
    
    '--- analyse du chargement ---
    For a = POSTES.P_CHGT_1 To POSTES.P_CHGT_2
        
        With TEtatsPostes(a)
            
            If .Condamnation = False Then
                
                If .EtatsChariots = E_PRESENT_VERROUILLE Then
                    
                    If .NumCharge >= CHARGES.C_NUM_MINI And .NumCharge <= CHARGES.C_NUM_MAXI Then
                        
                        If TEtatsCharges(.NumCharge).TGammesAnodisation.ChoixPosteAnodisation = C_C14_IMPOSE Then
                            
                            If TEtatsPostes(POSTES.P_C14).Condamnation = False Then     'une charge ne peut entrer si le
                                                                                                                                   'poste est condamné
                            
                                If TEtatsCharges(.NumCharge).ChargePrioritaire = True Then
                                    
                                    '--- la charge est prioritaire ---
                                    ChargePrioritaire = True
                                    ProchainNumeroPosteChargementSiAnodisationC14Impose = a
                                    Exit For
                                
                                Else
                                    
                                    '--- la charge n'est pas prioritaire donc contrôler la date d'entrée ---
                                    DateEntreeEnLigne = TEtatsCharges(.NumCharge).DateEntreeEnLigne
                                    If DateEntreeEnLigne <> Empty Then
                                        If MemDateEntreeEnLigne = Empty Then
                                            MemDateEntreeEnLigne = DateEntreeEnLigne
                                            ProchainNumeroPosteChargementSiAnodisationC14Impose = a
                                        Else
                                            If DateEntreeEnLigne < MemDateEntreeEnLigne Then
                                                MemDateEntreeEnLigne = DateEntreeEnLigne
                                                ProchainNumeroPosteChargementSiAnodisationC14Impose = a
                                            End If
                                        End If
                                    End If
                                
                                End If
                        
                            End If
                        
                        End If
                    
                    End If
                
                End If
            
            End If
        
        End With
    
    Next a
    
End Function

'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Rôle      : Recherche le prochain numéro de poste au chargement pour entrée la charge dans la ligne pour une
'                 gamme de production ou C15 est imposé
' Entrées :
' Retours :                                                                      ChargePrioritaire -> Indique que la charge est prioritaire
'                 ProchainNumeroPosteChargementSiAnodisationC15Impose ->            0 = Pas de charge en entrée de ligne
'                                                                                                                       C1 à C6 = Numéro de poste ou se trouve la
'                                                                                                                                        charge à rentrer
' Détails  :
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function ProchainNumeroPosteChargementSiAnodisationC15Impose(ByRef ChargePrioritaire As Boolean) As Integer

    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next

    '--- déclaration ---
    Dim a As Integer
    Dim DateEntreeEnLigne As Date, _
            MemDateEntreeEnLigne As Date
    
    '--- affectation ---
    ProchainNumeroPosteChargementSiAnodisationC15Impose = 0
    ChargePrioritaire = False
    
    '--- analyse du chargement ---
    For a = POSTES.P_CHGT_1 To POSTES.P_CHGT_2
        
        With TEtatsPostes(a)
            
            If .Condamnation = False Then
                
                If .EtatsChariots = E_PRESENT_VERROUILLE Then
                    
                    If .NumCharge >= CHARGES.C_NUM_MINI And .NumCharge <= CHARGES.C_NUM_MAXI Then
                        
                        If TEtatsCharges(.NumCharge).TGammesAnodisation.ChoixPosteAnodisation = C_C15_IMPOSE Then
                            
                            If TEtatsPostes(POSTES.P_C15).Condamnation = False Then     'une charge ne peut entrer si le
                                                                                                                                   'poste est condamné
                            
                                If TEtatsCharges(.NumCharge).ChargePrioritaire = True Then
                                    
                                    '--- la charge est prioritaire ---
                                    ChargePrioritaire = True
                                    ProchainNumeroPosteChargementSiAnodisationC15Impose = a
                                    Exit For
                                
                                Else
                                    
                                    '--- la charge n'est pas prioritaire donc contrôler la date d'entrée ---
                                    DateEntreeEnLigne = TEtatsCharges(.NumCharge).DateEntreeEnLigne
                                    If DateEntreeEnLigne <> Empty Then
                                        If MemDateEntreeEnLigne = Empty Then
                                            MemDateEntreeEnLigne = DateEntreeEnLigne
                                            ProchainNumeroPosteChargementSiAnodisationC15Impose = a
                                        Else
                                            If DateEntreeEnLigne < MemDateEntreeEnLigne Then
                                                MemDateEntreeEnLigne = DateEntreeEnLigne
                                                ProchainNumeroPosteChargementSiAnodisationC15Impose = a
                                            End If
                                        End If
                                    End If
                                
                                End If
                        
                            End If
                        
                        End If
                    
                    End If
                
                End If
            
            End If
        
        End With
    
    Next a
    
End Function

'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Rôle      : Recherche le prochain numéro de poste au chargement pour entrée la charge dans la ligne pour une
'                 gamme d'anodisation ou le chois du poste est automatique
' Entrées :
' Retours :                                                               ChargePrioritaire -> Indique que la charge est prioritaire
'                 ProchainNumeroPosteChargementSiAnodisationAutomatique ->           0 = Pas de charge en entrée de ligne
'                                                                                                               C1 à C6 = Numéro de poste ou se trouve la
'                                                                                                                                charge à rentrer
' Détails  :
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function ProchainNumeroPosteChargementSiAnodisationAutomatique(ByRef ChargePrioritaire As Boolean) As Integer

    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next

    '--- déclaration ---
    Dim a As Integer
    Dim DateEntreeEnLigne As Date, _
            MemDateEntreeEnLigne As Date
    
    '--- affectation ---
    ProchainNumeroPosteChargementSiAnodisationAutomatique = 0
    ChargePrioritaire = False
    
    '--- analyse du chargement ---
    For a = POSTES.P_CHGT_1 To POSTES.P_CHGT_2
        With TEtatsPostes(a)
            
            If .Condamnation = False Then
                
                If .EtatsChariots = E_PRESENT_VERROUILLE Then
                    
                    If .NumCharge >= CHARGES.C_NUM_MINI And .NumCharge <= CHARGES.C_NUM_MAXI Then
                        
                        If TEtatsCharges(.NumCharge).TGammesAnodisation.ChoixPosteAnodisation = C_AUTOMATIQUE Then
                            
                            If TEtatsPostes(POSTES.P_C13).Condamnation = False Or _
                               TEtatsPostes(POSTES.P_C14).Condamnation = False Or _
                               TEtatsPostes(POSTES.P_C15).Condamnation = False Or _
                               TEtatsPostes(POSTES.P_C16).Condamnation = False Then  'une charge ne peut entrer si les 4
                                                                                                                                'postes sont condamnés
                                If TEtatsCharges(.NumCharge).ChargePrioritaire = True Then
                                    
                                    '--- la charge est prioritaire ---
                                    ChargePrioritaire = True
                                    ProchainNumeroPosteChargementSiAnodisationAutomatique = a
                                    Exit For
                                
                                Else
                                    
                                    '--- la charge n'est pas prioritaire donc contrôler la date d'entrée ---
                                    DateEntreeEnLigne = TEtatsCharges(.NumCharge).DateEntreeEnLigne
                                    If DateEntreeEnLigne <> Empty Then
                                        If MemDateEntreeEnLigne = Empty Then
                                            MemDateEntreeEnLigne = DateEntreeEnLigne
                                            ProchainNumeroPosteChargementSiAnodisationAutomatique = a
                                        Else
                                            If DateEntreeEnLigne < MemDateEntreeEnLigne Then
                                                MemDateEntreeEnLigne = DateEntreeEnLigne
                                                ProchainNumeroPosteChargementSiAnodisationAutomatique = a
                                            End If
                                        End If
                                    End If
                                
                                End If
                        
                            End If
                        
                        End If
                    
                    End If
                
                End If
            
            End If
        
        End With
    Next a
    
End Function

'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Rôle      : Recherche le prochain numéro de poste au déchargement pour déposer une charge
' Entrées :
' Retours : ProchainNumeroPosteDechargement ->           0 = Pas de chariot vide pour déposer la charge
'                                                                                 D1 à D2 = Numéro de poste ou la dépose doit s'effectué
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function ProchainNumeroPosteDechargement() As Integer

    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next

    '--- déclaration ---
    Dim a As Integer
    
    '--- affectation ---
    ProchainNumeroPosteDechargement = 0
    
    '--- analyse du chargement ---
    For a = POSTES.P_D2 To POSTES.P_D1 Step -1
        With TEtatsPostes(a)
            If .Condamnation = False Then
                If .EtatsChariots = E_PRESENT_VERROUILLE And .NumCharge = 0 Then
                    ProchainNumeroPosteDechargement = a
                    Exit For
                End If
            End If
        End With
    Next a

End Function

'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Rôle      : Recherche le prochain numéro de poste au chargement pour déposer une charge
' Entrées :
' Retours : ProchainNumeroPosteChargement ->                          0 = Pas de chariot vide pour déposer la charge
'                                                                              CHGT1 à CHGT4 = Numéro de poste ou la dépose doit s'effectué
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function ProchainNumeroPosteChargement() As Integer

    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next

    '--- déclaration ---
    Dim a As Integer
    
    '--- affectation ---
    ProchainNumeroPosteChargement = 0
    
    '--- analyse du chargement ---
    For a = POSTES.P_CHGT_1 To POSTES.P_CHGT_2
        With TEtatsPostes(a)
            If .Condamnation = False Then
                If .EtatsChariots = E_PRESENT_VERROUILLE And .NumCharge = 0 Then
                    ProchainNumeroPosteChargement = a
                    Exit For
                End If
            End If
        End With
    Next a

End Function

'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Rôle      : Recherche le prochain numéro de poste d'anodisation disponible
' Entrées :                                 NumCharge -> numéro de la charge faisant l'objet de la demande
'                                               TypeDeZone -> FALSE = La zone est une ZONE de DEPART
'                                                                         TRUE = La zone est une ZONE d'ARRIVEE
' Retours : ProchainNumeroPosteAnodisation ->                              0 = Pas de poste d'anodisation disponible
'                                                                             C13 ou C14 ou C15 ou C16 = Numéro de poste d'anodisation libre
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function ProchainNumeroPosteAnodisation(ByVal NumCharge As Integer, _
                                                                                  ByVal TypeDeZone As Boolean) As Integer

    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next
    
    '--- constantes privées ---
    Const ZONE_DEPART As Boolean = False
    Const ZONE_ARRIVEE As Boolean = True

    '--- déclaration ---
    Dim a As Integer
    
    '--- affectation ---
    ProchainNumeroPosteAnodisation = 0
    
    '--- analyse sur les postes d'anodisation ---
    For a = POSTES.P_C13 To POSTES.P_C16
            
        '--- uniquement les postes d'anodisation ---
        If TEtatsPostes(a).Condamnation = False Then
        
            If TypeDeZone = ZONE_DEPART Then
                    
                '--- la zone d'anodisation est une zone de départ ---
                If NumCharge >= CHARGES.C_NUM_MINI And NumCharge <= CHARGES.C_NUM_MAXI Then
                    
                    '--- vérifier que le temps au poste est terminé ---
                    With TEtatsCharges(NumCharge)
                        If a = .TGammesAnodisation.TDetailsGammesAnodisation(.PtrZoneGammeAnodisation).NumPosteReel Then
                            If .TGammesAnodisation.TDetailsGammesAnodisation(.PtrZoneGammeAnodisation).FinDuTempsPosteReel = True Then
                                ProchainNumeroPosteAnodisation = a
                            End If
                        End If
                    End With
                
                End If
                
            Else
    
                '--- la zone d'anodisation est une zone d'arrivée ---
                If NumCharge = 0 Or (NumCharge >= CHARGES.C_NUM_MINI And NumCharge <= CHARGES.C_NUM_MAXI) Then
                
                    With TEtatsCharges(NumCharge)
                            
                        '--- recherche en fonction du choix du poste d'anodisation ---
                        Select Case .TGammesAnodisation.ChoixPosteAnodisation
                    
                            Case CHOIX_POSTE_ANODISATION.C_AUTOMATIQUE
                               '--- prendre le premier poste d'arrivée vide ---
                                If TEtatsPostes(a).NumCharge = 0 Then
                                    ProchainNumeroPosteAnodisation = a
                                End If
                            
                            Case CHOIX_POSTE_ANODISATION.C_C13_IMPOSE
                                '--- C13 est imposé ---
                                If a = POSTES.P_C13 Then
                                    If TEtatsPostes(a).NumCharge = 0 Then
                                        ProchainNumeroPosteAnodisation = a
                                    End If
                                End If
                            
                            Case CHOIX_POSTE_ANODISATION.C_C14_IMPOSE
                                '--- C14 est imposé ---
                                If a = POSTES.P_C14 Then
                                    If TEtatsPostes(a).NumCharge = 0 Then
                                        ProchainNumeroPosteAnodisation = a
                                    End If
                                End If
                            
                            Case CHOIX_POSTE_ANODISATION.C_C15_IMPOSE
                                '--- C15 est imposé ---
                                If a = POSTES.P_C15 Then
                                    If TEtatsPostes(a).NumCharge = 0 Then
                                        ProchainNumeroPosteAnodisation = a
                                    End If
                                End If
                            
                            Case CHOIX_POSTE_ANODISATION.C_C16_IMPOSE
                                '--- C16 est imposé ---
                                If a = POSTES.P_C16 Then
                                    If TEtatsPostes(a).NumCharge = 0 Then
                                        ProchainNumeroPosteAnodisation = a
                                    End If
                                End If
                
                            Case Else
                    
                        End Select
                    
                    End With
        
                End If
            
            End If
    
        End If

    Next a

End Function

'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Rôle      : Recherche le prochain numéro de poste de colmatage chaud
' Entrées :                                             NumCharge -> numéro de la charge faisant l'objet de la demande
'                                                           TypeDeZone -> FALSE = La zone est une ZONE de DEPART
'                                                                                     TRUE = La zone est une ZONE d'ARRIVEE
' Retours : ProchainNumeroPosteColmatageChaud ->                0 = Pas de poste de colmatage disponible
'                                                                                    C32 ou C33 = Numéro de poste de colmatage à chaud
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function ProchainNumeroPosteColmatageChaud(ByVal NumCharge As Integer, _
                                                                                          ByVal TypeDeZone As Boolean) As Integer

    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next
    
    '--- constantes privées ---
    Const ZONE_DEPART As Boolean = False
    Const ZONE_ARRIVEE As Boolean = True

    '--- déclaration ---
    Dim a As Integer
    
    '--- affectation ---
    ProchainNumeroPosteColmatageChaud = 0
    
    For a = POSTES.P_C31 To POSTES.P_C32
                
        If TEtatsPostes(a).Condamnation = False Then
        
            If TypeDeZone = ZONE_DEPART Then
                    
                '--- la zone est une zone de départ ---
                If NumCharge >= CHARGES.C_NUM_MINI And NumCharge <= CHARGES.C_NUM_MAXI Then
                    
                    '--- vérifier que le temps au poste est terminé ---
                    With TEtatsCharges(NumCharge)
                        If a = .TGammesAnodisation.TDetailsGammesAnodisation(.PtrZoneGammeAnodisation).NumPosteReel Then
                            If .TGammesAnodisation.TDetailsGammesAnodisation(.PtrZoneGammeAnodisation).FinDuTempsPosteReel = True Then
                                ProchainNumeroPosteColmatageChaud = a
                                Exit For
                            End If
                        End If
                    End With
                
                End If
                
            Else
    
                '--- la zone de colmatage est une zone d'arrivée ---
                If NumCharge = 0 Or (NumCharge >= CHARGES.C_NUM_MINI And NumCharge <= CHARGES.C_NUM_MAXI) Then
                
                    With TEtatsCharges(NumCharge)
                            
                        '--- prendre le premier poste d'arrivée vide ---
                        If TEtatsPostes(a).NumCharge = 0 Then
                            ProchainNumeroPosteColmatageChaud = a
                            Exit For
                        End If
                        
                    End With
                        
                End If
                                    
            End If
                                    
        End If
                                    
    Next a

    'LogCharge ("ProchainNumeroPosteColmatageChaud=" & ProchainNumeroPosteColmatageChaud)
End Function

'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Rôle      : Recherche le prochain numéro de poste de brillantage
' Entrées :                                    NumCharge -> numéro de la charge faisant l'objet de la demande
'                                                 TypeDeZone -> FALSE = La zone est une ZONE de DEPART
'                                                                            TRUE = La zone est une ZONE d'ARRIVEE
' Retours : ProchainNumeroPosteBrillantage ->                 0 = Pas de poste de brillantage disponible
'                                                                            C05 ou C07 = Numéro de poste du brillantage
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function ProchainNumeroPosteBrillantage(ByVal NumCharge As Integer, _
                                                                                 ByVal TypeDeZone As Boolean) As Integer

    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next
    
    '--- constantes privées ---
    Const ZONE_DEPART As Boolean = False
    Const ZONE_ARRIVEE As Boolean = True

    '--- déclaration ---
    Dim a As Integer
    
    '--- affectation ---
    ProchainNumeroPosteBrillantage = 0
    
    For a = POSTES.P_C05 To POSTES.P_C07
                
        Select Case a
        
            Case POSTES.P_C05, POSTES.P_C07
                '--- poste de brillantages uniquement ---
                If TEtatsPostes(a).Condamnation = False Then
                
                    If TypeDeZone = ZONE_DEPART Then
                            
                        '--- la zone est une zone de départ ---
                        If NumCharge >= CHARGES.C_NUM_MINI And NumCharge <= CHARGES.C_NUM_MAXI Then
                            
                            '--- vérifier que le temps au poste est terminé ---
                            With TEtatsCharges(NumCharge)
                                If a = .TGammesAnodisation.TDetailsGammesAnodisation(.PtrZoneGammeAnodisation).NumPosteReel Then
                                    If .TGammesAnodisation.TDetailsGammesAnodisation(.PtrZoneGammeAnodisation).FinDuTempsPosteReel = True Then
                                        ProchainNumeroPosteBrillantage = a
                                        Exit For
                                    End If
                                End If
                            End With
                        
                        End If
                        
                    Else
            
                        '--- la zone de colmatage est une zone d'arrivée ---
                        If NumCharge = 0 Or (NumCharge >= CHARGES.C_NUM_MINI And NumCharge <= CHARGES.C_NUM_MAXI) Then
                        
                            With TEtatsCharges(NumCharge)
                                    
                                '--- prendre le premier poste d'arrivée vide ---
                                If TEtatsPostes(a).NumCharge = 0 Then
                                    ProchainNumeroPosteBrillantage = a
                                    Exit For
                                End If
                                
                            End With
                                
                        End If
                                            
                    End If
                                            
                End If
                                    
            Case Else
        End Select
                                    
    Next a

End Function

'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Rôle      : Recherche le prochain numéro de poste du séchoir
' Entrées :                                       NumCharge -> numéro de la charge faisant l'objet de la demande
'                                                     TypeDeZone -> FALSE = La zone est une ZONE de DEPART
'                                                                                TRUE = La zone est une ZONE d'ARRIVEE
' Retours : ProchainNumeroPosteSechoir ->                  0 = Pas de poste de colmatage disponible
'                                                                       C33 ou C34 = Numéro de poste de colmatage à chaud
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function ProchainNumeroPosteSechoir(ByVal NumCharge As Integer, _
                                                                           ByVal TypeDeZone As Boolean) As Integer

    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next
    
    '--- constantes privées ---
    Const ZONE_DEPART As Boolean = False
    Const ZONE_ARRIVEE As Boolean = True

    '--- déclaration ---
    Dim a As Integer
    
    '--- affectation ---
    ProchainNumeroPosteSechoir = 0
    
    For a = POSTES.P_C33 To POSTES.P_C34
                
        If TEtatsPostes(a).Condamnation = False Then
        
            If TypeDeZone = ZONE_DEPART Then
                    
                '--- la zone est une zone de départ ---
                If NumCharge >= CHARGES.C_NUM_MINI And NumCharge <= CHARGES.C_NUM_MAXI Then
                    
                    '--- vérifier que le temps au poste est terminé ---
                    With TEtatsCharges(NumCharge)
                        If a = .TGammesAnodisation.TDetailsGammesAnodisation(.PtrZoneGammeAnodisation).NumPosteReel Then
                            If .TGammesAnodisation.TDetailsGammesAnodisation(.PtrZoneGammeAnodisation).FinDuTempsPosteReel = True Then
                                ProchainNumeroPosteSechoir = a
                                Exit For
                            End If
                        End If
                    End With
                
                End If
                
            Else
    
                '--- la zone du séchoir est une zone d'arrivée ---
                If NumCharge = 0 Or (NumCharge >= CHARGES.C_NUM_MINI And NumCharge <= CHARGES.C_NUM_MAXI) Then
                
                    With TEtatsCharges(NumCharge)
                            
                        '--- prendre le premier poste d'arrivée vide ---
                        If TEtatsPostes(a).NumCharge = 0 Then
                            ProchainNumeroPosteSechoir = a
                            Exit For
                        End If
                        
                    End With
                        
                End If
                                    
            End If
                                    
        End If
                                    
    Next a

End Function

'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Rôle      : Décompte un temps d'une charge dans un poste par rapport au temps de poste dans la gamme
' Entrées : NumPoste -> N° du poste ou l'on doit décompter le temps
' Retours :
' Détails  : La variable FinDuTempsPosteReel de la gamme d'anodisation monte quand le décompte du temps est
'                 inférieur ou égale à 0
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Sub DecompteDuTempsAuPosteSecondes(ByVal NumPoste As Integer)
    
    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next
    
    '--- constantes privées ---
    Const TEMPS_MOUVEMENT_AVANT_PRISE As Long = 5           'temps moyen correspondant à la fermeture des accroches et au début de montée
    
    '--- déclaration ---
    Dim NumCharge As Integer
    Dim TempsAuPosteSecondes As Long, _
            TempsDepuisEntreeDansLePosteSecondes As Long, _
            DecompteDuTempsSecondes As Long

    If NumPoste >= POSTES.P_CHGT_1 And NumPoste <= DERNIER_POSTE Then
    
        '--- affectation du numéro de charge ---
        NumCharge = TEtatsPostes(NumPoste).NumCharge
    
        If TEtatsPostes(NumPoste).DefinitionPoste.AvecTemps = True Then
        
            If NumCharge >= CHARGES.C_NUM_MINI And NumCharge <= CHARGES.C_NUM_MAXI Then

                '--- décompte du temps ---
                With TEtatsCharges(NumCharge)
                    If .PtrZoneGammeAnodisation > 0 And .NbrPostesTraites > 0 Then
                        If .TGammesAnodisation.TDetailsGammesAnodisation(.PtrZoneGammeAnodisation).NumPosteReel = .TDetailsFichesProduction(.NbrPostesTraites).NumPoste Then
                                        
                            '--- recherche du temps théorique dans la gamme ---
                            TempsAuPosteSecondes = .TGammesAnodisation.TDetailsGammesAnodisation(.PtrZoneGammeAnodisation).TempsAuPosteSecondes
                            
                            '--- calcul du temps à partir de la fiche de production ---
                            With .TDetailsFichesProduction(.NbrPostesTraites)
                                If .DateEntreePoste <> Empty And .DateSortiePoste <> Empty Then
                                    
                                    '--- calcul ---
                                    TempsDepuisEntreeDansLePosteSecondes = DateDiff("s", .DateEntreePoste, .DateSortiePoste)
                                    DecompteDuTempsSecondes = TempsAuPosteSecondes - TempsDepuisEntreeDansLePosteSecondes
                                    
                                    '--- affectation au bon endroit dans la gamme ---
                                    With TEtatsCharges(NumCharge)
                                        .TGammesAnodisation.TDetailsGammesAnodisation(.PtrZoneGammeAnodisation).DecompteDuTempsAuPosteReelSecondes = CStr(DecompteDuTempsSecondes)
                                        If (DecompteDuTempsSecondes - TEMPS_MOUVEMENT_AVANT_PRISE) <= 0 Then
                                            .TGammesAnodisation.TDetailsGammesAnodisation(.PtrZoneGammeAnodisation).FinDuTempsPosteReel = True
                                        End If
                                    End With
                                
                                End If
                            End With
                        
                        End If
                    End If
                End With

            End If

        End If
    
    End If

End Sub

'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Rôle      : Décompte un temps d'aletre dans un poste
' Entrées : NumPoste -> N° du poste ou l'on doit décompter le temps
' Retours :
' Détails  : La variable FinDuTempsPosteReel de la gamme d'anodisation monte quand le décompte du temps est
'                 inférieur ou égale à 0
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Sub DecompteDuTempsAlerteAuPosteSecondes(ByVal NumPoste As Integer)
    
    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next
    
    '--- constantes privées ---
    Const TEMPS_MOUVEMENT_AVANT_PRISE As Long = 5           'temps moyen correspondant à la fermeture des accroches et au début de montée
    
    '--- déclaration ---
    Dim DebutAlertePosteReel As Boolean                                       'indique le début de l'alerte au poste réel
    Dim UneAlerteEstEnCours As Boolean                                        'indique une alerte en cours
    
    Static MemAntiRebondKlaxon As Boolean                                  'mémoire anti-rebond de lancement du klaxon
    
    Dim NumCharge As Integer                                                          'représente un numéro de charge
    Dim NumPosteReel As Integer                                                     'représente un numéro de poste réel
    
    Dim TempsAlerteAuPosteSecondes As Long, _
            TempsDepuisEntreeDansLePosteSecondes As Long, _
            DecompteDuTempsSecondes As Long
    
    Dim TexteAlerte As String                                                            'texte de l'alerte destiné à l'afficheur

    If NumPoste >= POSTES.P_CHGT_1 And NumPoste <= DERNIER_POSTE Then
    
        '--- affectation ---
        NumCharge = TEtatsPostes(NumPoste).NumCharge
    
        If TEtatsPostes(NumPoste).DefinitionPoste.AvecTemps = True Then
        
            If NumCharge >= CHARGES.C_NUM_MINI And NumCharge <= CHARGES.C_NUM_MAXI Then

                '--- décompte du temps ---
                With TEtatsCharges(NumCharge)
                    
                    If .PtrZoneGammeAnodisation > 0 And .NbrPostesTraites > 0 Then
                        
                        If .TGammesAnodisation.TDetailsGammesAnodisation(.PtrZoneGammeAnodisation).NumPosteReel = .TDetailsFichesProduction(.NbrPostesTraites).NumPoste Then
                                        
                            '----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
                            
                            '--- recherche du temps théorique dans la gamme ---
                            TempsAlerteAuPosteSecondes = .TGammesAnodisation.TDetailsGammesAnodisation(.PtrZoneGammeAnodisation).TempsAlerteSecondes
                            
                            '--- affectation du bit de début d'alerte au poste ---
                            DebutAlertePosteReel = .TGammesAnodisation.TDetailsGammesAnodisation(.PtrZoneGammeAnodisation).DebutAlertePosteReel
                            
                            '----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
                            
                            '--- gestion de l'alerte ---
                            If DebutAlertePosteReel = True Then
                            
                                '--- affectation du numéro du poste ---
                                NumPosteReel = .TGammesAnodisation.TDetailsGammesAnodisation(.PtrZoneGammeAnodisation).NumPosteReel
                                
                                '--- affectation du texte de l'alerte ---
                                TexteAlerte = "ALERTE " & TEtatsPostes(NumPosteReel).DefinitionPoste.NomPoste & ": " & CTemps(.TGammesAnodisation.TDetailsGammesAnodisation(.PtrZoneGammeAnodisation).DecompteDuTempsAuPosteReelSecondes)
                                
                                '--- affectation indiquant l'alerte en cours ---
                                UneAlerteEstEnCours = True
                            
                            End If
                            
                            '----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
                            
                            '--- calcul du temps à partir de la fiche de production ---
                            With .TDetailsFichesProduction(.NbrPostesTraites)
                                
                                If .DateEntreePoste <> Empty And .DateSortiePoste <> Empty And TempsAlerteAuPosteSecondes > 0 And DebutAlertePosteReel = False Then
                                    
                                    '--- calcul ---
                                    TempsDepuisEntreeDansLePosteSecondes = DateDiff("s", .DateEntreePoste, .DateSortiePoste)
                                    DecompteDuTempsSecondes = TempsAlerteAuPosteSecondes - TempsDepuisEntreeDansLePosteSecondes
                                    
                                    '--- affectation au bon endroit dans la gamme ---
                                    With TEtatsCharges(NumCharge)
                                        
                                        '--- affectation du temps d'alerte réel en secondes ---
                                        .TGammesAnodisation.TDetailsGammesAnodisation(.PtrZoneGammeAnodisation).DecompteDuTempsAlerteReelSecondes = CStr(DecompteDuTempsSecondes)
                                        
                                        '--- montée du drapeau d'alerte ---
                                        If DecompteDuTempsSecondes <= 0 Then
                                            .TGammesAnodisation.TDetailsGammesAnodisation(.PtrZoneGammeAnodisation).DebutAlertePosteReel = True
                                        End If
                                    
                                    End With
                                
                                End If
                            
                            End With
                        
                            '----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
                        
                        End If
                    
                    End If
                
                End With

            End If

        End If
                    
        '----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
            
        '--- affichage de l'alerte sur l'afficheur ---
        If UneAlerteEstEnCours = True Then
                                
            '--- alerte en cours donc affectation de la priorité d'affichage des alertes ---
            PrioriteAfficheurPourAlertes = True
                                
            '--- affichage du message d'alerte ---
            Bidon = MessageAfficheur("B", TexteAlerte)
                                
            '--- lancement du klaxon ---
            If MemAntiRebondKlaxon = False Then
                
                '--- montée du bit du klaxon dans l'automate ---
                Bidon = APIEcritureVariableNommee("DEFAUTS", "M_Dem_PC_Klaxon", True)
        
                '--- affectation de la mémoire anti-rebond du klaxon ---
                MemAntiRebondKlaxon = True
        
            End If
        
        Else
    
            '--- pas d'alerte donc RAZ de la priorité d'affichage des alertes ---
            PrioriteAfficheurPourAlertes = False
            
            If MemAntiRebondKlaxon = True Then
            
                '--- effacement de l'afficheur ---
                Bidon = MessageAfficheur("B", "")
                    
                '--- RAZ de la mémoire anti-rebond du klaxon ---
                MemAntiRebondKlaxon = False
    
            End If
    
        End If
    
    End If

End Sub

Public Function LogCharge(ByVal msg As String)

  Dim nUnit As Integer
  nUnit = FreeFile
  ' This assumes write access to the directory containing the program '
  ' You will need to choose another directory if this is not possible '
  Dim str As String
  str = Format(Now, "yyyymmdd")
  Open App.Path & "\" & str & "_LOG.txt" For Append As nUnit
  ' For Append As nUnit
  Print #nUnit, "  " & msg
  Print #nUnit, Format$(Now)
  Print #nUnit, " --------------------------------------- " '& Format$(Now)
  Close nUnit
End Function

'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Rôle      : Recherche le prochain numéro de poste valide (la ou on peut prendre ou déposé une charge)
' Entrées :                            NumCharge -> numéro de la charge faisant l'objet de la demande
'                                               NumZone -> N° de la zone (départ ou arrivée)
'                                          TypeDeZone -> FALSE = La zone est une ZONE de DEPART
'                                                                     TRUE = La zone est une ZONE d'ARRIVEE
' Retours : ProchainNumeroPosteValide -> 0 = Pas de poste de valide sinon le numéro du poste
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function ProchainNumeroPosteValide(ByVal NumCharge As Integer, _
                                                                         ByVal NumZone As Integer, _
                                                                         ByVal TypeDeZone As Boolean) As Integer
    
    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next

    '--- constantes privées ---
    Const ZONE_DEPART As Boolean = False
    Const ZONE_ARRIVEE As Boolean = True
    
    '--- déclaration ---
    Dim NumPremierPosteZone As Integer, _
           NumDernierPosteZone As Integer
    
    '--- affectation ---
    ProchainNumeroPosteValide = 0
    
 
   
    'LogCharge ("ProchainNumeroPosteValide")
        
    If NumZone >= LIMITE_BASSE_ZONES And NumZone <= LIMITE_HAUTE_ZONES Then
    
        'FSynoptique.TextInfo = ("numzone ok")
        '--- affectation ---
        NumPremierPosteZone = TZones(NumZone).NumPremierPoste
        NumDernierPosteZone = TZones(NumZone).NumDernierPoste
        
        'LogCharge ("NumPremierPosteZone=" & NumPremierPosteZone)
        'LogCharge ("NumDernierPosteZone=" & NumDernierPosteZone)
        
        If TypeDeZone = ZONE_DEPART Then
    
            '**********************************************************************************************************
            '                                               LA ZONE EST UNE ZONE DE DEPART
            '**********************************************************************************************************
            'LogCharge (" LA ZONE EST UNE ZONE DE DEPART")
            If NumPremierPosteZone = POSTES.P_CHGT_1 And NumDernierPosteZone = POSTES.P_CHGT_2 Then
            'SZP2024
            'ElseIf NumPremierPosteZone = POSTES.P_C05 And NumDernierPosteZone = POSTES.P_C07 Then
                'LogCharge ("zone de brillantage  DEPART")
                '--- zone de brillantage ---
            '    ProchainNumeroPosteValide = ProchainNumeroPosteBrillantage(NumCharge, False)
            
            ElseIf NumPremierPosteZone = POSTES.P_C13 And NumDernierPosteZone = POSTES.P_C16 Then
                
                '--- zone d'anodisation ---
                ProchainNumeroPosteValide = ProchainNumeroPosteAnodisation(NumCharge, False)
                
            ' DEBUT MODIF 20200120 SZP
            'ElseIf NumPremierPosteZone = POSTES.P_C32 And NumDernierPosteZone = POSTES.P_C33 Then
            '    LogCharge ("ProchainNumeroPosteColmatageChaud DEPART")
                '--- colmatage chaud ---
            '    ProchainNumeroPosteValide = ProchainNumeroPosteColmatageChaud(NumCharge, False)
            ElseIf NumPremierPosteZone = POSTES.P_C31 And NumDernierPosteZone = POSTES.P_C32 Then
            
                '--- colmatage chaud ---
                ProchainNumeroPosteValide = ProchainNumeroPosteColmatageChaud(NumCharge, False)
            
            'ElseIf NumPremierPosteZone = POSTES.P_C33 And NumDernierPosteZone = POSTES.P_C34 Then
            
                '--- séchoir (poste 1 et 2) ---
                'ProchainNumeroPosteValide = ProchainNumeroPosteSechoir(NumCharge, False)
            
            
            ' FIN MODIF 20200120 SZP
            ElseIf NumPremierPosteZone = POSTES.P_D1 And NumDernierPosteZone = POSTES.P_D2 Then
                
                '--- zone du déchargement ---
            
            ElseIf NumPremierPosteZone = NumDernierPosteZone Then
                            
                '--- toutes les autres zones (poste simple) ---
                If TEtatsPostes(NumPremierPosteZone).Condamnation = False Then
                    If NumCharge >= CHARGES.C_NUM_MINI And NumCharge <= CHARGES.C_NUM_MAXI Then
                        
                        '--- vérifier que le temps au poste est terminé ---
                        With TEtatsCharges(NumCharge)
                            If NumPremierPosteZone = .TGammesAnodisation.TDetailsGammesAnodisation(.PtrZoneGammeAnodisation).NumPosteReel Then
                                If .TGammesAnodisation.TDetailsGammesAnodisation(.PtrZoneGammeAnodisation).FinDuTempsPosteReel = True Then
                                    ProchainNumeroPosteValide = NumPremierPosteZone
                                End If
                            End If
                        End With
                    
                    End If
                End If
            
            Else
                
                '--- CAS NORMALLEMENT IMPOSSIBLE ---
            
            End If
        
        Else
    
            '**********************************************************************************************************
            '                                               LA ZONE EST UNE ZONE D'ARRIVEE
            '**********************************************************************************************************
            'LogCharge (" LA ZONE EST UNE ZONE D'ARRIVEE")
       
            If NumPremierPosteZone = POSTES.P_CHGT_1 And NumDernierPosteZone = POSTES.P_CHGT_2 Then
            
                '--- zone de chargement (cas du déchargement à un des postes de chargement) ---
                ProchainNumeroPosteValide = ProchainNumeroPosteChargement()
            
            'SZP2024
            'ElseIf NumPremierPosteZone = POSTES.P_C05 And NumDernierPosteZone = POSTES.P_C07 Then
                'LogCharge ("zone de brillantage  ARRIVEE")
                '--- zone de brillantage ---
             '   ProchainNumeroPosteValide = ProchainNumeroPosteBrillantage(NumCharge, True)
            
            ElseIf NumPremierPosteZone = POSTES.P_C13 And NumDernierPosteZone = POSTES.P_C16 Then
                
                '--- zone d'anodisation ---
                ProchainNumeroPosteValide = ProchainNumeroPosteAnodisation(NumCharge, True)
            
            ' DEBUT MODIF 20200120 SZP **********************************************************************
            ElseIf NumPremierPosteZone = POSTES.P_C31 And NumDernierPosteZone = POSTES.P_C32 Then
                'LogCharge ("ProchainNumeroPosteColmatageChaud ARRIVEE")
                '--- colmatage chaud ---
                ProchainNumeroPosteValide = ProchainNumeroPosteColmatageChaud(NumCharge, True)
            'ElseIf NumPremierPosteZone = POSTES.P_C31 And NumDernierPosteZone = POSTES.P_C32 Then
            
                '--- colmatage chaud ---
                'ProchainNumeroPosteValide = ProchainNumeroPosteColmatageChaud(NumCharge, True)
            
            'ElseIf NumPremierPosteZone = POSTES.P_C33 And NumDernierPosteZone = POSTES.P_C34 Then
            
                '--- séchoir (poste 1 et 2) ---
                'ProchainNumeroPosteValide = ProchainNumeroPosteSechoir(NumCharge, True)
            ' FIN MODIF 20200120 SZP ***********************************************************************
            
            
            ElseIf NumPremierPosteZone = POSTES.P_D1 And NumDernierPosteZone = POSTES.P_D2 Then
                
                '--- zone du déchargement ---
                ProchainNumeroPosteValide = ProchainNumeroPosteDechargement()
            
            ElseIf NumPremierPosteZone = NumDernierPosteZone Then
            
                '--- toutes les autres zones (poste simple) ---
                With TEtatsPostes(NumPremierPosteZone)
                    If .Condamnation = False Then
                        If .NumCharge = 0 Then                    'vérifier si le poste d'arrivée est vide
                            ProchainNumeroPosteValide = NumPremierPosteZone
                        End If
                    End If
                End With
            
            Else
            
                '--- CAS NORMALLEMENT IMPOSSIBLE ---
                
            End If
        
        End If
    
    End If
    
End Function

'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Rôle      : Recherche le prochain numéro théorique d'un poste d'arrivée
' Entrées :                             NumCharge -> numéro de la charge faisant l'objet de la demande
'                                    NumZoneArrivee -> N° de la zone (départ ou arrivée)
' Retours : ProchainNumeroPosteValide -> 0 = Pas de poste de valide sinon le numéro du poste
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function ProchainNumTheoriquePosteArrivee(ByVal NumCharge As Integer, _
                                                                                     ByVal NumZoneArrivee As Integer)
    
    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next

    '--- constantes privées ---
    
    '--- déclaration ---
    Dim NumPremierPosteZone As Integer, _
           NumDernierPosteZone As Integer
    
    '--- affectation ---
    ProchainNumTheoriquePosteArrivee = 0
        
    If NumZoneArrivee >= LIMITE_BASSE_ZONES And NumZoneArrivee <= LIMITE_HAUTE_ZONES Then
    
        '--- affectation ---
        NumPremierPosteZone = TZones(NumZoneArrivee).NumPremierPoste
        NumDernierPosteZone = TZones(NumZoneArrivee).NumDernierPoste
        
        '--- analyse de la zone d'arrivée ---
        If NumPremierPosteZone = POSTES.P_CHGT_1 And NumDernierPosteZone = POSTES.P_CHGT_2 Then
        
            '--- zone de chargement (cas du déchargement à un des postes de chargement) ---
            ProchainNumTheoriquePosteArrivee = ProchainNumeroPosteChargement()
            
        ElseIf NumPremierPosteZone = POSTES.P_C13 And NumDernierPosteZone = POSTES.P_C16 Then
            
            '--- zone d'anodisation ---
            ProchainNumTheoriquePosteArrivee = ProchainNumeroPosteAnodisation(NumCharge, True)
        ' DEBUT MODIF SZV 20200120 ---------------------------------------------------------------------------
         ElseIf NumPremierPosteZone = POSTES.P_C32 And NumDernierPosteZone = POSTES.P_C33 Then
        
            '--- colmatage chaud ---
            ProchainNumTheoriquePosteArrivee = ProchainNumeroPosteColmatageChaud(NumCharge, True)
            
        'ElseIf NumPremierPosteZone = POSTES.P_C31 And NumDernierPosteZone = POSTES.P_C32 Then
        
            '--- colmatage chaud ---
            'ProchainNumTheoriquePosteArrivee = ProchainNumeroPosteColmatageChaud(NumCharge, True)
        
        'ElseIf NumPremierPosteZone = POSTES.P_C33 And NumDernierPosteZone = POSTES.P_C34 Then
        
            '--- séchoir (poste 1 et 2) ---
            'ProchainNumTheoriquePosteArrivee = ProchainNumeroPosteSechoir(NumCharge, True)
        ' FIN MODIF SZV 20200120 -----------------------------------------------------
        ElseIf NumPremierPosteZone = POSTES.P_D1 And NumDernierPosteZone = POSTES.P_D2 Then
            
            '--- zone du déchargement ---
            ProchainNumTheoriquePosteArrivee = ProchainNumeroPosteDechargement()
        
        ElseIf NumPremierPosteZone = NumDernierPosteZone Then
        
            '--- toutes les autres zones (poste simple) ---
            With TEtatsPostes(NumPremierPosteZone)
                If .Condamnation = False Then
                    '--- ne pas vérifier si le poste d'arrivée est vide car le poste d'arrivée est bien celui la malgré la présence d'une charge ---
                    ProchainNumTheoriquePosteArrivee = NumPremierPosteZone
                End If
            End With
        
        End If
    
    Else
        
        '--- CAS NORMALLEMENT IMPOSSIBLE ---
            
    End If
        
End Function

'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Rôle      : Transfert un cycle d'un pont vers l'automate
' Entrées :            NumPont -> Fonction de l'énumération PONTS
'                      TCyclePont() -> Tableau contenant le cycle du pont à transférer
' Retours : EnvoiCyclePont -> OK = tout va bien
'                                               ERREUR_COMMUNICATION_API = indique une erreur de communication avec l'API
'                                               CYCLE_DEJA_EN_COURS = l'envoi n'est pas possible car un cycle est déjà en cours
'                                               PONT_NON_AUTOMATIQUE = l'envoi n'est pas possible car le pont n'est pas en auto
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function EnvoiCyclePont(ByVal NumPont As Integer, _
                                                    ByRef TCyclePont() As Integer) As String
    
    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next
    
    '--- constantes privées ---
    
    '--- déclaration ---
    Dim a As Integer                                                                                    'pour les boucles FOR...NEXT
    Dim PtrAction As Integer                                                                        'pointeur d'une action
    Dim TValeurs(0 To NBR_LIGNES_CYCLES_PONTS - 1) As Integer      'tableau contenant les valeurs
    
    Dim RefGroupe As Long                                                                         'référence sur le groupe
    Dim TRefElements(0 To NBR_LIGNES_CYCLES_PONTS - 1) As Long 'référence sur un élément
    Dim NbrValeurs As Long                                                                        'nombre de valeurs
    Dim ValeurRetourneeAPI As Long                                                          'valeur retournée par une fonction
                                                                                                                    'concernant le dialogue avec l'automate
    
    Dim NomGroupe As String                                                                      'représente un nom de groupe
    Dim AdrPtrAction As String                                                                      'adresse du pointeur des actions

    Dim TEtatsCommunication As Variant                                                    'tableau contenant les états de la
                                                                                                                    'communication de chaque valeur transmise
    
    '--- affectation ---
    EnvoiCyclePont = ""

    '--- affectation en fonction du pont ---
    Select Case NumPont
        Case PONTS.P_1
            NomGroupe = "ACTIONS_P1"
            AdrPtrAction = "MW108"
        Case PONTS.P_2
            NomGroupe = "ACTIONS_P2"
            AdrPtrAction = "MW128"
        Case Else
            Exit Function
    End Select
    NbrValeurs = NBR_LIGNES_CYCLES_PONTS
    
    '--- contrôle sur la dimension du tableau ---
    If LBound(TCyclePont()) <> 1 And UBound(TCyclePont()) <> NBR_LIGNES_CYCLES_PONTS Then
        Exit Function
    End If
    
    If PROGRAMME_AVEC_AUTOMATE = True Then
    
        If TEtatsPonts(NumPont).ModePont = MODES_PONTS.M_AUTOMATIQUE Then
    
            '--- lecture du pointeur de l'action ---
            ValeurRetourneeAPI = APILectureMot(AdrPtrAction, PtrAction)
            
            If ValeurRetourneeAPI = 0 Then
    
                If PtrAction = 0 Then
        
                    '--- référence sur le groupe ---
                    RefGroupe = OccFPrincipale.AOCFPrincipale.GetGroupRef(RefServeur, NomGroupe)
                    If RefGroupe <= 0 Then
                        EnvoiCyclePont = ERREUR_COMMUNICATION_API
                        Exit Function
                    End If
        
                    '--- affectation du tableaux des références et des valeurs ---
                    For a = 1 To NBR_LIGNES_CYCLES_PONTS
        
                        '--- affectation de chaque élément ---
                        TRefElements(a - 1) = OccFPrincipale.AOCFPrincipale.GetItemRef(RefGroupe, "Action" & Right("00" & a, 2))
        
                        '--- affectation des valeurs ---
                        TValeurs(a - 1) = TCyclePont(a)
                    
                    Next a
        
                    '--- écriture dans l'automate ---
                    ValeurRetourneeAPI = OccFPrincipale.AOCFPrincipale.Write(NbrValeurs, TRefElements, TValeurs, TEtatsCommunication)
                    
                
                    If ValeurRetourneeAPI = 0 Then
                
                        '--- attente pour être sûr que la table soit déjà dans l'automate ---
                        Call Sleep(300)
                
                        '--- ecriture du pointeur de l'action ---
                        ValeurRetourneeAPI = APIEcritureMot(AdrPtrAction, 1)
                
                        If ValeurRetourneeAPI = 0 Then
                
                            '--- tout OK ---
                            EnvoiCyclePont = OK
                
                        Else
                        
                            '--- erreur de communication avec l'automate ---
                            EnvoiCyclePont = ERREUR_COMMUNICATION_API
                    
                        End If
                
                    Else
                    
                        '--- erreur de communication avec l'automate ---
                        EnvoiCyclePont = ERREUR_COMMUNICATION_API
                
                    End If
        
                Else
            
                    '--- cycle déja en cours ---
                    EnvoiCyclePont = CYCLE_DEJA_EN_COURS
            
                End If

            Else
            
                '--- erreur de communication avec l'automate ---
                EnvoiCyclePont = ERREUR_COMMUNICATION_API
        
            End If
    
        Else
    
            '--- le pont n'est pas en automatique ---
            EnvoiCyclePont = PONT_NON_AUTOMATIQUE
    
        End If
    Else
         EnvoiCyclePont = OK

    End If

End Function

'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Rôle      : Insertion d'un temps d'égouttage dans un cycle destiné à un pont
' Entrées :                         TempsEgouttageSecondes -> Représente le temps d'égouttage en secondes
'                                                               TCyclePont() -> Tableau contenant le cycle du pont à modifier
' Retours : InsertionTempsEgouttageDansCyclePont -> OK = la valeur a été inséré
'                                                                                        ""   = l'action TEMPO_EGOUT n'a pas été trouvé
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function InsertionTempsEgouttageDansCyclePont(ByVal TempsEgouttageSecondes As Integer, _
                                                                                             ByRef TCyclePont() As Integer) As String
    
    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next

    '--- déclaration ---
    Dim a As Integer, _
           NumAction As Integer
    
    '--- affectation ---
    InsertionTempsEgouttageDansCyclePont = ""
    
    '--- recherche de l'action du temps d'égouttage et enregistrement dans le cycle ---
    For a = LBound(TCyclePont()) To UBound(TCyclePont())
    
        '--- affectation ---
        NumAction = TCyclePont(a)
    
        If NumAction >= NUM_ACTION_NOP And NumAction <= NUM_ACTION_FCY Then
    
            '--- fixer la valeur ---
            If TActions(NumAction).CodeAction = CODE_TEMPO_EGOUTTAGE And a < NBR_LIGNES_CYCLES_PONTS Then
                TCyclePont(Succ(a)) = TempsEgouttageSecondes
                InsertionTempsEgouttageDansCyclePont = OK
                Exit For
            End If
        
            '--- incrément si paramètre pour pointer toujours une action ---
            If TActions(NumAction).ParametreOuiNon = True Then
                Inc a
            End If
        
        End If
    
    Next a

End Function

'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Rôle      : Insertion d'un délai supplémentaire de stabilisation de la charge dans un cycle destiné à un pont
' Entrées :                        DelaiSupStabilisationChargeSecondes -> Représente le délai supplémentaire
'                                                                                                           de stabilisation de la charge en secondes
'                                                                                  TCyclePont() -> Tableau contenant le cycle du pont à modifier
' Retours : InsertionDelaiSupStabilisationChargeDansCyclePont -> OK = la valeur a été inséré
'                                                                                                            ""  = l'action TEMPO_STAB n'a pas été trouvé
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function InsertionDelaiSupStabilisationChargeDansCyclePont(ByVal DelaiSupStabilisationChargeSecondes As Integer, _
                                                                                                                ByRef TCyclePont() As Integer) As String
    
    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next

    '--- déclaration ---
    Dim a As Integer, _
           NumAction As Integer
    
    '--- affectation ---
    InsertionDelaiSupStabilisationChargeDansCyclePont = ""
    
    '--- recherche de l'action temporisation de stabilisation et enregistrement dans le cycle ---
    For a = LBound(TCyclePont()) To UBound(TCyclePont())
    
        '--- affectation ---
        NumAction = TCyclePont(a)
    
        If NumAction >= NUM_ACTION_NOP And NumAction <= NUM_ACTION_FCY Then
    
            '--- fixer la valeur ---
            If TActions(NumAction).CodeAction = CODE_TEMPO_STABILISATION And a < NBR_LIGNES_CYCLES_PONTS Then
                TCyclePont(Succ(a)) = TCyclePont(Succ(a)) + DelaiSupStabilisationChargeSecondes
                InsertionDelaiSupStabilisationChargeDansCyclePont = OK
                Exit For
            End If
        
            '--- incrément si paramètre pour pointer toujours une action ---
            If TActions(NumAction).ParametreOuiNon = True Then
                Inc a
            End If
        
        End If
    
    Next a

End Function

'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Rôle      : Transmet le déplacement du pont au poste voulu en AUTOMATIQUE
' Entrées :                                  NumPont -> Numéro du pont concerné
'                                                NumPoste -> Numéro du poste souhaité
'                                      CouleurReponse -> Couleur de la réponse
' Retours : AutomatiqueDeplacementPont -> Message à retourner comme réponse
' Détails  :
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function AutomatiqueDeplacementPont(ByVal NumPont As Integer, _
                                                                           ByVal NumPoste As Integer, _
                                                                           ByRef CouleurReponse As Long) As String
    
    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next
    
    '--- constantes privées ---
    
    Dim Texte As String
    Texte = "AutomatiqueDeplacementPont " & NumPont & ", NumPoste: " & NumPoste
    AfficheRenseignementsDebug CouleurReponse, Texte & vbCrLf
    
    If NumPoste >= POSTES.P_C13 And NumPoste <= POSTES.P_C16 Then
        Call Log(Texte)
    End If
    
    
    '--- déclaration ---
    Dim TUnCyclePont(1 To NBR_LIGNES_CYCLES_PONTS) As Integer
    Dim Reponse As String, _
            ReponseEnvoiCyclePont As String
    
    '--- affectation ---
    AutomatiqueDeplacementPont = ""
    
    If NumPont = PONTS.P_1 Or NumPont = PONTS.P_2 Then
                    
        '--- vérification si le système cyclique ou IA dispose du contrôle du pont ---
        ' ATTENTION si l'opérateur dispose du contrôle des 2 ponts on utilise la fonction pour l'interprétation
        ' des commandes en passant par la fonction GestionCommandesOperateur
        If TEtatsPonts(NumPont).ControleParOperateur = False Or _
            (TEtatsPonts(PONTS.P_1).ControleParOperateur = True And _
            TEtatsPonts(PONTS.P_2).ControleParOperateur = True) And _
            TEtatsPonts(NumPont).PtrEtActionEnCoursAPI.PtrAction = 0 Then
        
            If NumPoste >= POSTES.P_CHGT_1 And NumPoste <= DERNIER_POSTE Then
                        
                '--- affectation ---
                Reponse = NOUVELLE_LIGNE & "DEPLACEMENT DU PONT " & _
                                  NumPont & _
                                   " EN " & _
                                  TEtatsPostes(NumPoste).DefinitionPoste.NomPoste & _
                                  vbCrLf & TABULATION_REPONSES
                
                '--- construction du cycle ---
                Erase TUnCyclePont()
                TUnCyclePont(1) = NumPoste
                TUnCyclePont(2) = NUM_ACTION_FCY
                
                
                '--- lancement du déplacement ---
                ReponseEnvoiCyclePont = EnvoiCyclePont(NumPont, TUnCyclePont)
                Select Case ReponseEnvoiCyclePont
                    
                    Case OK
                         '--- le cycle a été transféré avec succès, il faut remplir la fiche des paramètres ---
                         With TEtatsPonts(NumPont).TParametresCyclesPonts(CYCLES.C_ACTUEL)
                            .NumPosteDepart = TEtatsPonts(NumPont).PosteActuel
                            .NumPosteArrivee = NumPoste
                            .TypeCycle = TYPES_CYCLES.TC_DEPLACEMENT_PONT
                            .DelaiSupStabilisationChargeSecondes = 0
                            .TempsEgouttageSecondes = 0
                         End With
                        
                        '--- affectation de la réponse ---
                        CouleurReponse = COULEURS.BLEU_3
                        Reponse = OK
                    
                    Case Else
                        '--- le déplacement a été refusé / affectation de la réponse ---
                        CouleurReponse = COULEURS.ROUGE_3
                        Reponse = Reponse & "Déplacement au poste " & TEtatsPostes(NumPoste).DefinitionPoste.NomPoste & " REFUSE"
                        Reponse = Reponse & vbCrLf & TABULATION_REPONSES & ReponseEnvoiCyclePont
                
                End Select
                    
            Else
            
                '--- mauvaise formulation / affectation de la réponse ---
                CouleurReponse = COULEURS.ROUGE_3
                Reponse = MAUVAISE_FORMULATION
            
            End If
    
        Else
    
            '--- pas de disposition du pont / affectation de la réponse ---
            CouleurReponse = COULEURS.ROUGE_3
            Reponse = PAS_DE_DISPOSITION_DU_PONT_IA & " " & NumPont
    
        End If
    
    Else
        
        '--- mauvaise formulation / affectation de la réponse ---
        CouleurReponse = COULEURS.ROUGE_3
        Reponse = MAUVAISE_FORMULATION
    
    End If

    '--- valeur de retour ---
    AutomatiqueDeplacementPont = Reponse

End Function

'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Rôle      : Transmet le déplacement du pont au poste voulu en AUTOMATIQUE pour l'optimisation avant la fin
'                 d'un temps au poste
' Entrées :                                  NumPont -> Numéro du pont concerné
'                                                NumPoste -> Numéro du poste souhaité
'                                      CouleurReponse -> Couleur de la réponse
' Retours : AutomatiqueDeplacementPont -> Message à retourner comme réponse
' Détails  :
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function AutomatiqueDeplacementPontOptimisation(ByVal NumPont As Integer, _
                                                                                                ByVal NumPoste As Integer, _
                                                                                                ByRef CouleurReponse As Long) As String
    
    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next
    
    '--- constantes privées ---
    
    
    'Call Log("AutomatiqueDeplacementPontOptimisation entrée  pont:" & NumPont & " , NumPoste: " & NumPoste)
    '--- déclaration ---
    Dim TUnCyclePont(1 To NBR_LIGNES_CYCLES_PONTS) As Integer
    Dim Reponse As String, _
            ReponseEnvoiCyclePont As String
    
    '--- affectation ---
    AutomatiqueDeplacementPontOptimisation = ""
    
    If NumPont = PONTS.P_1 Or NumPont = PONTS.P_2 Then
                    
        '--- vérification si le système cyclique ou IA dispose du contrôle du pont ---
        ' ATTENTION si l'opérateur dispose du contrôle des 2 ponts on utilise la fonction pour l'interprétation
        ' des commandes en passant par la fonction GestionCommandesOperateur
        If TEtatsPonts(NumPont).ControleParOperateur = False Or _
            (TEtatsPonts(PONTS.P_1).ControleParOperateur = True And _
            TEtatsPonts(PONTS.P_2).ControleParOperateur = True) And _
            TEtatsPonts(NumPont).PtrEtActionEnCoursAPI.PtrAction = 0 Then
        
            If NumPoste >= POSTES.P_CHGT_1 And NumPoste <= DERNIER_POSTE Then
                        
                '--- ne pas faire de déplacement en optimisation si le poste est vide ou condamné ---
                If TEtatsPostes(NumPoste).NumCharge > 0 And TEtatsPostes(NumPoste).Condamnation = False Then
                
                    '--- affectation ---
                    Reponse = NOUVELLE_LIGNE & "DEPLACEMENT DU PONT " & _
                                      NumPont & _
                                       " EN " & _
                                      TEtatsPostes(NumPoste).DefinitionPoste.NomPoste & _
                                      vbCrLf & TABULATION_REPONSES
                    
                    '--- construction du cycle ---
                    Erase TUnCyclePont()
                    TUnCyclePont(1) = NumPoste
                    TUnCyclePont(2) = NUM_ACTION_FCY
                    
                    '--- lancement du déplacement ---
                    ReponseEnvoiCyclePont = EnvoiCyclePont(NumPont, TUnCyclePont)
                    Select Case ReponseEnvoiCyclePont
                        
                        Case OK
                             '--- le cycle a été transféré avec succès, il faut remplir la fiche des paramètres ---
                             With TEtatsPonts(NumPont).TParametresCyclesPonts(CYCLES.C_ACTUEL)
                                .NumPosteDepart = TEtatsPonts(NumPont).PosteActuel
                                .NumPosteArrivee = NumPoste
                                .TypeCycle = TYPES_CYCLES.TC_DEPLACEMENT_PONT
                                .DelaiSupStabilisationChargeSecondes = 0
                                .TempsEgouttageSecondes = 0
                             End With
                            
                            '--- affectation de la réponse ---
                            CouleurReponse = COULEURS.BLEU_3
                            Reponse = OK
                        
                        Case Else
                            '--- le déplacement a été refusé / affectation de la réponse ---
                            CouleurReponse = COULEURS.ROUGE_3
                            Reponse = Reponse & "Déplacement au poste " & TEtatsPostes(NumPoste).DefinitionPoste.NomPoste & " REFUSE"
                            Reponse = Reponse & vbCrLf & TABULATION_REPONSES & ReponseEnvoiCyclePont
                    
                    End Select
                                        
                Else
                            
                    '--- affectation de la réponse ---
                    CouleurReponse = COULEURS.BLEU_3
                    Reponse = "DEPLACEMENT INUTILE DU PONT " & NumPont
                    
                End If
                    
            Else
            
                '--- mauvaise formulation / affectation de la réponse ---
                CouleurReponse = COULEURS.ROUGE_3
                Reponse = MAUVAISE_FORMULATION
            
            End If
    
        Else
    
            '--- pas de disposition du pont / affectation de la réponse ---
            CouleurReponse = COULEURS.ROUGE_3
            Reponse = PAS_DE_DISPOSITION_DU_PONT_IA & " " & NumPont
    
        End If
    
    Else
        
        '--- mauvaise formulation / affectation de la réponse ---
        CouleurReponse = COULEURS.ROUGE_3
        Reponse = MAUVAISE_FORMULATION
    
    End If
     'Call Log("AutomatiqueDeplacementPontOptimisation sortie")
    '--- valeur de retour ---
    AutomatiqueDeplacementPontOptimisation = Reponse

End Function

'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Rôle      : Transfert d'une charge au poste voulu en AUTOMATIQUE
' Entrées :                                     NumPontImpose -> Numéro du pont imposé
'                                                                                    0 = pas de pont imposé, prendre le numéro du pont IA pour
'                                                                                          effectuer le transfert
'                                                                                    <> 0, pont imposé, prendre ce numéro de pont pour
'                                                                                          effectuer le transfert
'                                                    NumPosteDepart -> Numéro du poste de départ
'                                                   NumPosteArrivee -> Numéro du poste d'arrivée
'                                    TempsEgouttageSecondes -> Temps d'égouttage en secondes
'                 DelaiSupStabilisationChargeSecondes -> Délai de stabilisation supplémentaire de la charge
'                                                                                    en secondes
' Retours :                      NumPontReelDuTransfert -> numéro du pont réel qui va effectuer le transfert
'                                                    CouleurReponse -> Couleur de la réponse
'                                 AutomatiqueTransfertCharge -> Message à retourner comme réponse
' Détails  :
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function AutomatiqueTransfertCharge(ByVal NumPontImpose As Integer, _
                                                                         ByVal NumPosteDepart As Integer, _
                                                                         ByVal NumPosteArrivee As Integer, _
                                                                         ByVal TempsEgouttageSecondes As Integer, _
                                                                         ByVal DelaiSupStabilisationChargeSecondes As Integer, _
                                                                         ByRef NumPontReelDuTransfert As Integer, _
                                                                         ByRef CouleurReponse As Long) As String
    
    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next
    
    '--- constantes privées ---
    
    '--- déclaration ---
    Dim NumPont As Integer                                                                      'numéro du pont donnée par les diagrammes
                                                                                                                  'en cyclique (extrait de la prémisse)
    Dim NumPontIA As Integer                                                                   'numéro de pont donné par le moteur
                                                                                                                  'd'inférence (extrait de la prémisse)
    Dim TUnCyclePont(1 To NBR_LIGNES_CYCLES_PONTS) As Integer 'cycle d'un pont avec tous les temps à
                                                                                                                  'envoyer à l'automate
    Dim TempsCycleSecondes As Long                                                    'temps d'un cycle en secondes
    
    Dim Reponse As String                                                                        'correspond à la variable de retour de la
                                                                                                                  'fonction
    Dim ReponseExtraitPremisseDecodee As String                                'correspond à la réponse donnée à l'extraction
                                                                                                                  'd'une prémisse décodée
    Dim ReponseEnvoiCyclePont As String                                               'correspond à la réponse donnée à l'envoi
                                                                                                                  'd'un cycle d'un pont
            
    '--- affectation par défaut ---
    AutomatiqueTransfertCharge = ""
    NumPontReelDuTransfert = 0
    
    
        
    If NumPosteDepart >= POSTES.P_C13 And NumPosteDepart <= POSTES.P_C16 Then
        'Call Log("transfert de l'ano vers la prochaine cuve d'id poste" & NumPosteArrivee & "DelaiSupStabilisationChargeSecondes = " & DelaiSupStabilisationChargeSecondes)
    End If
    
    If NumPosteDepart >= POSTES.P_CHGT_1 And NumPosteDepart <= DERNIER_POSTE And _
       NumPosteArrivee >= POSTES.P_CHGT_1 And NumPosteArrivee <= DERNIER_POSTE Then
    
        '--- effacement du tableau ---
        Erase TUnCyclePont()
            
        '--- extraction de la prémisse ---
        ReponseExtraitPremisseDecodee = ExtraitPremisseDecodee(NumPosteDepart, _
                                                                                                            NumPosteArrivee, _
                                                                                                            NumPont, _
                                                                                                            NumPontIA, _
                                                                                                            TempsCycleSecondes, _
                                                                                                            TUnCyclePont())
        
        '*****************************************************************************************************************
        '                            détermination du numéro de pont rellement choisi pour le transfert
        '*****************************************************************************************************************
        If NumPontImpose = PONTS.P_1 Or NumPontImpose = PONTS.P_2 Then
            NumPontReelDuTransfert = NumPontImpose
        Else
            NumPontReelDuTransfert = NumPontIA
        End If
        
        '--- vérification si le système cyclique ou IA dispose du contrôle du pont ---
        ' ATTENTION si l'opérateur dispose du contrôle des 2 ponts on utilise la fonction pour l'interprétation
        ' des commandes en passant par la fonction GestionCommandesOperateur
        If TEtatsPonts(NumPontReelDuTransfert).ControleParOperateur = False Or _
            (TEtatsPonts(PONTS.P_1).ControleParOperateur = True And _
            TEtatsPonts(PONTS.P_2).ControleParOperateur = True) And _
            TEtatsPonts(NumPontReelDuTransfert).PtrEtActionEnCoursAPI.PtrAction = 0 Then
            
            '--- affectation ---
            Reponse = NOUVELLE_LIGNE & "TRANSFERT DE LA CHARGE DE " & _
                              TEtatsPostes(NumPosteDepart).DefinitionPoste.NomPoste & _
                              " EN " & _
                              TEtatsPostes(NumPosteArrivee).DefinitionPoste.NomPoste & _
                              " AVEC LE PONT " & NumPontReelDuTransfert & _
                              vbCrLf & TABULATION_REPONSES
        
            '--- vérification de l'existence de la règle ---
            If ReponseExtraitPremisseDecodee = OK Then

                '--- insertion du temps d'égouttage dans le cycle du pont ---
                If TempsEgouttageSecondes > 0 Then
                    Bidon = InsertionTempsEgouttageDansCyclePont(TempsEgouttageSecondes, TUnCyclePont())
                End If
                
                '--- insertion du délai de stabilisation supplémentaire de la charge dans le cycle du pont ---
                If DelaiSupStabilisationChargeSecondes > 0 Then
                    Bidon = InsertionDelaiSupStabilisationChargeDansCyclePont(DelaiSupStabilisationChargeSecondes, TUnCyclePont())
                End If
                
                '--- lancement du transfert ---
                ReponseEnvoiCyclePont = EnvoiCyclePont(NumPontReelDuTransfert, TUnCyclePont())
                Select Case ReponseEnvoiCyclePont
                    
                    Case OK
                         '--- le cycle a été transféré avec succès, il faut remplir la fiche des paramètres ---
                         With TEtatsPonts(NumPontReelDuTransfert).TParametresCyclesPonts(CYCLES.C_ACTUEL)
                            .NumPosteDepart = NumPosteDepart
                            .NumPosteArrivee = NumPosteArrivee
                            .TypeCycle = TYPES_CYCLES.TC_TRANSFERT_CHARGE
                            .DelaiSupStabilisationChargeSecondes = DelaiSupStabilisationChargeSecondes
                            .TempsEgouttageSecondes = TempsEgouttageSecondes
                         End With
                        
                        '--- toujours restitué la valeur du pont IA avec le n° de pont cyclique ---
                        With TPremisses(NumPosteDepart, NumPosteArrivee)
                            .NumPontIA = .NumPont
                        End With
                        
                        '--- affectation de la réponse ---
                        CouleurReponse = COULEURS.BLEU_3
                        Reponse = OK
                    
                    Case Else
                        '--- le transfert a été refusé / affectation de la réponse ---
                        CouleurReponse = COULEURS.ROUGE_3
                        Reponse = Reponse & "Transfert de la charge de " & TEtatsPostes(NumPosteDepart).DefinitionPoste.NomPoste & _
                                          " en " & TEtatsPostes(NumPosteArrivee).DefinitionPoste.NomPoste & _
                                          " avec le pont " & NumPontReelDuTransfert & _
                                          IIf(TempsEgouttageSecondes = 0, "", ", égouttage " & TempsEgouttageSecondes & " secondes") & _
                                          " REFUSE"
                        Reponse = Reponse & vbCrLf & TABULATION_REPONSES & ReponseEnvoiCyclePont
                
                End Select

            Else

                '--- mauvaise formulation / affectation de la réponse ---
                CouleurReponse = COULEURS.ROUGE_3
                Reponse = ReponseExtraitPremisseDecodee

            End If

        Else

            '--- pas de disposition du pont / affectation de la réponse ---
            CouleurReponse = COULEURS.ROUGE_3
            Reponse = PAS_DE_DISPOSITION_DU_PONT_IA & " " & NumPontReelDuTransfert

        End If

    Else
        
        '--- pas de disposition du pont / affectation de la réponse ---
        CouleurReponse = COULEURS.ROUGE_3
        Reponse = MAUVAISE_FORMULATION
    
    End If
    
   

    '--- valeur de retour ---
    AutomatiqueTransfertCharge = Reponse

End Function

'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Rôle      : Contrôle si une charge est dans un des bains prioritaire avant de lancer un transfert de charge
' Entrées :              NumPontImpose -> Numéro du pont imposé
'                                                             0 = pas de pont imposé, prendre le numéro du pont IA pour effectuer
'                                                            le transfert
'                                                            <> 0, pont imposé, prendre ce numéro de pont pour effectuer le transfert
'
'                             NumPosteDepart -> Numéro du poste de départ
'                            NumPosteArrivee -> Numéro du poste d'arrivée
' Retours :             CouleurReponse -> Couleur de la réponse
'                 ControleBainsPrioritaire -> Message à retourner comme réponse
' Détails  :
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function ControleBainsPrioritaires(ByVal NumPontImpose As Integer, _
                                                                   ByVal NumPosteDepart As Integer, _
                                                                   ByVal NumPosteArrivee As Integer, _
                                                                   ByRef CouleurReponse As Long) As String
    
    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next
    
    '--- constantes privées ---
    Const TOLERANCE_DEPLACEMENT_PONT As Integer = 80               'tolérance de déplacement d'un pont en secondes avant de confirmer la priorité d'un bain
    Const TEMPS_ANNULATION_PRIORITE As Integer = -15                    'temps dépassé dans un bain prioritaire ou l'on annule la priorité
    
    '--- déclaration ---
    Dim ChargeDansBainPrioritaireP1 As Boolean                                   'indique qu'il y a au moins une charge dans la zone du pont 1
                                                                                                                  'qui est dans un bain prioritaire
    Dim ChargeDansBainPrioritaireP2 As Boolean                                   'indique qu'il y a au moins une charge dans la zone du pont 2
                                                                                                                  'qui est dans un bain prioritaire
    Dim ChargePresenteAuPoste As Boolean                                            'indique une charge présente dans un poste
    
    Dim a As Integer                                                                                   'pour les boucles FOR...NEXT
    
    Dim NumPont As Integer                                                                      'numéro du pont donnée par les diagrammes
                                                                                                                  'en cyclique (extrait de la prémisse)
    Dim NumPontIA As Integer                                                                   'numéro de pont donné par le moteur
                                                                                                                  'd'inférence (extrait de la prémisse)
    Dim NumPontReelDuTransfert As Integer                                             'numéro du pont réel qui va effectuer le transfert

    Dim NumPosteBainPrioritaireP1 As Integer                                        'numéro du poste ou se trouve le bain prioritaire pour le pont 1
    Dim NumPosteBainPrioritaireP2 As Integer                                        'numéro du poste ou se trouve le bain prioritaire pour le pont 2
    
    Dim TUnCyclePont(1 To NBR_LIGNES_CYCLES_PONTS) As Integer 'cycle d'un pont avec tous les temps à
                                                                                                                  'envoyer à l'automate
    Dim TempsCycleSecondes As Long                                                    'temps d'un cycle en secondes
    
    Dim DecompteTempsPostePrioritaireP1 As Long                              'décompte du temps en secondes du poste prioritaire pour le pont 1
    Dim DecompteTempsPostePrioritaireP2 As Long                              'décompte du temps en secondes du poste prioritaire pour le pont 2
    
    Dim Reponse As String                                                                        'correspond à la variable de retour de la
                                                                                                                  'fonction
    Dim ReponseExtraitPremisseDecodee As String                                'correspond à la réponse donnée à l'extraction
                                                                                                                  'd'une prémisse décodée
    Dim ReponseEnvoiCyclePont As String                                               'correspond à la réponse donnée à l'envoi
                                                                                                                  'd'un cycle d'un pont
            
    '----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    
    '--- affectation par défaut ---
    ControleBainsPrioritaires = ""
    ChargeDansBainPrioritaireP1 = False
    ChargeDansBainPrioritaireP2 = False
    NumPontReelDuTransfert = 0
    
    '----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    
    '--- recherche si une charge est au moins dans un bain prioritaire pour le PONT 1 ---
    For a = PREMIER_BAIN To POSTES.P_C12
    
        With TEtatsPostes(a)
    
            If .DefinitionPoste.RespectTempsObligatoire = True And _
               .NumCharge >= CHARGES.C_NUM_MINI And .NumCharge <= CHARGES.C_NUM_MAXI Then
                
                '--- affectation du numéro de poste ou se trouve le bain prioritaire pour le pont 1 ---
                NumPosteBainPrioritaireP1 = a
                
                '--- affectation indiquant qu 'il y a au moins une charge dans la zone du pont 1 qui est dans un bain prioritaire ---
                ChargeDansBainPrioritaireP1 = True
    
                '--- sortie directe ---
                Exit For
    
            End If
    
        End With
    
    Next a
    
    '----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    
    '--- recherche si une charge est au moins dans un bain prioritaire pour le PONT 2 ---
    For a = POSTES.P_C17 To DERNIER_POSTE
    
        If (a <> POSTES.P_D1 And a <> POSTES.P_D2) Then
            With TEtatsPostes(a)
    
                If .DefinitionPoste.RespectTempsObligatoire = True And _
                   .NumCharge >= CHARGES.C_NUM_MINI And .NumCharge <= CHARGES.C_NUM_MAXI Then
                    
                    '--- affectation du numéro de poste ou se trouve le bain prioritaire pour le pont 1 ---
                    NumPosteBainPrioritaireP2 = a
        
                    '--- affectation indiquant qu 'il y a au moins une charge dans la zone du pont 1 qui est dans un bain prioritaire ---
                    ChargeDansBainPrioritaireP2 = True
        
                    '--- sortie directe ---
                    Exit For
        
                End If
            End With
        End If
        
        
    
        
    
    Next a
    
    '----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    
    '--- cas des condamnation de ponts ---
    If TEtatsPonts(PONTS.P_1).Condamnation = True Then
        If ChargeDansBainPrioritaireP2 = False Then
            ChargeDansBainPrioritaireP2 = ChargeDansBainPrioritaireP1
        End If
        ChargeDansBainPrioritaireP1 = False
    End If
    
    If TEtatsPonts(PONTS.P_2).Condamnation = True Then
        If ChargeDansBainPrioritaireP1 = False Then
            ChargeDansBainPrioritaireP1 = ChargeDansBainPrioritaireP2
        End If
        ChargeDansBainPrioritaireP2 = False
    End If
    
    '----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    
    '--- pas de charge dans un bain prioritaire donc sortie directe ---
    If ChargeDansBainPrioritaireP1 = False And ChargeDansBainPrioritaireP2 = False Then
                        
        '--- affectation de la réponse ---
        CouleurReponse = COULEURS.BLEU_3
        ControleBainsPrioritaires = OK
        
        '--- sortie de la fonction ---
        Exit Function
    
    End If
    
    '----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    
    '--- recherche du pont concerné par le transfert de charge ---
    If NumPosteDepart >= POSTES.P_CHGT_1 And NumPosteDepart <= DERNIER_POSTE And _
       NumPosteArrivee >= POSTES.P_CHGT_1 And NumPosteArrivee <= DERNIER_POSTE Then
    
        '--- effacement du tableau ---
        Erase TUnCyclePont()
            
        '--- extraction de la prémisse ---
        ReponseExtraitPremisseDecodee = ExtraitPremisseDecodee(NumPosteDepart, _
                                                                                                            NumPosteArrivee, _
                                                                                                            NumPont, _
                                                                                                            NumPontIA, _
                                                                                                            TempsCycleSecondes, _
                                                                                                            TUnCyclePont())
        
        '*****************************************************************************************************************
        '                            détermination du numéro de pont réellement choisi pour le transfert
        '*****************************************************************************************************************
        If NumPontImpose = PONTS.P_1 Or NumPontImpose = PONTS.P_2 Then
            NumPontReelDuTransfert = NumPontImpose
        Else
            NumPontReelDuTransfert = NumPontIA
        End If
    
        '*****************************************************************************************************************
        '                                                           traitement du PONT 1
        '*****************************************************************************************************************
        If NumPontReelDuTransfert = PONTS.P_1 Then
        
            '--- pas de charge prioritaire, alors retour de OK ---
            If ChargeDansBainPrioritaireP1 = False Then
                
                '--- affectation de la réponse ---
                CouleurReponse = COULEURS.BLEU_3
                Reponse = OK
            
            Else
            
                '--- transfert autorisé si le poste de départ est prioritaire ---
                If TEtatsPostes(NumPosteDepart).DefinitionPoste.RespectTempsObligatoire = True Then

                    '--- affectation de la réponse ---
                    CouleurReponse = COULEURS.BLEU_3
                    Reponse = OK

                Else

                    '--- affectation du décompte du temps en secondes du poste prioritaire ---
                    DecompteTempsPostePrioritaireP1 = RechercheDecompteTempsAuPoste(NumPosteBainPrioritaireP1, ChargePresenteAuPoste)
                    
                    If DecompteTempsPostePrioritaireP1 > TOLERANCE_DEPLACEMENT_PONT Or DecompteTempsPostePrioritaireP1 < TEMPS_ANNULATION_PRIORITE Then 'tolérance d'un déplacement si le temps est supérieur à X minutes
                                                                                                                                                                                                                                                                                 'DecompteTempsPostePrioritaireP1 < TEMPS_ANNULATION_PRIORITE
                                                                                                                                                                                                                                                                                 'car si une charge prioritaire ne peut avancer (cas du poste d'arrivée occupé)
                                                                                                                                                                                                                                                                                  'il faut forcer l'avance des autres charges le libérer le poste d'arrivée
                        '--- affectation de la réponse ---
                        CouleurReponse = COULEURS.BLEU_3
                        Reponse = OK
                    
                    Else
                    
                        '--- affectation ---
                        CouleurReponse = COULEURS.ROUGE_3
                        Reponse = NOUVELLE_LIGNE & "BAIN PRIORITAIRE - TRANSFERT REFUSE DE " & _
                                          TEtatsPostes(NumPosteDepart).DefinitionPoste.NomPoste & _
                                          " EN " & _
                                         TEtatsPostes(NumPosteArrivee).DefinitionPoste.NomPoste & _
                                         " AVEC LE PONT " & NumPontReelDuTransfert & _
                                         vbCrLf & TABULATION_REPONSES
                
                    End If
                
                End If

            End If
    
        End If
    
        '*****************************************************************************************************************
        '                                                           traitement du PONT 2
        '*****************************************************************************************************************
        If NumPontReelDuTransfert = PONTS.P_2 Then
        
            '--- pas de charge prioritaire, alors retour de OK ---
            If ChargeDansBainPrioritaireP2 = False Then
                
                '--- affectation de la réponse ---
                CouleurReponse = COULEURS.BLEU_3
                Reponse = OK
            
            Else
            
                '--- transfert autorisé si le poste de départ est prioritaire ---
                If TEtatsPostes(NumPosteDepart).DefinitionPoste.RespectTempsObligatoire = True Then

                    '--- affectation de la réponse ---
                    CouleurReponse = COULEURS.BLEU_3
                    Reponse = OK

                Else

                    '--- affectation du décompte du temps en secondes du poste prioritaire ---
                    DecompteTempsPostePrioritaireP2 = RechercheDecompteTempsAuPoste(NumPosteBainPrioritaireP2, ChargePresenteAuPoste)
                    
                    If DecompteTempsPostePrioritaireP2 > TOLERANCE_DEPLACEMENT_PONT Or DecompteTempsPostePrioritaireP2 < TEMPS_ANNULATION_PRIORITE Then 'tolérance d'un déplacement si le temps est supérieur à X minutes
                                                                                                                                                                                                                                                                                 'DecompteTempsPostePrioritaireP2 < TEMPS_ANNULATION_PRIORITE
                                                                                                                                                                                                                                                                                 'car si une charge prioritaire ne peut avancer (cas du poste d'arrivée occupé)
                                                                                                                                                                                                                                                                                 'il faut forcer l'avance des autres charges le libérer le poste d'arrivée
                    
                        '--- affectation de la réponse ---
                        CouleurReponse = COULEURS.BLEU_3
                        Reponse = OK
                    
                    Else
                    
                        '--- affectation ---
                        CouleurReponse = COULEURS.ROUGE_3
                        Reponse = NOUVELLE_LIGNE & "BAIN PRIORITAIRE - TRANSFERT REFUSE DE " & _
                                          TEtatsPostes(NumPosteDepart).DefinitionPoste.NomPoste & _
                                          " EN " & _
                                         TEtatsPostes(NumPosteArrivee).DefinitionPoste.NomPoste & _
                                         " AVEC LE PONT " & NumPontReelDuTransfert & _
                                         vbCrLf & TABULATION_REPONSES
                    
                    End If
                                
                End If

            End If
    
        End If
    
    End If
    
    '--- valeur de retour ---
    ControleBainsPrioritaires = Reponse

End Function



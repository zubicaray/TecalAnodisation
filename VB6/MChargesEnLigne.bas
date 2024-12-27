Attribute VB_Name = "MChargesEnLigne"
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' R�le                    : MODULE AIDANT A LA GESTION DES CHARGES EN LIGNE
' Nom                    : MChargesEnLigne.bas
' Date de cr�ation : 08/03/2011
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

'--- d�clarations obligatoires ---
Option Explicit

'--- options g�n�rales ---
Option Base 1
DefVar A-Z

'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' R�le      : Vide une charge dans le tableau des �tats des charges
' Entr�es : NumCharge -> Num�ro de la charge � initialiser
' Retours :
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Sub InitialisationCharge(ByVal NumCharge As Integer)
    
    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next
    
    '--- d�claration ---
    Dim a As Integer
    Dim FicheVideGammesAnodisation As EnrGammesAnodisation, _
            FicheVideDetailsCharges As DetailsCharges, _
            FicheVideDetailsGammesAnodisation As EnrDetailsGammesAnodisation, _
            FicheVideDetailsFichesProduction As DetailsFichesProduction
    
    '--- analyse en fonction du PC ---
    'If TypePC <> TYPES_PC. Then Exit Sub

    '--- contr�le avant d'initialiser la charge ---
    If NumCharge >= CHARGES.C_NUM_MINI And NumCharge <= CHARGES.C_NUM_MAXI Then

        With TEtatsCharges(NumCharge)
            
            '--- date d'entr�e en ligne ---
            .DateEntreeEnLigne = Empty
            
            '--- date d'arriv�e au d�chargement ---
            .DateArriveeAuDechargement = Empty
            
            '--- num�ro de barre ---
            .NumBarre = 0
            
            '--- charge prioritaire ---
            .ChargePrioritaire = False                                'indique qu'il sagit  d'une charge prioritaire
                                                                                      'cette option est valid� au chargement
            
            '--- d�lai suppl�mentaire de stabilisation de la charge ---
            .DelaiSupStabilisationChargeSecondes = 0
            
            '--- options 1 et 2 de la charge (vitesse de mont�e-descente, etc ...) ---
            .Options1 = 0
            .Options2 = 0
            
            '--- d�tails des charges ---
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
            
            '--- nombre de postes trait�s (pour la fiche de production) ---
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
' R�le      : Transmet la valeur du mot des options pour une charge
' Entr�es :                             NumCharge -> N� de la charge ou l'on veut envoyer les options
'                                                  Options1 -> Valeur de toutes les options 1
'                                                  Options2 -> Valeur de toutes les options 2
' Retours : EnvoiOptionsPourUneCharge -> OK = Transmission correcte
'                                                                       "" = Incident de transmission
' D�tails  :
'                           Poids FORT du mot transmis OPTIONS 1
'                           ---------------------------------------------------------------------------------------
'                           |  Bit 7 |  Bit 6 | Bit 5 | Bit 4 | Bit 3 | Bit 2 | Bit 1 | Bit 0 |
'                           ---------------------------------------------------------------------------------------
'                           |  128   |   64   |   32  |   16   |    8   |    4    |    2   |     1   |
'                           ---------------------------------------------------------------------------------------
'                                 |           |          |         |         |          |          |         |_____  forcer la mont�e en tr�s petite vitesse
'                                 |           |          |         |         |          |          |__________  forcer la mont�e en petite vitesse
'                                 |           |          |         |         |          |________________ forcer la descente en tr�s petite vitesse
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
'                                 |           |          |         |         |          |          |         |_____  gestion de l'�lectro-vanne du brillantage avec les gammes
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
    
    '--- constantes priv�es ---
    Const NOM_GROUPE As String = "SUIVI_LIGNE"
    Const OPTIONS_GAMME_1_POSTE As String = "OptionsGamme1Poste"     'variable options de la gamme partie 1 pour un poste
    Const OPTIONS_GAMME_2_POSTE As String = "OptionsGamme2Poste"     'variable options de la gamme partie 1 pour un poste
    Const OPTIONS_GAMME_1_PONT As String = "OptionsGamme1P"              'variable options de la gamme partie 1 pour un pont
    Const OPTIONS_GAMME_2_PONT As String = "OptionsGamme2P"              'variable options de la gamme partie 1 pour un pont
    
    '--- d�claration ---
    Dim NumPoste As Integer, _
           NumPont As Integer
    Dim ValeurRetourneeAPI As Long                  'valeur retourn�e par une fonction concernant le dialogue avec l'automate
    Dim NomVariableOptions1 As String              'nom de la variable pour les options 1
    Dim NomVariableOptions2 As String              'nom de la variable pour les options 2
    
    '--- affectation ---
    EnvoiOptionsPourUneCharge = ""

    If NumCharge >= CHARGES.C_NUM_MINI And NumCharge <= CHARGES.C_NUM_MAXI Then
    
        '--- recherche du poste ou se trouve la charge ---
        NumPoste = RechercheNumPostePourUneCharge(NumCharge)
    
        If NumPoste <> 0 Then
        
            If NumPoste < 0 Then
    
                '--- le n� de poste est n�gatif alors la charge est sur un des ponts ---
                NumPont = Abs(NumPoste)
                
                '--- calcul de l'adresse des options pour UN PONT ---
                NomVariableOptions1 = OPTIONS_GAMME_1_PONT & NumPont
                NomVariableOptions2 = OPTIONS_GAMME_2_PONT & NumPont
                
            ElseIf NumPoste >= POSTES.P_CHGT_1 And NumPoste <= DERNIER_POSTE Then
                        
                '--- calcul de l'adresse des options pour UN POSTE ---
                NomVariableOptions1 = OPTIONS_GAMME_1_POSTE & Right("00" & NumPoste, 2)
                NomVariableOptions2 = OPTIONS_GAMME_2_POSTE & Right("00" & NumPoste, 2)
            
            End If
    
            '--- �criture des options ---
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
' R�le      : Recherche le num�ro de poste ou se trouve une charge
' Entr�es :                                          NumCharge -> N� de la charge ou l'on recherche le poste actuel
' Retours : RechercheNumPostePourUneCharge -> 0 si pas de charge dans la ligne
'                                                                                 moins x, une valeur n�gative repr�sente le num�ro du pont si
'                                                                                 la charge se trouve sur un des ponts
'                                                                                 plus x,  une valeur positive repr�sente le num�ro du poste ou
'                                                                                 se trouve la charge
' D�tails  :
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function RechercheNumPostePourUneCharge(ByVal NumCharge As Integer) As Integer

    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next

    '--- d�claration ---
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
' R�le      : Recherche le temps pass� au poste pour une charge
' Entr�es :                                                 NumCharge -> N� de la charge
'                                                                   NumPoste -> N� du poste ou l'on recherche le temps pass�
' Retours : RechercheTempsAuPostePourUneCharge -> 0 si pas de temps au poste ou pas de passage dans ce
'                                                                                         poste
'                                                                                         Sinon le temps pass� au poste en secondes
'                                            DateEntreeDansLePoste -> Date compl�te d'entr�e dans le poste
' D�tails  :
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function RechercheTempsAuPostePourUneCharge(ByVal NumCharge As Integer, _
                                                                                              ByVal NumPoste As Integer, _
                                                                                              ByRef DateEntreeDansLePoste As Date) As Long

    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next

    '--- d�claration ---
    Dim a As Integer
    
    '--- affectation ---
    RechercheTempsAuPostePourUneCharge = 0
                                                                      
    If NumCharge >= CHARGES.C_NUM_MINI And NumCharge <= CHARGES.C_NUM_MAXI Then
                                                                      
        If NumPoste >= POSTES.P_CHGT_1 And NumPoste <= DERNIER_POSTE Then
        
            With TEtatsCharges(NumCharge)
        
                For a = LBound(.TDetailsFichesProduction()) To UBound(.TDetailsFichesProduction())
            
                    With .TDetailsFichesProduction(a)
        
                        '--- recherche du temps si le poste a �t� trouv� ---
                        If .NumPoste = NumPoste Then
            
                            '--- temps r�el au poste ---
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
' R�le      : Recherche le d�compte du temps pass� pour une charge dans un poste
' Entr�es :                                          NumPoste -> N� du poste ou l'on recherche le d�compte du temps pass�
' Retours : RechercheDecompteTempsAuPoste -> d�compte du temps pass� au poste en secondes
'                                 ChargePresenteAuPostee -> TRUE = il y a une charge dans ce poste
'                                                                               FALSE = pas de charge pr�sente dans ce poste
' D�tails  :
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function RechercheDecompteTempsAuPoste(ByVal NumPoste As Integer, _
                                                                                     ByRef ChargePresenteAuPoste As Boolean) As Long

    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next

    '--- d�claration ---
    Dim NumCharge As Integer                                                       'num�ro de charge

    '--- affectation par d�faut ---
    RechercheDecompteTempsAuPoste = 0
    ChargePresenteAuPoste = False
    
    If NumPoste >= POSTES.P_CHGT_1 And NumPoste <= DERNIER_POSTE Then
    
        If TEtatsPostes(NumPoste).DefinitionPoste.AvecTemps = True Then
        
            '--- affectation du num�ro de charge ---
            NumCharge = TEtatsPostes(NumPoste).NumCharge
    
            If NumCharge >= CHARGES.C_NUM_MINI And NumCharge <= CHARGES.C_NUM_MAXI Then
                                                                      
                '--- affectation du d�compte du temps ---
                With TEtatsCharges(NumCharge)
                    
                    If .PtrZoneGammeAnodisation > 0 And .NbrPostesTraites > 0 Then
                        
                        If .TGammesAnodisation.TDetailsGammesAnodisation(.PtrZoneGammeAnodisation).NumPosteReel = NumPoste Then
                                        
                            '--- affectation du d�compte du temps au poste ---
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
' R�le      : Recherche le temps pass� au poste d'anodisation
' Entr�es :                                      NumCharge -> N� de la charge
'                                              NumPosteAnodisation -> N� du poste d'anodisation
' Retours : RechercheTempsAuPosteDeAnodisation -> 0 si pas de temps au poste ou pas de passage dans au poste
'                                                                             d'anodisation
'                                                                             Sinon le temps pass� au poste en secondes
'                               DateEntreeAuPosteAnodisation -> Date compl�te d'entr�e dans le poste d'anodisation
' D�tails  :
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function RechercheTempsAuPosteAnodisation(ByVal NumCharge As Integer, _
                                                                                       ByRef NumPosteAnodisation As Integer, _
                                                                                       ByRef DateEntreeAuPosteAnodisation As Date) As Long

    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next

    '--- d�claration ---
    Dim a As Integer
    
    '--- affectation ---
    RechercheTempsAuPosteAnodisation = 0
                                                                      
    If NumCharge >= CHARGES.C_NUM_MINI And NumCharge <= CHARGES.C_NUM_MAXI Then
        
        With TEtatsCharges(NumCharge)
    
            For a = LBound(.TDetailsFichesProduction()) To UBound(.TDetailsFichesProduction())
        
                With .TDetailsFichesProduction(a)
    
                    '--- recherche du temps si le poste a �t� trouv� ---
                    If .NumPoste = POSTES.P_C13 Or .NumPoste = POSTES.P_C14 Or .NumPoste = POSTES.P_C15 Or .NumPoste = POSTES.P_C16 Then
        
                        '--- temps r�el au poste ---
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
' R�le      : Recherche le num�ro de charge le PLUS PETIT dans la ligne
' Entr�es :
' Retours : RechercheNumeroChargeLePlusPetit -> 0 si pas de charge dans la ligne sinon le num�ro le plus petit
' D�tails  :
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function RechercheNumeroChargeLePlusPetit() As Integer
    
    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next
    
    '--- d�claration ---
    Dim a As Integer, _
           NumCharge As Integer
                                                                      
    '--- affectation ---
    RechercheNumeroChargeLePlusPetit = 0
    NumCharge = CHARGES.C_NUM_MAXI + 1 'forcer � la valeur la plus �lev�e par rapport au n� de charge maxi
    
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
' R�le      : V�rifie l'existence d'un num�ro de charge dans la ligne
' Entr�es :                      NumCharge -> Num�ro de charge faisant l'objet du contr�le
' Retours : ExistenceNumeroCharge -> FALSE = La charge n'existe pas dans la ligne
'                                                               TRUE = La charge existe d�j� dans la ligne
' D�tails  :
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function ExistenceNumeroCharge(ByVal NumCharge As Integer) As Boolean

    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next
    
    '--- d�claration ---
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
' R�le      : V�rifie l'existence d'au moins une charge dans la ligne sans compter le d�chargement
' Entr�es :
' Retours : ExistenceNumeroChargeHorsDechargement -> FALSE = Aucune charge dans la ligne
'                                                                                               TRUE = Une charge au moins existe dans la ligne
' D�tails  :
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function ExistenceChargeEnLigneHorsDechargement() As Boolean

    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next
    
    '--- d�claration ---
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
' R�le      : V�rifie l'existence d'au moins une charge dans la ligne sans compter le chargement et le
'                 d�chargement
' Entr�es :
' Retours : ExistenceNumeroChargeHorsChargementDechargement -> FALSE = Aucune charge dans la ligne
'                                                                                                                 TRUE = Une charge au moins existe dans
'                                                                                                                               la ligne
' D�tails  :
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function ExistenceChargeEnLigneHorsChargementDechargement() As Boolean

    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next
    
    '--- d�claration ---
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
' R�le      : Enregistre le num�ro du poste r�el (cas des postes multiples notamment) dans la gamme d'anodisation
'                 d'une charge se trouvant � un poste
' Entr�es : NumPoste -> Num�ro du poste de contr�le
' Retours :
' D�tails  :
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Sub EnregistreNumPosteReelGamme(ByVal NumPoste As Integer)
    
    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next
    
    '--- d�claration ---
    Dim NumCharge As Integer, _
           NumZone As Integer, _
           NumPremierPosteZone As Integer, _
           NumDernierPosteZone As Integer

    '--- enregistrement ---
    If NumPoste >= POSTES.P_CHGT_1 And NumPoste <= DERNIER_POSTE Then
    
        '--- affectation du num�ro de charge ---
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
    
                        '--- affectation du num�ro ---
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
' R�le      : Incr�mente le pointeur de la zone d'anodisation une fois la charge arriv�e dans le poste
' Entr�es : NumPoste -> Num�ro du poste de contr�le
' Retours :
' D�tails  :
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Sub IncrementationPtrZoneGammeAnodisation(ByVal NumPoste As Integer)
    
    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next
    
    '--- d�claration ---
    Dim NumCharge As Integer, _
           NumProchaineZone As Integer, _
           NumPremierPosteProchaineZone As Integer, _
           NumDernierPosteProchaineZone As Integer

    '--- enregistrement ---
    If NumPoste >= POSTES.P_CHGT_1 And NumPoste <= DERNIER_POSTE Then
    
        '--- affectation du num�ro de charge ---
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
    
                        '--- affectation du num�ro ---
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
' R�le      : Recherche le num�ro de charge le PLUS GRAND dans la ligne
' Entr�es :
' Retours : RechercheNumeroChargeLePlusGrand -> 0 si pas de charge dans la ligne sinon le num�ro le plus
'                 grand
' D�tails  :
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function RechercheNumeroChargeLePlusGrand() As Integer
    
    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next
    
    '--- d�claration ---
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
' R�le      : Envoi dans l'automate d'un num�ro de charge � un POSTE
' Entr�es :    NumPoste -> Num�ro du poste concern�
'                 NumCharge -> Num�ro de charge
' Retours :
' D�tails  :
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function EnvoiNumeroChargePoste(ByVal NumPoste As Integer, _
                                                                      ByVal NumCharge As Integer) As String
    
    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next
    
    '--- constantes priv�es ---
    Const NOM_GROUPE As String = "SUIVI_LIGNE"
    
    '--- d�claration ---
    Dim ValeurRetourneeAPI As Long                  'valeur retourn�e par une fonction concernant le dialogue avec l'automate
    Dim NomVariable As String                            'nom de la variable
    
    '--- affectation ---
    EnvoiNumeroChargePoste = ""

    If NumPoste >= POSTES.P_CHGT_1 And NumPoste <= DERNIER_POSTE Then
    
        '--- affectation du nom de la variable ---
        NomVariable = "NumChargePoste" & Right("00" & NumPoste, 2)
                
        '--- �criture ---
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
' R�le      : Envoi dans l'automate d'un num�ro de charge � un POSTE avec les options
' Entr�es :                                               NumPoste -> Num�ro du poste concern�
'                                                            NumCharge -> Num�ro de charge
'                                                                Options1 -> Valeur de toutes les options 1
'                                                                Options2 -> Valeur de toutes les options 2
' Retours : EnvoiNumeroChargePosteAvecOptions -> OK = Transmission correcte
'                                                                                      "" = Incident de transmission
' D�tails  :
'                           Poids FORT du mot transmis
'                           ---------------------------------------------------------------------------------------
'                           |  Bit 7 |  Bit 6 | Bit 5 | Bit 4 | Bit 3 | Bit 2 | Bit 1 | Bit 0 |
'                           ---------------------------------------------------------------------------------------
'                           |  128   |   64   |   32  |   16   |    8   |    4    |    2   |     1   |
'                           ---------------------------------------------------------------------------------------
'                           |           |          |         |         |          |          |         |_____  forcer la mont�e en tr�s petite vitesse
'                           |           |          |         |         |          |          |__________  forcer la mont�e en petite vitesse
'                           |           |          |         |         |          |________________ forcer la descente en tr�s petite vitesse
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
    
    '--- constantes priv�es ---
    Const NOM_GROUPE As String = "SUIVI_LIGNE"
    Const OPTIONS_GAMME_1_POSTE As String = "OptionsGamme1Poste"     'variable options de la gamme partie 1 pour un poste
    Const OPTIONS_GAMME_2_POSTE As String = "OptionsGamme2Poste"     'variable options de la gamme partie 1 pour un poste
    
    '--- d�claration ---
    Dim ValeurRetourneeAPI As Long                                 'valeur retourn�e par une fonction concernant le dialogue avec l'automate
    Dim NomVariableNumChargePoste As String               'nom de la variable pour le num�ro de charge au poste
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
            
                '--- �criture du num�ro de charge ---
                ValeurRetourneeAPI = APIEcritureVariableNommee(NOM_GROUPE, NomVariableNumChargePoste, NumCharge)
                If ValeurRetourneeAPI = 0 Then
                
                    '--- �criture des options 1 ---
                    ValeurRetourneeAPI = APIEcritureVariableNommee(NOM_GROUPE, NomVariableOption1, Options1)
                    
                    If ValeurRetourneeAPI = 0 Then
                    
                        '--- �criture des options 2 ---
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
' R�le      : Envoi dans l'automate d'un num�ro de charge � un PONT
' Entr�es :      NumPont -> Num�ro du pont concern�
'                 NumCharge -> Num�ro de charge
' Retours :
' D�tails  :
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function EnvoiNumeroChargePont(ByVal NumPont As Integer, _
                                                                    ByVal NumCharge As Integer) As String
    
    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next
    
    '--- constantes priv�es ---
    Const NOM_GROUPE As String = "SUIVI_LIGNE"
    
    '--- d�claration ---
    Dim ValeurRetourneeAPI As Long                  'valeur retourn�e par une fonction concernant le dialogue avec l'automate
    Dim NomVariable As String                            'nom de la variable
    

    
    '--- affectation ---
    EnvoiNumeroChargePont = ""
    
    If NumPont = PONTS.P_1 Or NumPont = PONTS.P_2 Then
                    
        If NumCharge = 0 Or (NumCharge >= CHARGES.C_NUM_MINI And NumCharge <= CHARGES.C_NUM_MAXI) Then

            '--- affectation du nom de la variable ---
            NomVariable = "NumChargeP" & NumPont
                
            '--- �criture ---
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
' R�le      : Recherche le prochain num�ro de charge valide pour une entr�e dans la ligne
' Entr�es :
' Retours :
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function ProchainNumeroCharge() As Integer

    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next

    '--- d�claration ---
    Dim TNumChargesUtilisees(CHARGES.C_NUM_MINI To CHARGES.C_NUM_MAXI) As Boolean
    Dim a As Integer, _
           LeProchainNumeroCharge As Integer

    '--- affectation ---
    LeProchainNumeroCharge = 0
    
    '--- recherche du prochain num�ro de charge pour les ponts ---
    For a = LBound(TEtatsPonts) To UBound(TEtatsPonts())
        With TEtatsPonts(a)
            If .NumCharge >= CHARGES.C_NUM_MINI And .NumCharge <= CHARGES.C_NUM_MAXI Then
                
                '--- indique que le n� de charge est d�j� utilis� ---
                TNumChargesUtilisees(.NumCharge) = True
                
                '--- prendre le n� de charge le plus �lev� ---
                If .NumCharge > LeProchainNumeroCharge Then
                    LeProchainNumeroCharge = .NumCharge
                End If
            
            End If
        End With
    Next a
    
    '--- recherche du prochain num�ro de charge pour les postes ---
    For a = LBound(TEtatsPostes()) To UBound(TEtatsPostes())
        With TEtatsPostes(a)
            If .NumCharge >= CHARGES.C_NUM_MINI And .NumCharge <= CHARGES.C_NUM_MAXI Then
                
                '--- indique que le n� de charge est d�j� utilis� ---
                TNumChargesUtilisees(.NumCharge) = True
                
                '--- prendre le n� de charge le plus �lev� ---
                If .NumCharge > LeProchainNumeroCharge Then
                    LeProchainNumeroCharge = .NumCharge
                End If
            
            End If
        End With
    Next a

    '--- incr�mentation de la variable ---
    Inc LeProchainNumeroCharge

    '--- analyse des limites ---
    If LeProchainNumeroCharge > CHARGES.C_NUM_MAXI Then
        
        '--- rechercher le premier n� de charge non utilis� ---
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
' R�le      : Recherche le prochain num�ro de poste au chargement pour entr�e la charge dans la ligne pour une
'                 gamme de production ou C13 est impos�
' Entr�es :
' Retours :                                                                      ChargePrioritaire -> Indique que la charge est prioritaire
'                 ProchainNumeroPosteChargementSiAnodisationC13Impose ->            0 = Pas de charge en entr�e de ligne
'                                                                                                                       C1 � C6 = Num�ro de poste ou se trouve la
'                                                                                                                                        charge � rentrer
' D�tails  :
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function ProchainNumeroPosteChargementSiAnodisationC13Impose(ByRef ChargePrioritaire As Boolean) As Integer

    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next

    '--- d�claration ---
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
                                                                                                                                   'poste est condamn�
                                If TEtatsCharges(.NumCharge).ChargePrioritaire = True Then
                                    
                                    '--- la charge est prioritaire ---
                                    ChargePrioritaire = True
                                    ProchainNumeroPosteChargementSiAnodisationC13Impose = a
                                    Exit For
                                
                                Else
                                    
                                    '--- la charge n'est pas prioritaire donc contr�ler la date d'entr�e ---
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
' R�le      : Recherche le prochain num�ro de poste au chargement pour entr�e la charge dans la ligne pour une
'                 gamme d'anodisation ou C15 est impos�
' Entr�es :
' Retours :                                                             ChargePrioritaire -> Indique que la charge est prioritaire
'                 ProchainNumeroPosteChargementSiAnodisationC15Impose ->            0 = Pas de charge en entr�e de ligne
'                                                                                                             C1 � C6 = Num�ro de poste ou se trouve la
'                                                                                                                              charge � rentrer
' D�tails  :
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function ProchainNumeroPosteChargementSiAnodisationC16Impose(ByRef ChargePrioritaire As Boolean) As Integer

    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next

    '--- d�claration ---
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
                                                                                                                                   'poste est condamn�
                            
                                If TEtatsCharges(.NumCharge).ChargePrioritaire = True Then
                                    
                                    '--- la charge est prioritaire ---
                                    ChargePrioritaire = True
                                    ProchainNumeroPosteChargementSiAnodisationC16Impose = a
                                    Exit For
                                
                                Else
                                    
                                    '--- la charge n'est pas prioritaire donc contr�ler la date d'entr�e ---
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
' R�le      : Recherche le prochain num�ro de poste au chargement pour entr�e la charge dans la ligne pour une
'                 gamme de production ou C14 est impos�
' Entr�es :
' Retours :                                                                      ChargePrioritaire -> Indique que la charge est prioritaire
'                 ProchainNumeroPosteChargementSiAnodisationC14Impose ->            0 = Pas de charge en entr�e de ligne
'                                                                                                                       C1 � C6 = Num�ro de poste ou se trouve la
'                                                                                                                                        charge � rentrer
' D�tails  :
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function ProchainNumeroPosteChargementSiAnodisationC14Impose(ByRef ChargePrioritaire As Boolean) As Integer

    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next

    '--- d�claration ---
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
                                                                                                                                   'poste est condamn�
                            
                                If TEtatsCharges(.NumCharge).ChargePrioritaire = True Then
                                    
                                    '--- la charge est prioritaire ---
                                    ChargePrioritaire = True
                                    ProchainNumeroPosteChargementSiAnodisationC14Impose = a
                                    Exit For
                                
                                Else
                                    
                                    '--- la charge n'est pas prioritaire donc contr�ler la date d'entr�e ---
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
' R�le      : Recherche le prochain num�ro de poste au chargement pour entr�e la charge dans la ligne pour une
'                 gamme de production ou C15 est impos�
' Entr�es :
' Retours :                                                                      ChargePrioritaire -> Indique que la charge est prioritaire
'                 ProchainNumeroPosteChargementSiAnodisationC15Impose ->            0 = Pas de charge en entr�e de ligne
'                                                                                                                       C1 � C6 = Num�ro de poste ou se trouve la
'                                                                                                                                        charge � rentrer
' D�tails  :
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function ProchainNumeroPosteChargementSiAnodisationC15Impose(ByRef ChargePrioritaire As Boolean) As Integer

    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next

    '--- d�claration ---
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
                                                                                                                                   'poste est condamn�
                            
                                If TEtatsCharges(.NumCharge).ChargePrioritaire = True Then
                                    
                                    '--- la charge est prioritaire ---
                                    ChargePrioritaire = True
                                    ProchainNumeroPosteChargementSiAnodisationC15Impose = a
                                    Exit For
                                
                                Else
                                    
                                    '--- la charge n'est pas prioritaire donc contr�ler la date d'entr�e ---
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
' R�le      : Recherche le prochain num�ro de poste au chargement pour entr�e la charge dans la ligne pour une
'                 gamme d'anodisation ou le chois du poste est automatique
' Entr�es :
' Retours :                                                               ChargePrioritaire -> Indique que la charge est prioritaire
'                 ProchainNumeroPosteChargementSiAnodisationAutomatique ->           0 = Pas de charge en entr�e de ligne
'                                                                                                               C1 � C6 = Num�ro de poste ou se trouve la
'                                                                                                                                charge � rentrer
' D�tails  :
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function ProchainNumeroPosteChargementSiAnodisationAutomatique(ByRef ChargePrioritaire As Boolean) As Integer

    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next

    '--- d�claration ---
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
                                                                                                                                'postes sont condamn�s
                                If TEtatsCharges(.NumCharge).ChargePrioritaire = True Then
                                    
                                    '--- la charge est prioritaire ---
                                    ChargePrioritaire = True
                                    ProchainNumeroPosteChargementSiAnodisationAutomatique = a
                                    Exit For
                                
                                Else
                                    
                                    '--- la charge n'est pas prioritaire donc contr�ler la date d'entr�e ---
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
' R�le      : Recherche le prochain num�ro de poste au d�chargement pour d�poser une charge
' Entr�es :
' Retours : ProchainNumeroPosteDechargement ->           0 = Pas de chariot vide pour d�poser la charge
'                                                                                 D1 � D2 = Num�ro de poste ou la d�pose doit s'effectu�
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function ProchainNumeroPosteDechargement() As Integer

    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next

    '--- d�claration ---
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
' R�le      : Recherche le prochain num�ro de poste au chargement pour d�poser une charge
' Entr�es :
' Retours : ProchainNumeroPosteChargement ->                          0 = Pas de chariot vide pour d�poser la charge
'                                                                              CHGT1 � CHGT4 = Num�ro de poste ou la d�pose doit s'effectu�
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function ProchainNumeroPosteChargement() As Integer

    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next

    '--- d�claration ---
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
' R�le      : Recherche le prochain num�ro de poste d'anodisation disponible
' Entr�es :                                 NumCharge -> num�ro de la charge faisant l'objet de la demande
'                                               TypeDeZone -> FALSE = La zone est une ZONE de DEPART
'                                                                         TRUE = La zone est une ZONE d'ARRIVEE
' Retours : ProchainNumeroPosteAnodisation ->                              0 = Pas de poste d'anodisation disponible
'                                                                             C13 ou C14 ou C15 ou C16 = Num�ro de poste d'anodisation libre
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function ProchainNumeroPosteAnodisation(ByVal NumCharge As Integer, _
                                                                                  ByVal TypeDeZone As Boolean) As Integer

    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next
    
    '--- constantes priv�es ---
    Const ZONE_DEPART As Boolean = False
    Const ZONE_ARRIVEE As Boolean = True

    '--- d�claration ---
    Dim a As Integer
    
    '--- affectation ---
    ProchainNumeroPosteAnodisation = 0
    
    '--- analyse sur les postes d'anodisation ---
    For a = POSTES.P_C13 To POSTES.P_C16
            
        '--- uniquement les postes d'anodisation ---
        If TEtatsPostes(a).Condamnation = False Then
        
            If TypeDeZone = ZONE_DEPART Then
                    
                '--- la zone d'anodisation est une zone de d�part ---
                If NumCharge >= CHARGES.C_NUM_MINI And NumCharge <= CHARGES.C_NUM_MAXI Then
                    
                    '--- v�rifier que le temps au poste est termin� ---
                    With TEtatsCharges(NumCharge)
                        If a = .TGammesAnodisation.TDetailsGammesAnodisation(.PtrZoneGammeAnodisation).NumPosteReel Then
                            If .TGammesAnodisation.TDetailsGammesAnodisation(.PtrZoneGammeAnodisation).FinDuTempsPosteReel = True Then
                                ProchainNumeroPosteAnodisation = a
                            End If
                        End If
                    End With
                
                End If
                
            Else
    
                '--- la zone d'anodisation est une zone d'arriv�e ---
                If NumCharge = 0 Or (NumCharge >= CHARGES.C_NUM_MINI And NumCharge <= CHARGES.C_NUM_MAXI) Then
                
                    With TEtatsCharges(NumCharge)
                            
                        '--- recherche en fonction du choix du poste d'anodisation ---
                        Select Case .TGammesAnodisation.ChoixPosteAnodisation
                    
                            Case CHOIX_POSTE_ANODISATION.C_AUTOMATIQUE
                               '--- prendre le premier poste d'arriv�e vide ---
                                If TEtatsPostes(a).NumCharge = 0 Then
                                    ProchainNumeroPosteAnodisation = a
                                End If
                            
                            Case CHOIX_POSTE_ANODISATION.C_C13_IMPOSE
                                '--- C13 est impos� ---
                                If a = POSTES.P_C13 Then
                                    If TEtatsPostes(a).NumCharge = 0 Then
                                        ProchainNumeroPosteAnodisation = a
                                    End If
                                End If
                            
                            Case CHOIX_POSTE_ANODISATION.C_C14_IMPOSE
                                '--- C14 est impos� ---
                                If a = POSTES.P_C14 Then
                                    If TEtatsPostes(a).NumCharge = 0 Then
                                        ProchainNumeroPosteAnodisation = a
                                    End If
                                End If
                            
                            Case CHOIX_POSTE_ANODISATION.C_C15_IMPOSE
                                '--- C15 est impos� ---
                                If a = POSTES.P_C15 Then
                                    If TEtatsPostes(a).NumCharge = 0 Then
                                        ProchainNumeroPosteAnodisation = a
                                    End If
                                End If
                            
                            Case CHOIX_POSTE_ANODISATION.C_C16_IMPOSE
                                '--- C16 est impos� ---
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
' R�le      : Recherche le prochain num�ro de poste de colmatage chaud
' Entr�es :                                             NumCharge -> num�ro de la charge faisant l'objet de la demande
'                                                           TypeDeZone -> FALSE = La zone est une ZONE de DEPART
'                                                                                     TRUE = La zone est une ZONE d'ARRIVEE
' Retours : ProchainNumeroPosteColmatageChaud ->                0 = Pas de poste de colmatage disponible
'                                                                                    C32 ou C33 = Num�ro de poste de colmatage � chaud
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function ProchainNumeroPosteColmatageChaud(ByVal NumCharge As Integer, _
                                                                                          ByVal TypeDeZone As Boolean) As Integer

    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next
    
    '--- constantes priv�es ---
    Const ZONE_DEPART As Boolean = False
    Const ZONE_ARRIVEE As Boolean = True

    '--- d�claration ---
    Dim a As Integer
    
    '--- affectation ---
    ProchainNumeroPosteColmatageChaud = 0
    
    For a = POSTES.P_C31 To POSTES.P_C32
                
        If TEtatsPostes(a).Condamnation = False Then
        
            If TypeDeZone = ZONE_DEPART Then
                    
                '--- la zone est une zone de d�part ---
                If NumCharge >= CHARGES.C_NUM_MINI And NumCharge <= CHARGES.C_NUM_MAXI Then
                    
                    '--- v�rifier que le temps au poste est termin� ---
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
    
                '--- la zone de colmatage est une zone d'arriv�e ---
                If NumCharge = 0 Or (NumCharge >= CHARGES.C_NUM_MINI And NumCharge <= CHARGES.C_NUM_MAXI) Then
                
                    With TEtatsCharges(NumCharge)
                            
                        '--- prendre le premier poste d'arriv�e vide ---
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
' R�le      : Recherche le prochain num�ro de poste de brillantage
' Entr�es :                                    NumCharge -> num�ro de la charge faisant l'objet de la demande
'                                                 TypeDeZone -> FALSE = La zone est une ZONE de DEPART
'                                                                            TRUE = La zone est une ZONE d'ARRIVEE
' Retours : ProchainNumeroPosteBrillantage ->                 0 = Pas de poste de brillantage disponible
'                                                                            C05 ou C07 = Num�ro de poste du brillantage
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function ProchainNumeroPosteBrillantage(ByVal NumCharge As Integer, _
                                                                                 ByVal TypeDeZone As Boolean) As Integer

    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next
    
    '--- constantes priv�es ---
    Const ZONE_DEPART As Boolean = False
    Const ZONE_ARRIVEE As Boolean = True

    '--- d�claration ---
    Dim a As Integer
    
    '--- affectation ---
    ProchainNumeroPosteBrillantage = 0
    
    For a = POSTES.P_C05 To POSTES.P_C07
                
        Select Case a
        
            Case POSTES.P_C05, POSTES.P_C07
                '--- poste de brillantages uniquement ---
                If TEtatsPostes(a).Condamnation = False Then
                
                    If TypeDeZone = ZONE_DEPART Then
                            
                        '--- la zone est une zone de d�part ---
                        If NumCharge >= CHARGES.C_NUM_MINI And NumCharge <= CHARGES.C_NUM_MAXI Then
                            
                            '--- v�rifier que le temps au poste est termin� ---
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
            
                        '--- la zone de colmatage est une zone d'arriv�e ---
                        If NumCharge = 0 Or (NumCharge >= CHARGES.C_NUM_MINI And NumCharge <= CHARGES.C_NUM_MAXI) Then
                        
                            With TEtatsCharges(NumCharge)
                                    
                                '--- prendre le premier poste d'arriv�e vide ---
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
' R�le      : Recherche le prochain num�ro de poste du s�choir
' Entr�es :                                       NumCharge -> num�ro de la charge faisant l'objet de la demande
'                                                     TypeDeZone -> FALSE = La zone est une ZONE de DEPART
'                                                                                TRUE = La zone est une ZONE d'ARRIVEE
' Retours : ProchainNumeroPosteSechoir ->                  0 = Pas de poste de colmatage disponible
'                                                                       C33 ou C34 = Num�ro de poste de colmatage � chaud
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function ProchainNumeroPosteSechoir(ByVal NumCharge As Integer, _
                                                                           ByVal TypeDeZone As Boolean) As Integer

    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next
    
    '--- constantes priv�es ---
    Const ZONE_DEPART As Boolean = False
    Const ZONE_ARRIVEE As Boolean = True

    '--- d�claration ---
    Dim a As Integer
    
    '--- affectation ---
    ProchainNumeroPosteSechoir = 0
    
    For a = POSTES.P_C33 To POSTES.P_C34
                
        If TEtatsPostes(a).Condamnation = False Then
        
            If TypeDeZone = ZONE_DEPART Then
                    
                '--- la zone est une zone de d�part ---
                If NumCharge >= CHARGES.C_NUM_MINI And NumCharge <= CHARGES.C_NUM_MAXI Then
                    
                    '--- v�rifier que le temps au poste est termin� ---
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
    
                '--- la zone du s�choir est une zone d'arriv�e ---
                If NumCharge = 0 Or (NumCharge >= CHARGES.C_NUM_MINI And NumCharge <= CHARGES.C_NUM_MAXI) Then
                
                    With TEtatsCharges(NumCharge)
                            
                        '--- prendre le premier poste d'arriv�e vide ---
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
' R�le      : D�compte un temps d'une charge dans un poste par rapport au temps de poste dans la gamme
' Entr�es : NumPoste -> N� du poste ou l'on doit d�compter le temps
' Retours :
' D�tails  : La variable FinDuTempsPosteReel de la gamme d'anodisation monte quand le d�compte du temps est
'                 inf�rieur ou �gale � 0
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Sub DecompteDuTempsAuPosteSecondes(ByVal NumPoste As Integer)
    
    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next
    
    '--- constantes priv�es ---
    Const TEMPS_MOUVEMENT_AVANT_PRISE As Long = 5           'temps moyen correspondant � la fermeture des accroches et au d�but de mont�e
    
    '--- d�claration ---
    Dim NumCharge As Integer
    Dim TempsAuPosteSecondes As Long, _
            TempsDepuisEntreeDansLePosteSecondes As Long, _
            DecompteDuTempsSecondes As Long

    If NumPoste >= POSTES.P_CHGT_1 And NumPoste <= DERNIER_POSTE Then
    
        '--- affectation du num�ro de charge ---
        NumCharge = TEtatsPostes(NumPoste).NumCharge
    
        If TEtatsPostes(NumPoste).DefinitionPoste.AvecTemps = True Then
        
            If NumCharge >= CHARGES.C_NUM_MINI And NumCharge <= CHARGES.C_NUM_MAXI Then

                '--- d�compte du temps ---
                With TEtatsCharges(NumCharge)
                    If .PtrZoneGammeAnodisation > 0 And .NbrPostesTraites > 0 Then
                        If .TGammesAnodisation.TDetailsGammesAnodisation(.PtrZoneGammeAnodisation).NumPosteReel = .TDetailsFichesProduction(.NbrPostesTraites).NumPoste Then
                                        
                            '--- recherche du temps th�orique dans la gamme ---
                            TempsAuPosteSecondes = .TGammesAnodisation.TDetailsGammesAnodisation(.PtrZoneGammeAnodisation).TempsAuPosteSecondes
                            
                            '--- calcul du temps � partir de la fiche de production ---
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
' R�le      : D�compte un temps d'aletre dans un poste
' Entr�es : NumPoste -> N� du poste ou l'on doit d�compter le temps
' Retours :
' D�tails  : La variable FinDuTempsPosteReel de la gamme d'anodisation monte quand le d�compte du temps est
'                 inf�rieur ou �gale � 0
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Sub DecompteDuTempsAlerteAuPosteSecondes(ByVal NumPoste As Integer)
    
    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next
    
    '--- constantes priv�es ---
    Const TEMPS_MOUVEMENT_AVANT_PRISE As Long = 5           'temps moyen correspondant � la fermeture des accroches et au d�but de mont�e
    
    '--- d�claration ---
    Dim DebutAlertePosteReel As Boolean                                       'indique le d�but de l'alerte au poste r�el
    Dim UneAlerteEstEnCours As Boolean                                        'indique une alerte en cours
    
    Static MemAntiRebondKlaxon As Boolean                                  'm�moire anti-rebond de lancement du klaxon
    
    Dim NumCharge As Integer                                                          'repr�sente un num�ro de charge
    Dim NumPosteReel As Integer                                                     'repr�sente un num�ro de poste r�el
    
    Dim TempsAlerteAuPosteSecondes As Long, _
            TempsDepuisEntreeDansLePosteSecondes As Long, _
            DecompteDuTempsSecondes As Long
    
    Dim TexteAlerte As String                                                            'texte de l'alerte destin� � l'afficheur

    If NumPoste >= POSTES.P_CHGT_1 And NumPoste <= DERNIER_POSTE Then
    
        '--- affectation ---
        NumCharge = TEtatsPostes(NumPoste).NumCharge
    
        If TEtatsPostes(NumPoste).DefinitionPoste.AvecTemps = True Then
        
            If NumCharge >= CHARGES.C_NUM_MINI And NumCharge <= CHARGES.C_NUM_MAXI Then

                '--- d�compte du temps ---
                With TEtatsCharges(NumCharge)
                    
                    If .PtrZoneGammeAnodisation > 0 And .NbrPostesTraites > 0 Then
                        
                        If .TGammesAnodisation.TDetailsGammesAnodisation(.PtrZoneGammeAnodisation).NumPosteReel = .TDetailsFichesProduction(.NbrPostesTraites).NumPoste Then
                                        
                            '----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
                            
                            '--- recherche du temps th�orique dans la gamme ---
                            TempsAlerteAuPosteSecondes = .TGammesAnodisation.TDetailsGammesAnodisation(.PtrZoneGammeAnodisation).TempsAlerteSecondes
                            
                            '--- affectation du bit de d�but d'alerte au poste ---
                            DebutAlertePosteReel = .TGammesAnodisation.TDetailsGammesAnodisation(.PtrZoneGammeAnodisation).DebutAlertePosteReel
                            
                            '----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
                            
                            '--- gestion de l'alerte ---
                            If DebutAlertePosteReel = True Then
                            
                                '--- affectation du num�ro du poste ---
                                NumPosteReel = .TGammesAnodisation.TDetailsGammesAnodisation(.PtrZoneGammeAnodisation).NumPosteReel
                                
                                '--- affectation du texte de l'alerte ---
                                TexteAlerte = "ALERTE " & TEtatsPostes(NumPosteReel).DefinitionPoste.NomPoste & ": " & CTemps(.TGammesAnodisation.TDetailsGammesAnodisation(.PtrZoneGammeAnodisation).DecompteDuTempsAuPosteReelSecondes)
                                
                                '--- affectation indiquant l'alerte en cours ---
                                UneAlerteEstEnCours = True
                            
                            End If
                            
                            '----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
                            
                            '--- calcul du temps � partir de la fiche de production ---
                            With .TDetailsFichesProduction(.NbrPostesTraites)
                                
                                If .DateEntreePoste <> Empty And .DateSortiePoste <> Empty And TempsAlerteAuPosteSecondes > 0 And DebutAlertePosteReel = False Then
                                    
                                    '--- calcul ---
                                    TempsDepuisEntreeDansLePosteSecondes = DateDiff("s", .DateEntreePoste, .DateSortiePoste)
                                    DecompteDuTempsSecondes = TempsAlerteAuPosteSecondes - TempsDepuisEntreeDansLePosteSecondes
                                    
                                    '--- affectation au bon endroit dans la gamme ---
                                    With TEtatsCharges(NumCharge)
                                        
                                        '--- affectation du temps d'alerte r�el en secondes ---
                                        .TGammesAnodisation.TDetailsGammesAnodisation(.PtrZoneGammeAnodisation).DecompteDuTempsAlerteReelSecondes = CStr(DecompteDuTempsSecondes)
                                        
                                        '--- mont�e du drapeau d'alerte ---
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
                                
            '--- alerte en cours donc affectation de la priorit� d'affichage des alertes ---
            PrioriteAfficheurPourAlertes = True
                                
            '--- affichage du message d'alerte ---
            Bidon = MessageAfficheur("B", TexteAlerte)
                                
            '--- lancement du klaxon ---
            If MemAntiRebondKlaxon = False Then
                
                '--- mont�e du bit du klaxon dans l'automate ---
                Bidon = APIEcritureVariableNommee("DEFAUTS", "M_Dem_PC_Klaxon", True)
        
                '--- affectation de la m�moire anti-rebond du klaxon ---
                MemAntiRebondKlaxon = True
        
            End If
        
        Else
    
            '--- pas d'alerte donc RAZ de la priorit� d'affichage des alertes ---
            PrioriteAfficheurPourAlertes = False
            
            If MemAntiRebondKlaxon = True Then
            
                '--- effacement de l'afficheur ---
                Bidon = MessageAfficheur("B", "")
                    
                '--- RAZ de la m�moire anti-rebond du klaxon ---
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
' R�le      : Recherche le prochain num�ro de poste valide (la ou on peut prendre ou d�pos� une charge)
' Entr�es :                            NumCharge -> num�ro de la charge faisant l'objet de la demande
'                                               NumZone -> N� de la zone (d�part ou arriv�e)
'                                          TypeDeZone -> FALSE = La zone est une ZONE de DEPART
'                                                                     TRUE = La zone est une ZONE d'ARRIVEE
' Retours : ProchainNumeroPosteValide -> 0 = Pas de poste de valide sinon le num�ro du poste
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function ProchainNumeroPosteValide(ByVal NumCharge As Integer, _
                                                                         ByVal NumZone As Integer, _
                                                                         ByVal TypeDeZone As Boolean) As Integer
    
    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next

    '--- constantes priv�es ---
    Const ZONE_DEPART As Boolean = False
    Const ZONE_ARRIVEE As Boolean = True
    
    '--- d�claration ---
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
            
                '--- s�choir (poste 1 et 2) ---
                'ProchainNumeroPosteValide = ProchainNumeroPosteSechoir(NumCharge, False)
            
            
            ' FIN MODIF 20200120 SZP
            ElseIf NumPremierPosteZone = POSTES.P_D1 And NumDernierPosteZone = POSTES.P_D2 Then
                
                '--- zone du d�chargement ---
            
            ElseIf NumPremierPosteZone = NumDernierPosteZone Then
                            
                '--- toutes les autres zones (poste simple) ---
                If TEtatsPostes(NumPremierPosteZone).Condamnation = False Then
                    If NumCharge >= CHARGES.C_NUM_MINI And NumCharge <= CHARGES.C_NUM_MAXI Then
                        
                        '--- v�rifier que le temps au poste est termin� ---
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
            
                '--- zone de chargement (cas du d�chargement � un des postes de chargement) ---
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
            
                '--- s�choir (poste 1 et 2) ---
                'ProchainNumeroPosteValide = ProchainNumeroPosteSechoir(NumCharge, True)
            ' FIN MODIF 20200120 SZP ***********************************************************************
            
            
            ElseIf NumPremierPosteZone = POSTES.P_D1 And NumDernierPosteZone = POSTES.P_D2 Then
                
                '--- zone du d�chargement ---
                ProchainNumeroPosteValide = ProchainNumeroPosteDechargement()
            
            ElseIf NumPremierPosteZone = NumDernierPosteZone Then
            
                '--- toutes les autres zones (poste simple) ---
                With TEtatsPostes(NumPremierPosteZone)
                    If .Condamnation = False Then
                        If .NumCharge = 0 Then                    'v�rifier si le poste d'arriv�e est vide
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
' R�le      : Recherche le prochain num�ro th�orique d'un poste d'arriv�e
' Entr�es :                             NumCharge -> num�ro de la charge faisant l'objet de la demande
'                                    NumZoneArrivee -> N� de la zone (d�part ou arriv�e)
' Retours : ProchainNumeroPosteValide -> 0 = Pas de poste de valide sinon le num�ro du poste
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function ProchainNumTheoriquePosteArrivee(ByVal NumCharge As Integer, _
                                                                                     ByVal NumZoneArrivee As Integer)
    
    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next

    '--- constantes priv�es ---
    
    '--- d�claration ---
    Dim NumPremierPosteZone As Integer, _
           NumDernierPosteZone As Integer
    
    '--- affectation ---
    ProchainNumTheoriquePosteArrivee = 0
        
    If NumZoneArrivee >= LIMITE_BASSE_ZONES And NumZoneArrivee <= LIMITE_HAUTE_ZONES Then
    
        '--- affectation ---
        NumPremierPosteZone = TZones(NumZoneArrivee).NumPremierPoste
        NumDernierPosteZone = TZones(NumZoneArrivee).NumDernierPoste
        
        '--- analyse de la zone d'arriv�e ---
        If NumPremierPosteZone = POSTES.P_CHGT_1 And NumDernierPosteZone = POSTES.P_CHGT_2 Then
        
            '--- zone de chargement (cas du d�chargement � un des postes de chargement) ---
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
        
            '--- s�choir (poste 1 et 2) ---
            'ProchainNumTheoriquePosteArrivee = ProchainNumeroPosteSechoir(NumCharge, True)
        ' FIN MODIF SZV 20200120 -----------------------------------------------------
        ElseIf NumPremierPosteZone = POSTES.P_D1 And NumDernierPosteZone = POSTES.P_D2 Then
            
            '--- zone du d�chargement ---
            ProchainNumTheoriquePosteArrivee = ProchainNumeroPosteDechargement()
        
        ElseIf NumPremierPosteZone = NumDernierPosteZone Then
        
            '--- toutes les autres zones (poste simple) ---
            With TEtatsPostes(NumPremierPosteZone)
                If .Condamnation = False Then
                    '--- ne pas v�rifier si le poste d'arriv�e est vide car le poste d'arriv�e est bien celui la malgr� la pr�sence d'une charge ---
                    ProchainNumTheoriquePosteArrivee = NumPremierPosteZone
                End If
            End With
        
        End If
    
    Else
        
        '--- CAS NORMALLEMENT IMPOSSIBLE ---
            
    End If
        
End Function

'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' R�le      : Transfert un cycle d'un pont vers l'automate
' Entr�es :            NumPont -> Fonction de l'�num�ration PONTS
'                      TCyclePont() -> Tableau contenant le cycle du pont � transf�rer
' Retours : EnvoiCyclePont -> OK = tout va bien
'                                               ERREUR_COMMUNICATION_API = indique une erreur de communication avec l'API
'                                               CYCLE_DEJA_EN_COURS = l'envoi n'est pas possible car un cycle est d�j� en cours
'                                               PONT_NON_AUTOMATIQUE = l'envoi n'est pas possible car le pont n'est pas en auto
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function EnvoiCyclePont(ByVal NumPont As Integer, _
                                                    ByRef TCyclePont() As Integer) As String
    
    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next
    
    '--- constantes priv�es ---
    
    '--- d�claration ---
    Dim a As Integer                                                                                    'pour les boucles FOR...NEXT
    Dim PtrAction As Integer                                                                        'pointeur d'une action
    Dim TValeurs(0 To NBR_LIGNES_CYCLES_PONTS - 1) As Integer      'tableau contenant les valeurs
    
    Dim RefGroupe As Long                                                                         'r�f�rence sur le groupe
    Dim TRefElements(0 To NBR_LIGNES_CYCLES_PONTS - 1) As Long 'r�f�rence sur un �l�ment
    Dim NbrValeurs As Long                                                                        'nombre de valeurs
    Dim ValeurRetourneeAPI As Long                                                          'valeur retourn�e par une fonction
                                                                                                                    'concernant le dialogue avec l'automate
    
    Dim NomGroupe As String                                                                      'repr�sente un nom de groupe
    Dim AdrPtrAction As String                                                                      'adresse du pointeur des actions

    Dim TEtatsCommunication As Variant                                                    'tableau contenant les �tats de la
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
    
    '--- contr�le sur la dimension du tableau ---
    If LBound(TCyclePont()) <> 1 And UBound(TCyclePont()) <> NBR_LIGNES_CYCLES_PONTS Then
        Exit Function
    End If
    
    If PROGRAMME_AVEC_AUTOMATE = True Then
    
        If TEtatsPonts(NumPont).ModePont = MODES_PONTS.M_AUTOMATIQUE Then
    
            '--- lecture du pointeur de l'action ---
            ValeurRetourneeAPI = APILectureMot(AdrPtrAction, PtrAction)
            
            If ValeurRetourneeAPI = 0 Then
    
                If PtrAction = 0 Then
        
                    '--- r�f�rence sur le groupe ---
                    RefGroupe = OccFPrincipale.AOCFPrincipale.GetGroupRef(RefServeur, NomGroupe)
                    If RefGroupe <= 0 Then
                        EnvoiCyclePont = ERREUR_COMMUNICATION_API
                        Exit Function
                    End If
        
                    '--- affectation du tableaux des r�f�rences et des valeurs ---
                    For a = 1 To NBR_LIGNES_CYCLES_PONTS
        
                        '--- affectation de chaque �l�ment ---
                        TRefElements(a - 1) = OccFPrincipale.AOCFPrincipale.GetItemRef(RefGroupe, "Action" & Right("00" & a, 2))
        
                        '--- affectation des valeurs ---
                        TValeurs(a - 1) = TCyclePont(a)
                    
                    Next a
        
                    '--- �criture dans l'automate ---
                    ValeurRetourneeAPI = OccFPrincipale.AOCFPrincipale.Write(NbrValeurs, TRefElements, TValeurs, TEtatsCommunication)
                    
                
                    If ValeurRetourneeAPI = 0 Then
                
                        '--- attente pour �tre s�r que la table soit d�j� dans l'automate ---
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
            
                    '--- cycle d�ja en cours ---
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
' R�le      : Insertion d'un temps d'�gouttage dans un cycle destin� � un pont
' Entr�es :                         TempsEgouttageSecondes -> Repr�sente le temps d'�gouttage en secondes
'                                                               TCyclePont() -> Tableau contenant le cycle du pont � modifier
' Retours : InsertionTempsEgouttageDansCyclePont -> OK = la valeur a �t� ins�r�
'                                                                                        ""   = l'action TEMPO_EGOUT n'a pas �t� trouv�
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function InsertionTempsEgouttageDansCyclePont(ByVal TempsEgouttageSecondes As Integer, _
                                                                                             ByRef TCyclePont() As Integer) As String
    
    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next

    '--- d�claration ---
    Dim a As Integer, _
           NumAction As Integer
    
    '--- affectation ---
    InsertionTempsEgouttageDansCyclePont = ""
    
    '--- recherche de l'action du temps d'�gouttage et enregistrement dans le cycle ---
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
        
            '--- incr�ment si param�tre pour pointer toujours une action ---
            If TActions(NumAction).ParametreOuiNon = True Then
                Inc a
            End If
        
        End If
    
    Next a

End Function

'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' R�le      : Insertion d'un d�lai suppl�mentaire de stabilisation de la charge dans un cycle destin� � un pont
' Entr�es :                        DelaiSupStabilisationChargeSecondes -> Repr�sente le d�lai suppl�mentaire
'                                                                                                           de stabilisation de la charge en secondes
'                                                                                  TCyclePont() -> Tableau contenant le cycle du pont � modifier
' Retours : InsertionDelaiSupStabilisationChargeDansCyclePont -> OK = la valeur a �t� ins�r�
'                                                                                                            ""  = l'action TEMPO_STAB n'a pas �t� trouv�
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function InsertionDelaiSupStabilisationChargeDansCyclePont(ByVal DelaiSupStabilisationChargeSecondes As Integer, _
                                                                                                                ByRef TCyclePont() As Integer) As String
    
    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next

    '--- d�claration ---
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
        
            '--- incr�ment si param�tre pour pointer toujours une action ---
            If TActions(NumAction).ParametreOuiNon = True Then
                Inc a
            End If
        
        End If
    
    Next a

End Function

'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' R�le      : Transmet le d�placement du pont au poste voulu en AUTOMATIQUE
' Entr�es :                                  NumPont -> Num�ro du pont concern�
'                                                NumPoste -> Num�ro du poste souhait�
'                                      CouleurReponse -> Couleur de la r�ponse
' Retours : AutomatiqueDeplacementPont -> Message � retourner comme r�ponse
' D�tails  :
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function AutomatiqueDeplacementPont(ByVal NumPont As Integer, _
                                                                           ByVal NumPoste As Integer, _
                                                                           ByRef CouleurReponse As Long) As String
    
    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next
    
    '--- constantes priv�es ---
    
    Dim Texte As String
    Texte = "AutomatiqueDeplacementPont " & NumPont & ", NumPoste: " & NumPoste
    AfficheRenseignementsDebug CouleurReponse, Texte & vbCrLf
    
    If NumPoste >= POSTES.P_C13 And NumPoste <= POSTES.P_C16 Then
        Call Log(Texte)
    End If
    
    
    '--- d�claration ---
    Dim TUnCyclePont(1 To NBR_LIGNES_CYCLES_PONTS) As Integer
    Dim Reponse As String, _
            ReponseEnvoiCyclePont As String
    
    '--- affectation ---
    AutomatiqueDeplacementPont = ""
    
    If NumPont = PONTS.P_1 Or NumPont = PONTS.P_2 Then
                    
        '--- v�rification si le syst�me cyclique ou IA dispose du contr�le du pont ---
        ' ATTENTION si l'op�rateur dispose du contr�le des 2 ponts on utilise la fonction pour l'interpr�tation
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
                
                
                '--- lancement du d�placement ---
                ReponseEnvoiCyclePont = EnvoiCyclePont(NumPont, TUnCyclePont)
                Select Case ReponseEnvoiCyclePont
                    
                    Case OK
                         '--- le cycle a �t� transf�r� avec succ�s, il faut remplir la fiche des param�tres ---
                         With TEtatsPonts(NumPont).TParametresCyclesPonts(CYCLES.C_ACTUEL)
                            .NumPosteDepart = TEtatsPonts(NumPont).PosteActuel
                            .NumPosteArrivee = NumPoste
                            .TypeCycle = TYPES_CYCLES.TC_DEPLACEMENT_PONT
                            .DelaiSupStabilisationChargeSecondes = 0
                            .TempsEgouttageSecondes = 0
                         End With
                        
                        '--- affectation de la r�ponse ---
                        CouleurReponse = COULEURS.BLEU_3
                        Reponse = OK
                    
                    Case Else
                        '--- le d�placement a �t� refus� / affectation de la r�ponse ---
                        CouleurReponse = COULEURS.ROUGE_3
                        Reponse = Reponse & "D�placement au poste " & TEtatsPostes(NumPoste).DefinitionPoste.NomPoste & " REFUSE"
                        Reponse = Reponse & vbCrLf & TABULATION_REPONSES & ReponseEnvoiCyclePont
                
                End Select
                    
            Else
            
                '--- mauvaise formulation / affectation de la r�ponse ---
                CouleurReponse = COULEURS.ROUGE_3
                Reponse = MAUVAISE_FORMULATION
            
            End If
    
        Else
    
            '--- pas de disposition du pont / affectation de la r�ponse ---
            CouleurReponse = COULEURS.ROUGE_3
            Reponse = PAS_DE_DISPOSITION_DU_PONT_IA & " " & NumPont
    
        End If
    
    Else
        
        '--- mauvaise formulation / affectation de la r�ponse ---
        CouleurReponse = COULEURS.ROUGE_3
        Reponse = MAUVAISE_FORMULATION
    
    End If

    '--- valeur de retour ---
    AutomatiqueDeplacementPont = Reponse

End Function

'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' R�le      : Transmet le d�placement du pont au poste voulu en AUTOMATIQUE pour l'optimisation avant la fin
'                 d'un temps au poste
' Entr�es :                                  NumPont -> Num�ro du pont concern�
'                                                NumPoste -> Num�ro du poste souhait�
'                                      CouleurReponse -> Couleur de la r�ponse
' Retours : AutomatiqueDeplacementPont -> Message � retourner comme r�ponse
' D�tails  :
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function AutomatiqueDeplacementPontOptimisation(ByVal NumPont As Integer, _
                                                                                                ByVal NumPoste As Integer, _
                                                                                                ByRef CouleurReponse As Long) As String
    
    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next
    
    '--- constantes priv�es ---
    
    
    'Call Log("AutomatiqueDeplacementPontOptimisation entr�e  pont:" & NumPont & " , NumPoste: " & NumPoste)
    '--- d�claration ---
    Dim TUnCyclePont(1 To NBR_LIGNES_CYCLES_PONTS) As Integer
    Dim Reponse As String, _
            ReponseEnvoiCyclePont As String
    
    '--- affectation ---
    AutomatiqueDeplacementPontOptimisation = ""
    
    If NumPont = PONTS.P_1 Or NumPont = PONTS.P_2 Then
                    
        '--- v�rification si le syst�me cyclique ou IA dispose du contr�le du pont ---
        ' ATTENTION si l'op�rateur dispose du contr�le des 2 ponts on utilise la fonction pour l'interpr�tation
        ' des commandes en passant par la fonction GestionCommandesOperateur
        If TEtatsPonts(NumPont).ControleParOperateur = False Or _
            (TEtatsPonts(PONTS.P_1).ControleParOperateur = True And _
            TEtatsPonts(PONTS.P_2).ControleParOperateur = True) And _
            TEtatsPonts(NumPont).PtrEtActionEnCoursAPI.PtrAction = 0 Then
        
            If NumPoste >= POSTES.P_CHGT_1 And NumPoste <= DERNIER_POSTE Then
                        
                '--- ne pas faire de d�placement en optimisation si le poste est vide ou condamn� ---
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
                    
                    '--- lancement du d�placement ---
                    ReponseEnvoiCyclePont = EnvoiCyclePont(NumPont, TUnCyclePont)
                    Select Case ReponseEnvoiCyclePont
                        
                        Case OK
                             '--- le cycle a �t� transf�r� avec succ�s, il faut remplir la fiche des param�tres ---
                             With TEtatsPonts(NumPont).TParametresCyclesPonts(CYCLES.C_ACTUEL)
                                .NumPosteDepart = TEtatsPonts(NumPont).PosteActuel
                                .NumPosteArrivee = NumPoste
                                .TypeCycle = TYPES_CYCLES.TC_DEPLACEMENT_PONT
                                .DelaiSupStabilisationChargeSecondes = 0
                                .TempsEgouttageSecondes = 0
                             End With
                            
                            '--- affectation de la r�ponse ---
                            CouleurReponse = COULEURS.BLEU_3
                            Reponse = OK
                        
                        Case Else
                            '--- le d�placement a �t� refus� / affectation de la r�ponse ---
                            CouleurReponse = COULEURS.ROUGE_3
                            Reponse = Reponse & "D�placement au poste " & TEtatsPostes(NumPoste).DefinitionPoste.NomPoste & " REFUSE"
                            Reponse = Reponse & vbCrLf & TABULATION_REPONSES & ReponseEnvoiCyclePont
                    
                    End Select
                                        
                Else
                            
                    '--- affectation de la r�ponse ---
                    CouleurReponse = COULEURS.BLEU_3
                    Reponse = "DEPLACEMENT INUTILE DU PONT " & NumPont
                    
                End If
                    
            Else
            
                '--- mauvaise formulation / affectation de la r�ponse ---
                CouleurReponse = COULEURS.ROUGE_3
                Reponse = MAUVAISE_FORMULATION
            
            End If
    
        Else
    
            '--- pas de disposition du pont / affectation de la r�ponse ---
            CouleurReponse = COULEURS.ROUGE_3
            Reponse = PAS_DE_DISPOSITION_DU_PONT_IA & " " & NumPont
    
        End If
    
    Else
        
        '--- mauvaise formulation / affectation de la r�ponse ---
        CouleurReponse = COULEURS.ROUGE_3
        Reponse = MAUVAISE_FORMULATION
    
    End If
     'Call Log("AutomatiqueDeplacementPontOptimisation sortie")
    '--- valeur de retour ---
    AutomatiqueDeplacementPontOptimisation = Reponse

End Function

'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' R�le      : Transfert d'une charge au poste voulu en AUTOMATIQUE
' Entr�es :                                     NumPontImpose -> Num�ro du pont impos�
'                                                                                    0 = pas de pont impos�, prendre le num�ro du pont IA pour
'                                                                                          effectuer le transfert
'                                                                                    <> 0, pont impos�, prendre ce num�ro de pont pour
'                                                                                          effectuer le transfert
'                                                    NumPosteDepart -> Num�ro du poste de d�part
'                                                   NumPosteArrivee -> Num�ro du poste d'arriv�e
'                                    TempsEgouttageSecondes -> Temps d'�gouttage en secondes
'                 DelaiSupStabilisationChargeSecondes -> D�lai de stabilisation suppl�mentaire de la charge
'                                                                                    en secondes
' Retours :                      NumPontReelDuTransfert -> num�ro du pont r�el qui va effectuer le transfert
'                                                    CouleurReponse -> Couleur de la r�ponse
'                                 AutomatiqueTransfertCharge -> Message � retourner comme r�ponse
' D�tails  :
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
    
    '--- constantes priv�es ---
    
    '--- d�claration ---
    Dim NumPont As Integer                                                                      'num�ro du pont donn�e par les diagrammes
                                                                                                                  'en cyclique (extrait de la pr�misse)
    Dim NumPontIA As Integer                                                                   'num�ro de pont donn� par le moteur
                                                                                                                  'd'inf�rence (extrait de la pr�misse)
    Dim TUnCyclePont(1 To NBR_LIGNES_CYCLES_PONTS) As Integer 'cycle d'un pont avec tous les temps �
                                                                                                                  'envoyer � l'automate
    Dim TempsCycleSecondes As Long                                                    'temps d'un cycle en secondes
    
    Dim Reponse As String                                                                        'correspond � la variable de retour de la
                                                                                                                  'fonction
    Dim ReponseExtraitPremisseDecodee As String                                'correspond � la r�ponse donn�e � l'extraction
                                                                                                                  'd'une pr�misse d�cod�e
    Dim ReponseEnvoiCyclePont As String                                               'correspond � la r�ponse donn�e � l'envoi
                                                                                                                  'd'un cycle d'un pont
            
    '--- affectation par d�faut ---
    AutomatiqueTransfertCharge = ""
    NumPontReelDuTransfert = 0
    
    
        
    If NumPosteDepart >= POSTES.P_C13 And NumPosteDepart <= POSTES.P_C16 Then
        'Call Log("transfert de l'ano vers la prochaine cuve d'id poste" & NumPosteArrivee & "DelaiSupStabilisationChargeSecondes = " & DelaiSupStabilisationChargeSecondes)
    End If
    
    If NumPosteDepart >= POSTES.P_CHGT_1 And NumPosteDepart <= DERNIER_POSTE And _
       NumPosteArrivee >= POSTES.P_CHGT_1 And NumPosteArrivee <= DERNIER_POSTE Then
    
        '--- effacement du tableau ---
        Erase TUnCyclePont()
            
        '--- extraction de la pr�misse ---
        ReponseExtraitPremisseDecodee = ExtraitPremisseDecodee(NumPosteDepart, _
                                                                                                            NumPosteArrivee, _
                                                                                                            NumPont, _
                                                                                                            NumPontIA, _
                                                                                                            TempsCycleSecondes, _
                                                                                                            TUnCyclePont())
        
        '*****************************************************************************************************************
        '                            d�termination du num�ro de pont rellement choisi pour le transfert
        '*****************************************************************************************************************
        If NumPontImpose = PONTS.P_1 Or NumPontImpose = PONTS.P_2 Then
            NumPontReelDuTransfert = NumPontImpose
        Else
            NumPontReelDuTransfert = NumPontIA
        End If
        
        '--- v�rification si le syst�me cyclique ou IA dispose du contr�le du pont ---
        ' ATTENTION si l'op�rateur dispose du contr�le des 2 ponts on utilise la fonction pour l'interpr�tation
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
        
            '--- v�rification de l'existence de la r�gle ---
            If ReponseExtraitPremisseDecodee = OK Then

                '--- insertion du temps d'�gouttage dans le cycle du pont ---
                If TempsEgouttageSecondes > 0 Then
                    Bidon = InsertionTempsEgouttageDansCyclePont(TempsEgouttageSecondes, TUnCyclePont())
                End If
                
                '--- insertion du d�lai de stabilisation suppl�mentaire de la charge dans le cycle du pont ---
                If DelaiSupStabilisationChargeSecondes > 0 Then
                    Bidon = InsertionDelaiSupStabilisationChargeDansCyclePont(DelaiSupStabilisationChargeSecondes, TUnCyclePont())
                End If
                
                '--- lancement du transfert ---
                ReponseEnvoiCyclePont = EnvoiCyclePont(NumPontReelDuTransfert, TUnCyclePont())
                Select Case ReponseEnvoiCyclePont
                    
                    Case OK
                         '--- le cycle a �t� transf�r� avec succ�s, il faut remplir la fiche des param�tres ---
                         With TEtatsPonts(NumPontReelDuTransfert).TParametresCyclesPonts(CYCLES.C_ACTUEL)
                            .NumPosteDepart = NumPosteDepart
                            .NumPosteArrivee = NumPosteArrivee
                            .TypeCycle = TYPES_CYCLES.TC_TRANSFERT_CHARGE
                            .DelaiSupStabilisationChargeSecondes = DelaiSupStabilisationChargeSecondes
                            .TempsEgouttageSecondes = TempsEgouttageSecondes
                         End With
                        
                        '--- toujours restitu� la valeur du pont IA avec le n� de pont cyclique ---
                        With TPremisses(NumPosteDepart, NumPosteArrivee)
                            .NumPontIA = .NumPont
                        End With
                        
                        '--- affectation de la r�ponse ---
                        CouleurReponse = COULEURS.BLEU_3
                        Reponse = OK
                    
                    Case Else
                        '--- le transfert a �t� refus� / affectation de la r�ponse ---
                        CouleurReponse = COULEURS.ROUGE_3
                        Reponse = Reponse & "Transfert de la charge de " & TEtatsPostes(NumPosteDepart).DefinitionPoste.NomPoste & _
                                          " en " & TEtatsPostes(NumPosteArrivee).DefinitionPoste.NomPoste & _
                                          " avec le pont " & NumPontReelDuTransfert & _
                                          IIf(TempsEgouttageSecondes = 0, "", ", �gouttage " & TempsEgouttageSecondes & " secondes") & _
                                          " REFUSE"
                        Reponse = Reponse & vbCrLf & TABULATION_REPONSES & ReponseEnvoiCyclePont
                
                End Select

            Else

                '--- mauvaise formulation / affectation de la r�ponse ---
                CouleurReponse = COULEURS.ROUGE_3
                Reponse = ReponseExtraitPremisseDecodee

            End If

        Else

            '--- pas de disposition du pont / affectation de la r�ponse ---
            CouleurReponse = COULEURS.ROUGE_3
            Reponse = PAS_DE_DISPOSITION_DU_PONT_IA & " " & NumPontReelDuTransfert

        End If

    Else
        
        '--- pas de disposition du pont / affectation de la r�ponse ---
        CouleurReponse = COULEURS.ROUGE_3
        Reponse = MAUVAISE_FORMULATION
    
    End If
    
   

    '--- valeur de retour ---
    AutomatiqueTransfertCharge = Reponse

End Function

'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' R�le      : Contr�le si une charge est dans un des bains prioritaire avant de lancer un transfert de charge
' Entr�es :              NumPontImpose -> Num�ro du pont impos�
'                                                             0 = pas de pont impos�, prendre le num�ro du pont IA pour effectuer
'                                                            le transfert
'                                                            <> 0, pont impos�, prendre ce num�ro de pont pour effectuer le transfert
'
'                             NumPosteDepart -> Num�ro du poste de d�part
'                            NumPosteArrivee -> Num�ro du poste d'arriv�e
' Retours :             CouleurReponse -> Couleur de la r�ponse
'                 ControleBainsPrioritaire -> Message � retourner comme r�ponse
' D�tails  :
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function ControleBainsPrioritaires(ByVal NumPontImpose As Integer, _
                                                                   ByVal NumPosteDepart As Integer, _
                                                                   ByVal NumPosteArrivee As Integer, _
                                                                   ByRef CouleurReponse As Long) As String
    
    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next
    
    '--- constantes priv�es ---
    Const TOLERANCE_DEPLACEMENT_PONT As Integer = 80               'tol�rance de d�placement d'un pont en secondes avant de confirmer la priorit� d'un bain
    Const TEMPS_ANNULATION_PRIORITE As Integer = -15                    'temps d�pass� dans un bain prioritaire ou l'on annule la priorit�
    
    '--- d�claration ---
    Dim ChargeDansBainPrioritaireP1 As Boolean                                   'indique qu'il y a au moins une charge dans la zone du pont 1
                                                                                                                  'qui est dans un bain prioritaire
    Dim ChargeDansBainPrioritaireP2 As Boolean                                   'indique qu'il y a au moins une charge dans la zone du pont 2
                                                                                                                  'qui est dans un bain prioritaire
    Dim ChargePresenteAuPoste As Boolean                                            'indique une charge pr�sente dans un poste
    
    Dim a As Integer                                                                                   'pour les boucles FOR...NEXT
    
    Dim NumPont As Integer                                                                      'num�ro du pont donn�e par les diagrammes
                                                                                                                  'en cyclique (extrait de la pr�misse)
    Dim NumPontIA As Integer                                                                   'num�ro de pont donn� par le moteur
                                                                                                                  'd'inf�rence (extrait de la pr�misse)
    Dim NumPontReelDuTransfert As Integer                                             'num�ro du pont r�el qui va effectuer le transfert

    Dim NumPosteBainPrioritaireP1 As Integer                                        'num�ro du poste ou se trouve le bain prioritaire pour le pont 1
    Dim NumPosteBainPrioritaireP2 As Integer                                        'num�ro du poste ou se trouve le bain prioritaire pour le pont 2
    
    Dim TUnCyclePont(1 To NBR_LIGNES_CYCLES_PONTS) As Integer 'cycle d'un pont avec tous les temps �
                                                                                                                  'envoyer � l'automate
    Dim TempsCycleSecondes As Long                                                    'temps d'un cycle en secondes
    
    Dim DecompteTempsPostePrioritaireP1 As Long                              'd�compte du temps en secondes du poste prioritaire pour le pont 1
    Dim DecompteTempsPostePrioritaireP2 As Long                              'd�compte du temps en secondes du poste prioritaire pour le pont 2
    
    Dim Reponse As String                                                                        'correspond � la variable de retour de la
                                                                                                                  'fonction
    Dim ReponseExtraitPremisseDecodee As String                                'correspond � la r�ponse donn�e � l'extraction
                                                                                                                  'd'une pr�misse d�cod�e
    Dim ReponseEnvoiCyclePont As String                                               'correspond � la r�ponse donn�e � l'envoi
                                                                                                                  'd'un cycle d'un pont
            
    '----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    
    '--- affectation par d�faut ---
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
                
                '--- affectation du num�ro de poste ou se trouve le bain prioritaire pour le pont 1 ---
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
                    
                    '--- affectation du num�ro de poste ou se trouve le bain prioritaire pour le pont 1 ---
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
                        
        '--- affectation de la r�ponse ---
        CouleurReponse = COULEURS.BLEU_3
        ControleBainsPrioritaires = OK
        
        '--- sortie de la fonction ---
        Exit Function
    
    End If
    
    '----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    
    '--- recherche du pont concern� par le transfert de charge ---
    If NumPosteDepart >= POSTES.P_CHGT_1 And NumPosteDepart <= DERNIER_POSTE And _
       NumPosteArrivee >= POSTES.P_CHGT_1 And NumPosteArrivee <= DERNIER_POSTE Then
    
        '--- effacement du tableau ---
        Erase TUnCyclePont()
            
        '--- extraction de la pr�misse ---
        ReponseExtraitPremisseDecodee = ExtraitPremisseDecodee(NumPosteDepart, _
                                                                                                            NumPosteArrivee, _
                                                                                                            NumPont, _
                                                                                                            NumPontIA, _
                                                                                                            TempsCycleSecondes, _
                                                                                                            TUnCyclePont())
        
        '*****************************************************************************************************************
        '                            d�termination du num�ro de pont r�ellement choisi pour le transfert
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
                
                '--- affectation de la r�ponse ---
                CouleurReponse = COULEURS.BLEU_3
                Reponse = OK
            
            Else
            
                '--- transfert autoris� si le poste de d�part est prioritaire ---
                If TEtatsPostes(NumPosteDepart).DefinitionPoste.RespectTempsObligatoire = True Then

                    '--- affectation de la r�ponse ---
                    CouleurReponse = COULEURS.BLEU_3
                    Reponse = OK

                Else

                    '--- affectation du d�compte du temps en secondes du poste prioritaire ---
                    DecompteTempsPostePrioritaireP1 = RechercheDecompteTempsAuPoste(NumPosteBainPrioritaireP1, ChargePresenteAuPoste)
                    
                    If DecompteTempsPostePrioritaireP1 > TOLERANCE_DEPLACEMENT_PONT Or DecompteTempsPostePrioritaireP1 < TEMPS_ANNULATION_PRIORITE Then 'tol�rance d'un d�placement si le temps est sup�rieur � X minutes
                                                                                                                                                                                                                                                                                 'DecompteTempsPostePrioritaireP1 < TEMPS_ANNULATION_PRIORITE
                                                                                                                                                                                                                                                                                 'car si une charge prioritaire ne peut avancer (cas du poste d'arriv�e occup�)
                                                                                                                                                                                                                                                                                  'il faut forcer l'avance des autres charges le lib�rer le poste d'arriv�e
                        '--- affectation de la r�ponse ---
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
                
                '--- affectation de la r�ponse ---
                CouleurReponse = COULEURS.BLEU_3
                Reponse = OK
            
            Else
            
                '--- transfert autoris� si le poste de d�part est prioritaire ---
                If TEtatsPostes(NumPosteDepart).DefinitionPoste.RespectTempsObligatoire = True Then

                    '--- affectation de la r�ponse ---
                    CouleurReponse = COULEURS.BLEU_3
                    Reponse = OK

                Else

                    '--- affectation du d�compte du temps en secondes du poste prioritaire ---
                    DecompteTempsPostePrioritaireP2 = RechercheDecompteTempsAuPoste(NumPosteBainPrioritaireP2, ChargePresenteAuPoste)
                    
                    If DecompteTempsPostePrioritaireP2 > TOLERANCE_DEPLACEMENT_PONT Or DecompteTempsPostePrioritaireP2 < TEMPS_ANNULATION_PRIORITE Then 'tol�rance d'un d�placement si le temps est sup�rieur � X minutes
                                                                                                                                                                                                                                                                                 'DecompteTempsPostePrioritaireP2 < TEMPS_ANNULATION_PRIORITE
                                                                                                                                                                                                                                                                                 'car si une charge prioritaire ne peut avancer (cas du poste d'arriv�e occup�)
                                                                                                                                                                                                                                                                                 'il faut forcer l'avance des autres charges le lib�rer le poste d'arriv�e
                    
                        '--- affectation de la r�ponse ---
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



Attribute VB_Name = "MGammesAnodisation"
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' R�le                    : MODULE AIDANT A LA GESTION DES GAMMES D'ANODISATION
' Nom                    : MGammesAnodisation.bas
' Date de cr�ation : 13/10/2010
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

'--- d�clarations obligatoires ---
Option Explicit

'--- options g�n�rales ---
Option Base 1
DefVar A-Z

'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' R�le      : Recherche le passage dans le poste de brillantage 'une gamme
' Entr�es : TGammesAnodisation -> Une gamme d'anodisation du type EnrGammesAnodisation
' Retours :    PassageBrillantage ->  TRUE = Passage dans le brillantage dans la gamme
'                                                        FALSE = Pas de passage dans le brillantage dans la gamme
' D�tails  :
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function PassageBrillantage(ByRef TGammesAnodisation As EnrGammesAnodisation) As Boolean

    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next
    
    '--- constantes priv�es ---

    '--- d�claration ---
    Dim a As Integer                                    'pour les boucles FOR...NEXT
    Dim Codezone As String                        'repr�sente le code d'une zone

    '--- affectation par d�faut ---
    PassageBrillantage = False

    '--- analyse dans la gamme ---
    For a = LBound(TGammesAnodisation.TDetailsGammesAnodisation()) To UBound(TGammesAnodisation.TDetailsGammesAnodisation())
    
        If TGammesAnodisation.TDetailsGammesAnodisation(a).NumZone <> 0 Then
    
            '--- affectation ---
            Codezone = TZones(TGammesAnodisation.TDetailsGammesAnodisation(a).NumZone).Codezone
        
            '--- sortie directe si la zone est d�tect�e ---
            If Codezone = "C05 ou C07" Or Codezone = "C07" Then
                PassageBrillantage = True
                Exit For
            End If
        
        End If
        
    Next a

End Function

'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' R�le      : Calcul les temps principaux d'une gamme d'anodisation SANS LES TEMPS DE MOUVEMENTS DES PONTS
' Entr�es :                                             TGammesAnodisation -> Une gamme d'anodisation du type EnrGammesAnodisation
' Retours :        TempsAvantPostePrincipalSansPontsSecondes -> Temps avant arriv�e au Anodisation en secondes
'                             TempsPostePrincipalSansPontsSecondes -> Temps dans le poste d'anodisation en secondes
'                        TempsApresPostePrincipalSansPontsSecondes -> Temps apr�s le poste d'anodisation en secondes
'                        TempsTotalPostesSansPontsSecondes -> Temps total des postes en secondes
'                 TempsTotalEgouttagesSansPontsSecondes -> Temps total des �gouttages en secondes
'                       TempsTotalGammeSansPontsSecondes -> Temps total de la gamme en secondes
' D�tails  :
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Sub CalculTempsGammeAnodisationSansPonts(ByRef TGammesAnodisation As EnrGammesAnodisation, _
                                                                                         ByRef TempsAvantPostePrincipalSansPontsSecondes As Long, _
                                                                                         ByRef TempsPostePrincipalSansPontsSecondes As Long, _
                                                                                         ByRef TempsApresPostePrincipalSansPontsSecondes As Long, _
                                                                                         ByRef TempsTotalPostesSansPontsSecondes As Long, _
                                                                                         ByRef TempsTotalEgouttagesSansPontsSecondes As Long, _
                                                                                         ByRef TempsTotalGammeSansPontsSecondes As Long)
    
    
    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next
    
    '--- constantes priv�es ---

    
    '--- d�claration ---
    Dim PresenceZoneAnodisation As Boolean
    Dim a As Integer, _
           NumZone As Integer
                                                                      
    '--- RAZ des temps ---
    TempsAvantPostePrincipalSansPontsSecondes = 0
    TempsPostePrincipalSansPontsSecondes = 0
    TempsApresPostePrincipalSansPontsSecondes = 0
    TempsTotalPostesSansPontsSecondes = 0
    TempsTotalEgouttagesSansPontsSecondes = 0
    TempsTotalGammeSansPontsSecondes = 0
                                                                      
    '--- calcul des temps ---
    For a = LBound(TGammesAnodisation.TDetailsGammesAnodisation()) To UBound(TGammesAnodisation.TDetailsGammesAnodisation())
    
        With TGammesAnodisation.TDetailsGammesAnodisation(a)
    
            '--- affectation ---
            NumZone = .NumZone
    
            If NumZone >= LIMITE_BASSE_ZONES And NumZone <= LIMITE_HAUTE_ZONES Then
            
                '--- temps d'anodisation ---
                If TZones(NumZone).Codezone = CODE_ZONE_ANODISATION Then
                    TempsPostePrincipalSansPontsSecondes = TempsPostePrincipalSansPontsSecondes + .TempsAuPosteSecondes + .TempsEgouttageSecondes
                    PresenceZoneAnodisation = True
                End If
        
                '--- temps avant Anodisation ---
                If PresenceZoneAnodisation = False Then
                    TempsAvantPostePrincipalSansPontsSecondes = TempsAvantPostePrincipalSansPontsSecondes + .TempsAuPosteSecondes + .TempsEgouttageSecondes
                End If
    
                '--- temps apr�s Anodisation ---
                If TZones(NumZone).Codezone <> CODE_ZONE_ANODISATION And PresenceZoneAnodisation = True Then
                    TempsApresPostePrincipalSansPontsSecondes = TempsApresPostePrincipalSansPontsSecondes + .TempsAuPosteSecondes + .TempsEgouttageSecondes
                End If
            
                '--- temps total des postes ---
                TempsTotalPostesSansPontsSecondes = TempsTotalPostesSansPontsSecondes + .TempsAuPosteSecondes
                
                '--- temps total des �gouttages ---
                TempsTotalEgouttagesSansPontsSecondes = TempsTotalEgouttagesSansPontsSecondes + .TempsEgouttageSecondes
            
                '--- temps total de la gamme ---
                TempsTotalGammeSansPontsSecondes = TempsTotalPostesSansPontsSecondes + TempsTotalEgouttagesSansPontsSecondes
            
            End If
    
        End With
    
    Next a
    
    '--- annulation des temps si pas de passage au Anodisation ---
    If PresenceZoneAnodisation = False Then
        TempsAvantPostePrincipalSansPontsSecondes = 0
        TempsPostePrincipalSansPontsSecondes = 0
        TempsApresPostePrincipalSansPontsSecondes = 0
    End If
    
End Sub

'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' R�le      : Calcul les temps principaux d'une gamme d'anodisation AVEC LES TEMPS DE MOUVEMENTS DES PONTS
'                 le calcul se fait avec le mode d'apprentissage des mouvements
' Entr�es :                                          TGammesAnodisation -> Une gamme d'anodisation du type EnrGammesAnodisation
' Retours : TempsMouvementsAvantPostePrincipalSecondes -> Temps des mouvements avant arriv�e au Anodisation en secondes
'                     TempsAvantPostePrincipalAvecPontsSecondes -> Temps avant arriv�e au Anodisation avec les ponts en secondes
'                          TempsPostePrincipalAvecPontsSecondes -> Temps dans le poste d'anodisation avec les ponts en secondes
'                 TempsMouvementsApresPostePrincipalSecondes -> Temps des mouvements apr�s le poste d'anodisation en secondes
'                     TempsApresPostePrincipalAvecPontsSecondes -> Temps apr�s le poste d'anodisation avec les ponts en secondes
'                     TempsTotalPostesAvecPontsSecondes -> Temps total des postes avec les ponts en secondes
'              TempsTotalEgouttagesAvecPontsSecondes -> Temps total des �gouttages avec les ponts en secondes
'                            TempsTotalMouvementsSecondes -> Temps total des cycles (total du temps des actions) en secondes
'                    TempsTotalGammeAvecPontsSecondes -> Temps total de la gamme avec les ponts en secondes
' D�tails  :
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Sub CalculTempsGammeAnodisationAvecPonts(ByRef TGammesAnodisation As EnrGammesAnodisation, _
                                                                                         ByRef TempsMouvementsAvantPostePrincipalSecondes As Long, _
                                                                                         ByRef TempsAvantPostePrincipalAvecPontsSecondes As Long, _
                                                                                         ByRef TempsPostePrincipalAvecPontsSecondes As Long, _
                                                                                         ByRef TempsMouvementsApresPostePrincipalSecondes As Long, _
                                                                                         ByRef TempsApresPostePrincipalAvecPontsSecondes As Long, _
                                                                                         ByRef TempsTotalPostesAvecPontsSecondes As Long, _
                                                                                         ByRef TempsTotalEgouttagesAvecPontsSecondes As Long, _
                                                                                         ByRef TempsTotalMouvementsSecondes As Long, _
                                                                                         ByRef TempsTotalGammeAvecPontsSecondes As Long)

    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next
    
    '--- constantes priv�es ---
    
    '--- d�claration ---
    Dim PresenceZoneAnodisation As Boolean
    Dim a As Integer, _
           NumZoneDepart As Integer, _
           NumZoneArrivee As Integer, _
           NumPosteDepart As Integer, _
           NumPosteArrivee As Integer
    Dim TempsAvantPostePrincipalSansPontsSecondes As Long, _
            TempsPostePrincipalSansPontsSecondes As Long, _
            TempsApresPostePrincipalSansPontsSecondes As Long, _
            TempsTotalPostesSansPontsSecondes As Long, _
            TempsTotalEgouttagesSansPontsSecondes As Long, _
            TempsTotalGammeSansPontsSecondes As Long
    Dim TempsCycleSecondes As Long                             'temps d'un cycle complet en secondes
    Dim CodeZoneDepart As String, _
            CodeZoneArrivee As String
                                                                      
    '--- RAZ des temps transmis par r�f�rence ---
    TempsMouvementsAvantPostePrincipalSecondes = 0
    TempsAvantPostePrincipalAvecPontsSecondes = 0
    TempsPostePrincipalAvecPontsSecondes = 0
    TempsMouvementsApresPostePrincipalSecondes = 0
    TempsApresPostePrincipalAvecPontsSecondes = 0
    TempsTotalPostesAvecPontsSecondes = 0
    TempsTotalEgouttagesAvecPontsSecondes = 0
    TempsTotalMouvementsSecondes = 0
    TempsTotalGammeAvecPontsSecondes = 0
                                                                      
    '--- appel de la routine de calcul des temps SANS les ponts ---
    CalculTempsGammeAnodisationSansPonts TGammesAnodisation, _
                                                                           TempsAvantPostePrincipalSansPontsSecondes, _
                                                                           TempsPostePrincipalSansPontsSecondes, _
                                                                           TempsApresPostePrincipalSansPontsSecondes, _
                                                                           TempsTotalPostesSansPontsSecondes, _
                                                                           TempsTotalEgouttagesSansPontsSecondes, _
                                                                           TempsTotalGammeSansPontsSecondes

    '--- affectation de base avec les valeurs sans ponts ---
    TempsAvantPostePrincipalAvecPontsSecondes = TempsAvantPostePrincipalSansPontsSecondes
    TempsPostePrincipalAvecPontsSecondes = TempsPostePrincipalSansPontsSecondes
    TempsApresPostePrincipalAvecPontsSecondes = TempsApresPostePrincipalSansPontsSecondes
    TempsTotalPostesAvecPontsSecondes = TempsTotalPostesSansPontsSecondes
    TempsTotalEgouttagesAvecPontsSecondes = TempsTotalEgouttagesSansPontsSecondes
    TempsTotalGammeAvecPontsSecondes = TempsTotalGammeSansPontsSecondes

    '--- calcul des temps EN AJOUTANT LES MOUVEMENTS DES PONTS ---
    For a = LBound(TGammesAnodisation.TDetailsGammesAnodisation()) To Pred(UBound(TGammesAnodisation.TDetailsGammesAnodisation()))

        '--- affectation ---
        NumZoneDepart = TGammesAnodisation.TDetailsGammesAnodisation(a).NumZone
        NumZoneArrivee = TGammesAnodisation.TDetailsGammesAnodisation(Succ(a)).NumZone
 
        If NumZoneDepart >= LIMITE_BASSE_ZONES And NumZoneDepart <= LIMITE_HAUTE_ZONES And _
           NumZoneArrivee >= LIMITE_BASSE_ZONES And NumZoneArrivee <= LIMITE_HAUTE_ZONES Then

            '--- affectation ---
            CodeZoneDepart = TZones(NumZoneDepart).Codezone
            CodeZoneArrivee = TZones(NumZoneArrivee).Codezone
            
            '--- recherche du poste de d�part de la zone (ATTENTION AUX ZONES A POSTES MULTIPLES) ---
            Select Case CodeZoneDepart
                Case "CHGT1 � CHGT4": NumPosteDepart = POSTES.P_CHGT_1                         'poste � distance moyenne
                Case CODE_ZONE_ANODISATION: NumPosteDepart = POSTES.P_C13                'poste le plus loin en zone de d�part
                Case "D1 � D2": NumPosteDepart = POSTES.P_D1                                               'poste � distance moyenne
                Case Else: NumPosteDepart = TZones(NumZoneDepart).NumPremierPoste
            End Select
            
            '--- recherche du poste d'arriv�e de la zone (ATTENTION AUX ZONES A POSTES MULTIPLES) ---
            Select Case CodeZoneArrivee
                Case "CHGT1 � CHGT4": NumPosteArrivee = POSTES.P_CHGT_1                         'poste � distance moyenne
                Case CODE_ZONE_ANODISATION: NumPosteArrivee = POSTES.P_C13                'poste le plus loin en zone d'arriv�e
                Case "D1 � D2": NumPosteArrivee = POSTES.P_D2                                               'poste � distance moyenne
                Case Else: NumPosteArrivee = TZones(NumZoneArrivee).NumPremierPoste
            End Select
            
            '--- calcul du temps et affectation dans les pr�misses pour mise � jour ---
            With TPremisses(NumPosteDepart, NumPosteArrivee)
                If CalculTempsCyclePremisse(NumPosteDepart, NumPosteArrivee, TempsCycleSecondes) = OK Then
                    .TempsCycleSecondes = TempsCycleSecondes
                Else
                    .TempsCycleSecondes = 0
                End If
            End With
            
            '--- calcul du temps total des cycles en secondes ---
            TempsTotalMouvementsSecondes = TempsTotalMouvementsSecondes + TempsCycleSecondes

            '--- affectation ---
            If CodeZoneDepart = CODE_ZONE_ANODISATION Then
                PresenceZoneAnodisation = True
            End If
            
            If PresenceZoneAnodisation = False Then
                
                '--- temps avant Anodisation ---
                TempsAvantPostePrincipalAvecPontsSecondes = TempsAvantPostePrincipalAvecPontsSecondes + TempsCycleSecondes
            
            Else

                '--- temps apr�s Anodisation ---
                TempsApresPostePrincipalAvecPontsSecondes = TempsApresPostePrincipalAvecPontsSecondes + TempsCycleSecondes
            
            End If
    
        End If

    Next a
    
    '--- temps total de la gamme ---
    TempsTotalGammeAvecPontsSecondes = TempsTotalGammeAvecPontsSecondes + TempsTotalMouvementsSecondes
    
    If PresenceZoneAnodisation = False Then
        
        '--- annulation des temps si pas de passage au Anodisation ---
        TempsMouvementsAvantPostePrincipalSecondes = 0
        TempsAvantPostePrincipalAvecPontsSecondes = 0
        TempsPostePrincipalAvecPontsSecondes = 0
        TempsMouvementsApresPostePrincipalSecondes = 0
        TempsApresPostePrincipalAvecPontsSecondes = 0
    
    Else
    
        '--- temps des mouvements des ponts avant Anodisation ---
        TempsMouvementsAvantPostePrincipalSecondes = TempsAvantPostePrincipalAvecPontsSecondes - TempsAvantPostePrincipalSansPontsSecondes
        
        '--- temps des mouvements des ponts apr�s Anodisation ---
        TempsMouvementsApresPostePrincipalSecondes = TempsApresPostePrincipalAvecPontsSecondes - TempsApresPostePrincipalSansPontsSecondes
    
    End If
    
End Sub



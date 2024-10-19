Attribute VB_Name = "MGammesProduction"
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' R�le                    : MODULE AIDANT A LA GESTION DES GAMMES DE PRODUCTION
' Nom                    : MGammesProduction.bas
' Date de cr�ation : 13/10/2010
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

'--- d�clarations obligatoires ---
Option Explicit

'--- options g�n�rales ---
Option Base 1
DefVar A-Z

'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' R�le      : Calcul les temps principaux d'une gamme d'anodisation SANS LES TEMPS DE MOUVEMENTS DES PONTS
' Entr�es :                                             TGammesAnodisation -> Une gamme d'anodisation du type EnrGammesAnodisation
' Retours :        TempsAvantAnodisationSansPontsSecondes -> Temps avant arriv�e au Anodisation en secondes
'                             TempsAuAnodisationSansPontsSecondes -> Temps dans le poste d'anodisation en secondes
'                        TempsApresAnodisationSansPontsSecondes -> Temps apr�s le poste d'anodisation en secondes
'                        TempsTotalPostesSansPontsSecondes -> Temps total des postes en secondes
'                 TempsTotalEgouttagesSansPontsSecondes -> Temps total des �gouttages en secondes
'                       TempsTotalGammeSansPontsSecondes -> Temps total de la gamme en secondes
' D�tails  :
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Sub CalculTempsGammeAnodisationSansPonts(ByRef TGammesAnodisation As EnrGammesAnodisation, _
                                                                                         ByRef TempsAvantAnodisationSansPontsSecondes As Long, _
                                                                                         ByRef TempsAuAnodisationSansPontsSecondes As Long, _
                                                                                         ByRef TempsApresAnodisationSansPontsSecondes As Long, _
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
    TempsAvantAnodisationSansPontsSecondes = 0
    TempsAuAnodisationSansPontsSecondes = 0
    TempsApresAnodisationSansPontsSecondes = 0
    TempsTotalPostesSansPontsSecondes = 0
    TempsTotalEgouttagesSansPontsSecondes = 0
    TempsTotalGammeSansPontsSecondes = 0
                                                                      
    '--- calcul des temps ---
    For a = LBound(TGammesAnodisation.TDetailsGammesAnodisation()) To UBound(TGammesAnodisation.TDetailsGammesAnodisation())
    
        With TGammesAnodisation.TDetailsGammesAnodisation(a)
    
            '--- affectation ---
            NumZone = .NumZone
    
            If NumZone >= LIMITE_BASSE_ZONES And NumZone <= LIMITE_HAUTE_ZONES Then
            
                '--- temps du Anodisation ---
                If TZones(NumZone).Codezone = TEXTE_CODE_ZONE_Anodisation Then
                    TempsAuAnodisationSansPontsSecondes = TempsAuAnodisationSansPontsSecondes + .TempsAuPosteSecondes + .TempsEgouttageSecondes
                    PresenceZoneAnodisation = True
                End If
        
                '--- temps avant Anodisation ---
                If PresenceZoneAnodisation = False Then
                    TempsAvantAnodisationSansPontsSecondes = TempsAvantAnodisationSansPontsSecondes + .TempsAuPosteSecondes + .TempsEgouttageSecondes
                End If
    
                '--- temps apr�s Anodisation ---
                If TZones(NumZone).Codezone <> TEXTE_CODE_ZONE_Anodisation And PresenceZoneAnodisation = True Then
                    TempsApresAnodisationSansPontsSecondes = TempsApresAnodisationSansPontsSecondes + .TempsAuPosteSecondes + .TempsEgouttageSecondes
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
        TempsAvantAnodisationSansPontsSecondes = 0
        TempsAuAnodisationSansPontsSecondes = 0
        TempsApresAnodisationSansPontsSecondes = 0
    End If
    
End Sub

'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' R�le      : Calcul les temps principaux d'une gamme d'anodisation AVEC LES TEMPS DE MOUVEMENTS DES PONTS
'                 le calcul se fait avec le mode d'apprentissage des mouvements
' Entr�es :                                          TGammesAnodisation -> Une gamme d'anodisation du type EnrGammesAnodisation
' Retours : TempsMouvementsAvantAnodisationSecondes -> Temps des mouvements avant arriv�e au Anodisation en secondes
'                     TempsAvantAnodisationAvecPontsSecondes -> Temps avant arriv�e au Anodisation avec les ponts en secondes
'                          TempsAuAnodisationAvecPontsSecondes -> Temps dans le poste d'anodisation avec les ponts en secondes
'                 TempsMouvementsApresAnodisationSecondes -> Temps des mouvements apr�s le poste d'anodisation en secondes
'                     TempsApresAnodisationAvecPontsSecondes -> Temps apr�s le poste d'anodisation avec les ponts en secondes
'                     TempsTotalPostesAvecPontsSecondes -> Temps total des postes avec les ponts en secondes
'              TempsTotalEgouttagesAvecPontsSecondes -> Temps total des �gouttages avec les ponts en secondes
'                            TempsTotalMouvementsSecondes -> Temps total des cycles (total du temps des actions) en secondes
'                    TempsTotalGammeAvecPontsSecondes -> Temps total de la gamme avec les ponts en secondes
' D�tails  :
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Sub CalculTempsGammeAnodisationAvecPonts(ByRef TGammesAnodisation As EnrGammesAnodisation, _
                                                                                ByRef TempsMouvementsAvantAnodisationSecondes As Long, _
                                                                                ByRef TempsAvantAnodisationAvecPontsSecondes As Long, _
                                                                                ByRef TempsAuAnodisationAvecPontsSecondes As Long, _
                                                                                ByRef TempsMouvementsApresAnodisationSecondes As Long, _
                                                                                ByRef TempsApresAnodisationAvecPontsSecondes As Long, _
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
    Dim TempsAvantAnodisationSansPontsSecondes As Long, _
            TempsAuAnodisationSansPontsSecondes As Long, _
            TempsApresAnodisationSansPontsSecondes As Long, _
            TempsTotalPostesSansPontsSecondes As Long, _
            TempsTotalEgouttagesSansPontsSecondes As Long, _
            TempsTotalGammeSansPontsSecondes As Long
    Dim TempsCycleSecondes As Long                             'temps d'un cycle complet en secondes
    Dim CodeZoneDepart As String, _
            CodeZoneArrivee As String
                                                                      
    '--- RAZ des temps transmis par r�f�rence ---
    TempsMouvementsAvantAnodisationSecondes = 0
    TempsAvantAnodisationAvecPontsSecondes = 0
    TempsAuAnodisationAvecPontsSecondes = 0
    TempsMouvementsApresAnodisationSecondes = 0
    TempsApresAnodisationAvecPontsSecondes = 0
    TempsTotalPostesAvecPontsSecondes = 0
    TempsTotalEgouttagesAvecPontsSecondes = 0
    TempsTotalMouvementsSecondes = 0
    TempsTotalGammeAvecPontsSecondes = 0
                                                                      
    '--- appel de la routine de calcul des temps SANS les ponts ---
    CalculTempsGammeAnodisationSansPonts TGammesAnodisation, _
                                                                  TempsAvantAnodisationSansPontsSecondes, _
                                                                  TempsAuAnodisationSansPontsSecondes, _
                                                                  TempsApresAnodisationSansPontsSecondes, _
                                                                  TempsTotalPostesSansPontsSecondes, _
                                                                  TempsTotalEgouttagesSansPontsSecondes, _
                                                                  TempsTotalGammeSansPontsSecondes

    '--- affectation de base avec les valeurs sans ponts ---
    TempsAvantAnodisationAvecPontsSecondes = TempsAvantAnodisationSansPontsSecondes
    TempsAuAnodisationAvecPontsSecondes = TempsAuAnodisationSansPontsSecondes
    TempsApresAnodisationAvecPontsSecondes = TempsApresAnodisationSansPontsSecondes
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
            'Select Case CodeZoneDepart
            '    Case "C1 � C6": NumPosteDepart = POSTES.P_C3        'poste � distance moyenne
            '    Case TEXTE_CODE_ZONE_Anodisation: NumPosteDepart = POSTES.P_C13   'poste le plus loin en zone de d�part
            '    Case "D1 � D6": NumPosteDepart = POSTES.P_D3        'poste � distance moyenne
            '    Case Else: NumPosteDepart = TZones(NumZoneDepart).NumPremierPoste
            'End Select
            
            '--- recherche du poste d'arriv�e de la zone (ATTENTION AUX ZONES A POSTES MULTIPLES) ---
            'Select Case CodeZoneArrivee
            '    Case "C1 � C6": NumPosteArrivee = POSTES.P_C3        'poste � distance moyenne
            '    Case TEXTE_CODE_ZONE_Anodisation: NumPosteArrivee = POSTES.P_C15   'poste le plus loin en zone d'arriv�e
            '    Case "D1 � D6": NumPosteArrivee = POSTES.P_D3        'poste � distance moyenne
            '    Case Else: NumPosteArrivee = TZones(NumZoneArrivee).NumPremierPoste
            'End Select
            
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
            If CodeZoneDepart = TEXTE_CODE_ZONE_Anodisation Then
                PresenceZoneAnodisation = True
            End If
            
            If PresenceZoneAnodisation = False Then
                
                '--- temps avant Anodisation ---
                TempsAvantAnodisationAvecPontsSecondes = TempsAvantAnodisationAvecPontsSecondes + TempsCycleSecondes
            
            Else

                '--- temps apr�s Anodisation ---
                TempsApresAnodisationAvecPontsSecondes = TempsApresAnodisationAvecPontsSecondes + TempsCycleSecondes
            
            End If
    
        End If

    Next a
    
    '--- temps total de la gamme ---
    TempsTotalGammeAvecPontsSecondes = TempsTotalGammeAvecPontsSecondes + TempsTotalMouvementsSecondes
    
    If PresenceZoneAnodisation = False Then
        
        '--- annulation des temps si pas de passage au Anodisation ---
        TempsMouvementsAvantAnodisationSecondes = 0
        TempsAvantAnodisationAvecPontsSecondes = 0
        TempsAuAnodisationAvecPontsSecondes = 0
        TempsMouvementsApresAnodisationSecondes = 0
        TempsApresAnodisationAvecPontsSecondes = 0
    
    Else
    
        '--- temps des mouvements des ponts avant Anodisation ---
        TempsMouvementsAvantAnodisationSecondes = TempsAvantAnodisationAvecPontsSecondes - TempsAvantAnodisationSansPontsSecondes
        
        '--- temps des mouvements des ponts apr�s Anodisation ---
        TempsMouvementsApresAnodisationSecondes = TempsApresAnodisationAvecPontsSecondes - TempsApresAnodisationSansPontsSecondes
    
    End If
    
End Sub



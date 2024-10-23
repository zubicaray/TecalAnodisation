Attribute VB_Name = "MEntreesCharges"
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' R�le                    : MODULE GERANT L'ENTREES DES CHARGES
' Nom                    : MEntreesCharges.bas
' Date de cr�ation : 03/02/2010
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

'--- d�clarations obligatoires ---
Option Explicit

'--- options g�n�rales ---
Option Base 1
DefVar A-Z

'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' R�le      : Recherche le nombre de charges dans la zone de brillantage
' Entr�es :
' Retours : RechercheNbrChargesEnBrillantage -> Le nombre de charges dans la zone de brillantage
' D�tails  :
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function RechercheNbrChargesEnBrillantage() As Integer

    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next

    '--- d�claration ---
    Dim a As Integer                                                                  'r�serv� pour les boucles FOR ... NEXT
    Dim NumCharge As Integer                                                 'indique un num�ro de charge
    
    '--- affectation par d�faut ---
    RechercheNbrChargesEnBrillantage = 0

    '********************************************************************************************************************
    '                                                 V�rification pour les postes de la ligne
    '********************************************************************************************************************
    For a = POSTES.P_C05 To POSTES.P_C09
         
        '--- affectation du num�ro de charge ---
        NumCharge = TEtatsPostes(a).NumCharge

        If NumCharge >= CHARGES.C_NUM_MINI And NumCharge <= CHARGES.C_NUM_MAXI Then

            '--- incr�mentation du nombre de charges dans la zone concern� ---
            Inc RechercheNbrChargesEnBrillantage

        End If

    Next a
    
End Function

'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' R�le      : Recherche dans la zone de pr�paration le nombre de charges avec passage dans la coloration noir
' Entr�es :
' Retours : RechercheNbrChargesEnBrillantage -> Le nombre de charges dans la zone de brillantage
' D�tails  :
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function RechercheEnPreparationNbrChargesGammeNoir() As Integer

    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next

    '--- constantes priv�es ---
    Const ZONE_CONCERNEE As String = "C28"

    '--- d�claration ---
    Dim a As Integer                                                                  'r�serv� pour les boucles FOR ... NEXT
    Dim b As Integer                                                                  'r�serv� pour les boucles FOR ... NEXT
    Dim NumCharge As Integer                                                 'indique un num�ro de charge
    Dim NumZone As Integer                                                    'repr�sente un num�ro de zone
    
    '--- affectation par d�faut ---
    RechercheEnPreparationNbrChargesGammeNoir = 0

    '********************************************************************************************************************
    '                                                 V�rification pour les postes de la ligne
    '********************************************************************************************************************
    For a = POSTES.P_CHGT_1 To POSTES.P_C12
         
        '--- affectation du num�ro de charge ---
        NumCharge = TEtatsPostes(a).NumCharge

        If NumCharge >= CHARGES.C_NUM_MINI And NumCharge <= CHARGES.C_NUM_MAXI Then

            For b = 1 To NBR_LIGNES_DETAILS_GAMMES_PRODUCTION
            
                '--- affectation du num�ro de zone ---
                NumZone = TEtatsCharges(NumCharge).TGammesAnodisation.TDetailsGammesAnodisation(b).NumZone
                
                If NumZone = 0 Then
                
                    '--- sortie directe si plus de zone dans la gamme ---
                    Exit For
                
                Else
                    
                    If Trim(TZones(NumZone).Codezone) = ZONE_CONCERNEE Then
                        
                        '--- incr�mentation du nombre de charges dans la zone concern� ---
                        Inc RechercheEnPreparationNbrChargesGammeNoir

                    End If
        
                End If
            
            Next b
            

        End If
                
        '--- sortie directe si plus de zone dans la gamme ---
        If NumZone = 0 Then Exit For

    Next a
    
End Function

'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' R�le      : Recherche dans la zone de pr�paration le nombre de charges avec passage dans la spectrocoloration
' Entr�es :
' Retours : RechercheNbrChargesEnBrillantage -> Le nombre de charges dans la zone de brillantage
' D�tails  :
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function RechercheEnPreparationNbrChargesGammeSpectrocoloration() As Integer

    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next

    '--- constantes priv�es ---
    Const ZONE_CONCERNEE As String = "C19"

    '--- d�claration ---
    Dim a As Integer                                                                  'r�serv� pour les boucles FOR ... NEXT
    Dim b As Integer                                                                  'r�serv� pour les boucles FOR ... NEXT
    Dim NumCharge As Integer                                                 'indique un num�ro de charge
    Dim NumZone As Integer                                                    'repr�sente un num�ro de zone
    
    '--- affectation par d�faut ---
    RechercheEnPreparationNbrChargesGammeSpectrocoloration = 0

    '********************************************************************************************************************
    '                                                 V�rification pour les postes de la ligne
    '********************************************************************************************************************
    For a = POSTES.P_CHGT_1 To POSTES.P_C12
         
        '--- affectation du num�ro de charge ---
        NumCharge = TEtatsPostes(a).NumCharge

        If NumCharge >= CHARGES.C_NUM_MINI And NumCharge <= CHARGES.C_NUM_MAXI Then

            For b = 1 To NBR_LIGNES_DETAILS_GAMMES_PRODUCTION
            
                '--- affectation du num�ro de zone ---
                NumZone = TEtatsCharges(NumCharge).TGammesAnodisation.TDetailsGammesAnodisation(b).NumZone
                
                If NumZone = 0 Then
                
                    '--- sortie directe si plus de zone dans la gamme ---
                    Exit For
                
                Else
                    
                    If Trim(TZones(NumZone).Codezone) = ZONE_CONCERNEE Then
                        
                        '--- incr�mentation du nombre de charges dans la zone concern� ---
                        Inc RechercheEnPreparationNbrChargesGammeSpectrocoloration

                    End If
        
                End If
            
            Next b
            

        End If
                
        '--- sortie directe si plus de zone dans la gamme ---
        If NumZone = 0 Then Exit For

    Next a
    
End Function

'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' R�le      : Recherche dans la zone de pr�paration le nombre de charges avec passage dans la coloration or
' Entr�es :
' Retours : RechercheNbrChargesEnBrillantage -> Le nombre de charges dans la zone de brillantage
' D�tails  :
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function RechercheEnPreparationNbrChargesGammeOr() As Integer

    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next

    '--- constantes priv�es ---
    Const ZONE_CONCERNEE As String = "C22"

    '--- d�claration ---
    Dim a As Integer                                                                  'r�serv� pour les boucles FOR ... NEXT
    Dim b As Integer                                                                  'r�serv� pour les boucles FOR ... NEXT
    Dim NumCharge As Integer                                                 'indique un num�ro de charge
    Dim NumZone As Integer                                                    'repr�sente un num�ro de zone
    
    '--- affectation par d�faut ---
    RechercheEnPreparationNbrChargesGammeOr = 0

    '********************************************************************************************************************
    '                                                 V�rification pour les postes de la ligne
    '********************************************************************************************************************
    For a = POSTES.P_CHGT_1 To POSTES.P_C12
         
        '--- affectation du num�ro de charge ---
        NumCharge = TEtatsPostes(a).NumCharge

        If NumCharge >= CHARGES.C_NUM_MINI And NumCharge <= CHARGES.C_NUM_MAXI Then

            For b = 1 To NBR_LIGNES_DETAILS_GAMMES_PRODUCTION
            
                '--- affectation du num�ro de zone ---
                NumZone = TEtatsCharges(NumCharge).TGammesAnodisation.TDetailsGammesAnodisation(b).NumZone
                
                If NumZone = 0 Then
                
                    '--- sortie directe si plus de zone dans la gamme ---
                    Exit For
                
                Else
                    
                    If Trim(TZones(NumZone).Codezone) = ZONE_CONCERNEE Then
                        
                        '--- incr�mentation du nombre de charges dans la zone concern� ---
                        Inc RechercheEnPreparationNbrChargesGammeOr

                    End If
        
                End If
            
            Next b
            

        End If
                
        '--- sortie directe si plus de zone dans la gamme ---
        If NumZone = 0 Then Exit For

    Next a
    
End Function

'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' R�le      : Recherche le nombre de charges dans la zone de d�graissage / satinage
' Entr�es :
' Retours : RechercheNbrChargesEnBrillantage -> Le nombre de charges dans la zone de brillantage
' D�tails  :
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function RechercheNbrChargesEnDegraissageSatinage() As Integer

    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next

    '--- d�claration ---
    Dim a As Integer                                                                  'r�serv� pour les boucles FOR ... NEXT
    Dim NumCharge As Integer                                                 'indique un num�ro de charge
    
    '--- affectation par d�faut ---
    RechercheNbrChargesEnDegraissageSatinage = 0

    '********************************************************************************************************************
    '                                                 V�rification pour les postes de la ligne
    '********************************************************************************************************************
    For a = POSTES.P_C00 To POSTES.P_C04
         
        '--- affectation du num�ro de charge ---
        NumCharge = TEtatsPostes(a).NumCharge

        If NumCharge >= CHARGES.C_NUM_MINI And NumCharge <= CHARGES.C_NUM_MAXI Then

            '--- incr�mentation du nombre de charges dans la zone concern� ---
            Inc RechercheNbrChargesEnDegraissageSatinage

        End If

    Next a
    
End Function

'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' R�le      : Recherche le nombre de charges dans la zone de pr�paration
' Entr�es :
' Retours : RechercheNbrChargesEnPreparation -> Le nombre de charges dans la zone de pr�paration
' D�tails  :
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function RechercheNbrChargesEnPreparation() As Integer
    
    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next

    '--- d�claration ---
    Dim a As Integer                                                                  'r�serv� pour les boucles FOR ... NEXT
    Dim NumCharge As Integer                                                 'indique un num�ro de charge
    Dim NumPosteAnodisation As Integer                                'num�ro du poste d'anodisation
    Dim NbrChargesEnPreparation As Integer                          'nombre de charges dans la zone de pr�paration
    
    Dim DateEntreeAuPosteAnodisation As Date                      'date entr�e au poste d'anodisation
                        
    '--- affectation par d�faut ---
    NbrChargesEnPreparation = 0
    RechercheNbrChargesEnPreparation = 0
                        
    '********************************************************************************************************************
    '                                                 V�rification pour les postes de la ligne
    '********************************************************************************************************************
    For a = POSTES.P_C00 To POSTES.P_C12
         
        '--- affectation du num�ro de charge ---
        NumCharge = TEtatsPostes(a).NumCharge

        If NumCharge >= CHARGES.C_NUM_MINI And NumCharge <= CHARGES.C_NUM_MAXI Then

            '--- incr�mentation du nombre de charges en pr�paration ---
            Inc NbrChargesEnPreparation

        End If

    Next a
    
    '********************************************************************************************************************
    '                                                             V�rification pour les ponts
    '********************************************************************************************************************
    For a = LBound(TEtatsPonts()) To UBound(TEtatsPonts())
         
        '--- affectation du num�ro de charge ---
        NumCharge = TEtatsPonts(a).NumCharge
                 
        If NumCharge >= CHARGES.C_NUM_MINI And NumCharge <= CHARGES.C_NUM_MAXI Then
             
            '--- contr�ler que l'on n'est jamais passer au anodisation ---
            If RechercheTempsAuPosteAnodisation(NumCharge, NumPosteAnodisation, DateEntreeAuPosteAnodisation) = 0 Then
            
                '--- incr�mentation du nombre de charges en pr�paration ---
                Inc NbrChargesEnPreparation
            
            End If
                 
        End If
                 
    Next a

    '--- valeur de retour ---
    RechercheNbrChargesEnPreparation = NbrChargesEnPreparation

End Function

'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' R�le      : V�rification qu'une charge en pr�paration � un poste d'anodisation impos� dans sa gamme
' Entr�es :
' D�tails  :
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function VerificationChargeEnPreparationAvecAnodisationImpose(ByVal NumPoste As POSTES) As Boolean

    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next

    '--- d�claration ---
    Dim a As Integer                                                                  'r�serv� pour les boucles FOR ... NEXT
    Dim NumCharge As Integer                                                 'indique un num�ro de charge
    Dim NumPosteAnodisation As Integer                                 'num�ro du poste d'anodisation
    
    Dim DateEntreeAuPosteAnodisation As Date                      'date entr�e au poste d'anodisation
                        
    '--- affectation par d�faut ---
    VerificationChargeEnPreparationAvecAnodisationImpose = False
                        
    '********************************************************************************************************************
    '                                                V�rification pour les postes de la ligne
    '********************************************************************************************************************
    For a = POSTES.P_C00 To POSTES.P_C12
         
        '--- ne prendre que la partie pr�paration ---
        NumCharge = TEtatsPostes(a).NumCharge
        
        If NumCharge >= CHARGES.C_NUM_MINI And NumCharge <= CHARGES.C_NUM_MAXI Then
        
           Select Case NumPoste
        
               Case POSTES.P_C13
                   '--- postes C13 ---
                   If TEtatsCharges(NumCharge).TGammesAnodisation.ChoixPosteAnodisation = C_C13_IMPOSE Then
                       VerificationChargeEnPreparationAvecAnodisationImpose = True
                       Exit For
                   End If
        
               Case POSTES.P_C14
                   '--- postes C14 ---
                   If TEtatsCharges(NumCharge).TGammesAnodisation.ChoixPosteAnodisation = C_C14_IMPOSE Then
                       VerificationChargeEnPreparationAvecAnodisationImpose = True
                       Exit For
                   End If
        
               Case POSTES.P_C15
                   '--- postes C15 ---
                   If TEtatsCharges(NumCharge).TGammesAnodisation.ChoixPosteAnodisation = C_C15_IMPOSE Then
                       VerificationChargeEnPreparationAvecAnodisationImpose = True
                       Exit For
                   End If
               
               'Case POSTES.P_C16
                   '--- postes C16 ---
                   'If TEtatsCharges(NumCharge).TGammesAnodisation.ChoixPosteAnodisation = C_C16_IMPOSE Then
                   '    VerificationChargeEnPreparationAvecAnodisationImpose = True
                   '    Exit For
                   'End If
        
                Case Else
           End Select
        
        End If
                         
    Next a
    
    '********************************************************************************************************************
    '                                                                V�rification pour les ponts
    '********************************************************************************************************************
    For a = LBound(TEtatsPonts()) To UBound(TEtatsPonts())
         
        '--- affectation du num�ro de charge ---
        NumCharge = TEtatsPonts(a).NumCharge
                 
        If NumCharge >= CHARGES.C_NUM_MINI And NumCharge <= CHARGES.C_NUM_MAXI Then
             
            '--- contr�ler que l'on n'est jamais passer au Anodisation ---
            If RechercheTempsAuPosteAnodisation(NumCharge, NumPosteAnodisation, DateEntreeAuPosteAnodisation) = 0 Then
            
                Select Case NumPoste
                     
                    Case POSTES.P_C13
                        '--- postes C13 ---
                        If TEtatsCharges(NumCharge).TGammesAnodisation.ChoixPosteAnodisation = C_C13_IMPOSE Then
                            VerificationChargeEnPreparationAvecAnodisationImpose = True
                            Exit For
                        End If
                    
                    Case POSTES.P_C14
                        '--- postes C14 ---
                        If TEtatsCharges(NumCharge).TGammesAnodisation.ChoixPosteAnodisation = C_C14_IMPOSE Then
                            VerificationChargeEnPreparationAvecAnodisationImpose = True
                            Exit For
                        End If
                    
                    Case POSTES.P_C15
                        '--- postes C15 ---
                        If TEtatsCharges(NumCharge).TGammesAnodisation.ChoixPosteAnodisation = C_C15_IMPOSE Then
                            VerificationChargeEnPreparationAvecAnodisationImpose = True
                            Exit For
                        End If
                    
                    'Case POSTES.P_C16
                        '--- postes C16 ---
                        'If TEtatsCharges(NumCharge).TGammesAnodisation.ChoixPosteAnodisation = C_C16_IMPOSE Then
                        '    VerificationChargeEnPreparationAvecAnodisationImpose = True
                        '    Exit For
                        'End If
                     
                     Case Else
                End Select
                 
            End If
                 
        End If
                 
    Next a

End Function

'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' R�le      : V�rification qu'une charge en pr�paration � un poste d'anodisation impos� dans sa gamme
' Entr�es :
' D�tails  :
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function NbrChargesEnPreparationAvecAnodisationImpose(ByVal NumPoste As POSTES) As Integer

    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next

    '--- d�claration ---
    Dim a As Integer                                                                  'r�serv� pour les boucles FOR ... NEXT
    Dim NumCharge As Integer                                                 'indique un num�ro de charge
    Dim NumPosteAnodisation As Integer                                         'num�ro du poste d'anodisation
    
    Dim DateEntreeAuPosteAnodisation As Date                               'date entr�e au poste d'anodisation
                        
    '--- affectation par d�faut ---
    NbrChargesEnPreparationAvecAnodisationImpose = 0
                        
    '********************************************************************************************************************
    '                                                V�rification pour les postes de la ligne
    '********************************************************************************************************************
    For a = POSTES.P_C00 To POSTES.P_C12
         
        '--- ne prendre que la partie pr�paration ---
        NumCharge = TEtatsPostes(a).NumCharge
                 
        If NumCharge >= CHARGES.C_NUM_MINI And NumCharge <= CHARGES.C_NUM_MAXI Then
        
           Select Case NumPoste
                
               Case POSTES.P_C13
                   '--- postes C13 ---
                   If TEtatsCharges(NumCharge).TGammesAnodisation.ChoixPosteAnodisation = C_C13_IMPOSE Then
                       Inc NbrChargesEnPreparationAvecAnodisationImpose
                   End If
               
               Case POSTES.P_C14
                   '--- postes C14 ---
                   If TEtatsCharges(NumCharge).TGammesAnodisation.ChoixPosteAnodisation = C_C14_IMPOSE Then
                       Inc NbrChargesEnPreparationAvecAnodisationImpose
                   End If
               
               Case POSTES.P_C15
                   '--- postes C15 ---
                   If TEtatsCharges(NumCharge).TGammesAnodisation.ChoixPosteAnodisation = C_C15_IMPOSE Then
                       Inc NbrChargesEnPreparationAvecAnodisationImpose
                   End If
               
               'Case POSTES.P_C16
                   '--- postes C16 ---
                   'If TEtatsCharges(NumCharge).TGammesAnodisation.ChoixPosteAnodisation = C_C16_IMPOSE Then
                       'Inc NbrChargesEnPreparationAvecAnodisationImpose
                   'End If
                
                Case Else
           End Select
        
        End If
    
    Next a
    
    '********************************************************************************************************************
    '                                                                V�rification pour les ponts
    '********************************************************************************************************************
    For a = LBound(TEtatsPonts()) To UBound(TEtatsPonts())
         
        '--- affectation du num�ro de charge ---
        NumCharge = TEtatsPonts(a).NumCharge
                 
        If NumCharge >= CHARGES.C_NUM_MINI And NumCharge <= CHARGES.C_NUM_MAXI Then
             
            '--- contr�ler que l'on n'est jamais passer au Anodisation ---
            If RechercheTempsAuPosteAnodisation(NumCharge, NumPosteAnodisation, DateEntreeAuPosteAnodisation) = 0 Then
            
                Select Case NumPoste
                     
                    Case POSTES.P_C13
                        '--- postes C13 ---
                        If TEtatsCharges(NumCharge).TGammesAnodisation.ChoixPosteAnodisation = C_C13_IMPOSE Then
                            Inc NbrChargesEnPreparationAvecAnodisationImpose
                        End If
                    
                    Case POSTES.P_C14
                        '--- postes C14 ---
                        If TEtatsCharges(NumCharge).TGammesAnodisation.ChoixPosteAnodisation = C_C14_IMPOSE Then
                            Inc NbrChargesEnPreparationAvecAnodisationImpose
                        End If
                    
                    Case POSTES.P_C15
                        '--- postes C15 ---
                        If TEtatsCharges(NumCharge).TGammesAnodisation.ChoixPosteAnodisation = C_C15_IMPOSE Then
                            Inc NbrChargesEnPreparationAvecAnodisationImpose
                        End If
                    
                    'Case POSTES.P_C16
                        '--- postes C16 ---
                        'If TEtatsCharges(NumCharge).TGammesAnodisation.ChoixPosteAnodisation = C_C16_IMPOSE Then
                            'Inc NbrChargesEnPreparationAvecAnodisationImpose
                        'End If
                     
                     Case Else
                End Select
                 
            End If
                 
        End If
                 
    Next a

End Function

'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' R�le      : V�rification qu'une charge est au moins en pr�paration
' Entr�es :
' D�tails  :
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function VerificationChargeEnPreparation() As Boolean

    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next

    '--- d�claration ---
    Dim a As Integer                                                                  'r�serv� pour les boucles FOR ... NEXT
                        
    '--- affectation par d�faut ---
    VerificationChargeEnPreparation = False
                        
    '********************************************************************************************************************
    '                                                V�rification pour les postes de la ligne
    '********************************************************************************************************************
    For a = POSTES.P_C00 To POSTES.P_C34
         
        Select Case a
             
            Case POSTES.P_C00 To POSTES.P_C12
                '--- ne prendre que la partie pr�paration ---
                If TEtatsPostes(a).NumCharge >= CHARGES.C_NUM_MINI Then
                    VerificationChargeEnPreparation = True
                    Exit Function
                 End If
                 
            Case Else
        End Select
                         
    Next a
    
    '********************************************************************************************************************
    '                                                                V�rification pour les ponts
    '********************************************************************************************************************
    For a = LBound(TEtatsPonts()) To UBound(TEtatsPonts())
        If TEtatsPonts(a).NumCharge >= CHARGES.C_NUM_MINI Then
            VerificationChargeEnPreparation = True
            Exit Function
        End If
    Next a

End Function

'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' R�le      : V�rification qu'une charge au chargement � un poste d'anodisation impos� dans sa gamme
' Entr�es :
' D�tails  :
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function VerificationChargeAuChargementAvecAnodisationImpose(ByVal NumPoste As POSTES) As Boolean

    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next

    '--- d�claration ---
    Dim a As Integer                                                                  'r�serv� pour les boucles FOR ... NEXT
    Dim NumCharge As Integer                                                 'indique un num�ro de charge
                        
    '--- affectation par d�faut ---
    VerificationChargeAuChargementAvecAnodisationImpose = False
                        
    '********************************************************************************************************************
    '                                                V�rification pour les postes de la ligne
    '********************************************************************************************************************
    For a = POSTES.P_CHGT_1 To POSTES.P_CHGT_4
         
        '--- affectation du num�ro de charge ---
        NumCharge = TEtatsPostes(a).NumCharge
                 
        If NumCharge >= CHARGES.C_NUM_MINI And NumCharge <= CHARGES.C_NUM_MAXI Then
             
            Select Case NumPoste
                         
                Case POSTES.P_C13
                    '--- postes C13 ---
                    If TEtatsCharges(NumCharge).TGammesAnodisation.ChoixPosteAnodisation = C_C13_IMPOSE Then
                        VerificationChargeAuChargementAvecAnodisationImpose = True
                        Exit For
                    End If
                
                Case POSTES.P_C14
                    '--- postes C14 ---
                    If TEtatsCharges(NumCharge).TGammesAnodisation.ChoixPosteAnodisation = C_C14_IMPOSE Then
                        VerificationChargeAuChargementAvecAnodisationImpose = True
                        Exit For
                    End If
                
                Case POSTES.P_C15
                    '--- postes C15 ---
                    If TEtatsCharges(NumCharge).TGammesAnodisation.ChoixPosteAnodisation = C_C15_IMPOSE Then
                        VerificationChargeAuChargementAvecAnodisationImpose = True
                        Exit For
                    End If
                
                'Case POSTES.P_C16
                    '--- postes C16 ---
                    'If TEtatsCharges(NumCharge).TGammesAnodisation.ChoixPosteAnodisation = C_C65_IMPOSE Then
                    '    VerificationChargeAuChargementAvecAnodisationImpose = True
                    '    Exit For
                    'End If
                 
                 Case Else
            End Select
                 
        End If
                         
    Next a

End Function

'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' R�le      : V�rification de la ligne d'anodisation (occupation des postes) pour autoriser l'entr�e de l'une des charges
'                 pr�sentes au chargement, ceci afin d'�viter les conflits de postes et de lib�ration du pont
'                 (pont libre = possibilit� de mouvements) dans la partie pr�paration de la ligne
'                 d�s qu'une charge peut �tre rentr� en ligne cette fonction modifie la variable
'                 ProchainNumPosteChargement du tableau du moteur d'inf�rence
' Entr�es :
' D�tails  :
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Sub VerificationLignePourEntreeCharge()

    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next
    
    '--- constantes priv�es ---
    Const NBR_CHARGES_MAXI_EN_PREPARATION As Integer = 4
    
    '--- d�claration ---
    Static MemAffichageMessages As Boolean                                    'm�moire d'affichage des messages
    
    Dim SortieModule As Boolean                                                         'indique qu'il faut sortir de ce module
    Dim ChargeEnZonePreparation As Boolean                                    'indique qu'une charge est en zone de pr�paration
    
    Dim EntreePossibleChargeAvecAnodisationAutomatique As Boolean 'indique la possibilit� d'entr�e une charge avec anodisation sur automatique
    
    Dim PassageZoneSpectrocoloration As Boolean                            'indique le passage dans la zone de spectrocoloration
    Dim PassageZoneOr As Boolean                                                     'indique le passage dans la zone d'or
    Dim PassageZoneNoir As Boolean                                                  'indique le passage dans la zone de noir
    
    Dim a As Integer                                                                               'r�serv� pour les boucles FOR ... NEXT
    Dim NumCharge As Integer                                                              'indique un num�ro de charge
    
    Dim NumChargePosteChargementPourC13 As Integer                               'indique un num�ro de charge au poste de chargement pour la cuve C13
    Dim NumChargePosteChargementPourC14 As Integer                               'indique un num�ro de charge au poste de chargement pour la cuve C14
    Dim NumChargePosteChargementPourC15 As Integer                               'indique un num�ro de charge au poste de chargement pour la cuve C15
    Dim NumChargePosteChargementPourC16 As Integer                               'indique un num�ro de charge au poste de chargement pour la cuve C16
    Dim NumChargePosteChargementSiAnodisationAutomatique As Integer  'indique un num�ro de charge au poste de chargement si le poste d'anodisation est automatique
    
    Dim NumChargeALancerPourC13 As Integer                                     'indique le num�ro de charge � lancer pour C13
    Dim NumChargeALancerPourC14 As Integer                                     'indique le num�ro de charge � lancer pour C14
    Dim NumChargeALancerPourC15 As Integer                                     'indique le num�ro de charge � lancer pour C15
    Dim NumChargeALancerPourC16 As Integer                                     'indique le num�ro de charge � lancer pour C16
    
    Dim CptPostes As Integer                                                                   'compteur des postes pour pointer dans le tableau
                                                                                                                'de l'ordre de sortie des charges
    Dim PtrZoneGammeAnodisation As Integer                                       'pointeur de la zone de la gamme d'anodisation

    Dim NbrChargesEnPreparation As Integer                                         'indique le nombre de charges en pr�paration
    Static MemNbrChargesEnPreparation As Integer                               'm�moire du nombre de charges en pr�paration

    Dim NbrChargesEnDegraissageSatinage As Integer                         'nombre de charges dans la zone de d�graissage / satinage
    Static MemNbrChargesEnDegraissageSatinage As Integer              'm�moire du nombre de charges dans la zone de d�graissage / satinage
    
    Dim NbrChargesEnBrillantage As Integer                                          'nombre de charges dans la zone de brillantage
    Static MemNbrChargesEnBrillantage As Integer                                'm�moire du nombre de charges dans la zone de brillantage
    
    Dim TempsMouvementsAvantPostePrincipalSecondes As Long      'temps des mouvements avant le poste principal en secondes
    Dim TempsAvantPostePrincipalAvecPontsSecondes As Long          'temps avant le poste principal avec les ponts en secondes
    Dim TempsPostePrincipalAvecPontsSecondes As Long                   'temps au poste principal avec les ponts en secondes
    Dim TempsMouvementsApresPostePrincipalSecondes As Long      'temps des mouvements apr�s le poste principal en secondes
    Dim TempsApresPostePrincipalAvecPontsSecondes As Long          'temps apr�s le poste principal avec les ponts en secondes
    Dim TempsTotalPostesAvecPontsSecondes As Long                       'temps total des postes avec les ponts en secondes
    Dim TempsTotalEgouttagesAvecPontsSecondes As Long                'temps total des �gouttages avec les ponts en secondes
    Dim TempsTotalMouvementsSecondes As Long                              'temps total des mouvements en secondes
    Dim TempsTotalGammeAvecPontsSecondes As Long                      'temps total de la gamme avec les ponts en secondes

    Dim TGammesAnodisation As EnrGammesAnodisation                   'repr�sente une gamme d'anodisation

                  '********** CORRESPOND AUX DETAILS DES GAMMES d'anodisation DES CHARGES **********

    Dim NumPosteReel As Integer                                                         'N� de poste r�el utilis� dans la zone (cas des postes multiples)
                                                                                                              
    Dim DecompteDuTempsAuPosteReelSecondes As String              'repr�sente la diff�rence entre le temps th�orique
                                                                                                              'au poste et le temps r�el pass� dans le poste
                                                                                                              'un nombre n�gatif apparait si la charge est rest� plus
                                                                                                              'longtemps dans le poste que le temps th�orique pr�vu
                                                                                                              'ATTENTION variable du type String volontairement
                                                                                                              'Si "" alors il n'y a pas eu de temps de d�compter
    Dim FicheVideInformationsPostesAnodisation As VarInformationsPostesAnodisation 'fiche vide des informations sur les postes d'anodisation

    '----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

    '********************************************************************************************************************
    '  Sortie directe de la routine si une charge doit d�j� rentrer en ligne (pointeur de zone de la gamme est � 1)
    '********************************************************************************************************************
    For a = POSTES.P_CHGT_1 To POSTES.P_CHGT_4
        NumCharge = TEtatsPostes(a).NumCharge
        If NumCharge >= CHARGES.C_NUM_MINI And NumCharge <= CHARGES.C_NUM_MAXI Then
            If TEtatsCharges(NumCharge).PtrZoneGammeAnodisation = 1 Then
                SortieModule = True
                Exit For
            End If
        End If
    Next a

    '--- sortie du module car une charge doit d�j� rentrer en ligne (pointeur de zone de la gamme est � 1) ---
    If SortieModule = True Then
        
        '--- affichage des informations sur les entr�es des charges avec un anti-rebond ---
        If MemAffichageMessages = False Then
            AfficheRenseignementsEntreesCharges VERT_4, _
                                                                             "Plus de calculs pour les entr�es - La charge " & NumCharge & " est d�j� s�lectionn�e" & _
                                                                             vbCrLf
            MemAffichageMessages = True
        End If
        
        '--- sortie de la routine ---
        Exit Sub
    
    Else
    
        '--- RAZ de la m�moire d'affichage des messages ---
        MemAffichageMessages = False
    
    End If

    '----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    
    With TMoteurInference
    
        '********************************************************************************************************************
        '                               Remplissage du tableau des informations sur les postes d'anodisation
        '********************************************************************************************************************
        For a = LBound(.TInformationsPostesAnodisation()) To UBound(.TInformationsPostesAnodisation())

            Select Case a
            
                Case POSTES.P_C13, POSTES.P_C14, POSTES.P_C15 ', POSTES.P_C16
                    '--- ne prendre que les postes d'anodisation ---
                    With TEtatsPostes(a)

                        '--- affectation du n� de charge ---
                        NumCharge = .NumCharge

                        If NumCharge >= CHARGES.C_NUM_MINI And NumCharge <= CHARGES.C_NUM_MAXI Then

                            '--- affectation du num�ro de charge et de la condamnation ---
                            With TMoteurInference.TInformationsPostesAnodisation(a)
                                .NumCharge = NumCharge
                                .Condamnation = TEtatsPostes(a).Condamnation
                            End With
                            
                            '--- affectation du pointeur de la zone de la gamme d'anodisation ---
                            PtrZoneGammeAnodisation = TEtatsCharges(NumCharge).PtrZoneGammeAnodisation
        
                            If PtrZoneGammeAnodisation > 0 Then
        
                                With TEtatsCharges(NumCharge).TGammesAnodisation.TDetailsGammesAnodisation(PtrZoneGammeAnodisation)
        
                                    '--- affectation du n� du poste r�el ---
                                    NumPosteReel = .NumPosteReel
        
                                    If a = .NumPosteReel Then               'v�rifier la concordance entre le poste scrut� et le poste r�el
        
                                        '--- affectation d�compte du temps au poste ---
                                        DecompteDuTempsAuPosteReelSecondes = .DecompteDuTempsAuPosteReelSecondes
        
                                        '--- remplir le tableau avec le n� de charge ainsi que le temps de d�compte de celui-ci ---
                                        If IsNumeric(DecompteDuTempsAuPosteReelSecondes) = True Then
        
                                            '--- compl�ment de la fiche ---
                                            With TMoteurInference.TInformationsPostesAnodisation(a)
                                                .DecompteDuTempsAuPosteReelSecondes = DecompteDuTempsAuPosteReelSecondes
                                            End With
        
                                        End If
        
                                    End If
        
                                End With
                        
                            End If

                        Else
        
                            '--- effacement de la fiche ---
                            TMoteurInference.TInformationsPostesAnodisation(a) = FicheVideInformationsPostesAnodisation
        
                        End If
        
                    End With

                Case Else
            End Select

        Next a
    
        '----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
        
        '**************************************************************************************************************
        '                                        Recherche du nombre de charges en pr�paration
        '**************************************************************************************************************
        NbrChargesEnPreparation = RechercheNbrChargesEnPreparation()
        
        If MemNbrChargesEnPreparation <> NbrChargesEnPreparation Then
    
            '--- affichage avec anti-rebond ---
            AfficheRenseignementsEntreesCharges ROUGE_3, "Nombre de charges en pr�paration = " & NbrChargesEnPreparation & vbCrLf
        
            '--- affectation de la m�moire du nombre de charges en pr�paration ---
            MemNbrChargesEnPreparation = NbrChargesEnPreparation
        
        End If
        
        '----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
        
        '**************************************************************************************************************
        '                              Recherche du nombre de charges en d�graissage / satinage
        '**************************************************************************************************************
        NbrChargesEnDegraissageSatinage = RechercheNbrChargesEnDegraissageSatinage()
        
        If MemNbrChargesEnDegraissageSatinage <> NbrChargesEnDegraissageSatinage Then
    
            '--- affichage avec anti-rebond ---
            AfficheRenseignementsEntreesCharges ROUGE_3, "Nombre de charges en d�graissage / satinage = " & NbrChargesEnDegraissageSatinage & vbCrLf
        
            '--- affectation de la m�moire du nombre de charges en d�graissage / satinage ---
            MemNbrChargesEnDegraissageSatinage = NbrChargesEnDegraissageSatinage
        
        End If
        
        '----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
        
        '**************************************************************************************************************
        '                   Recherche du  nombre de charges dans la zone de brillantage
        '**************************************************************************************************************
        NbrChargesEnBrillantage = RechercheNbrChargesEnBrillantage()
        
        If MemNbrChargesEnBrillantage <> NbrChargesEnBrillantage Then
    
            '--- affichage avec anti-rebond ---
            AfficheRenseignementsEntreesCharges ROUGE_3, "Nombre de charges en brillantage = " & NbrChargesEnBrillantage & vbCrLf
        
            '--- affectation de la m�moire du nombre de charges en brillantage ---
            MemNbrChargesEnBrillantage = NbrChargesEnBrillantage
        
        End If
        
        '----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
        
        '**************************************************************************************************************
        '                                              Analyse avec anodisation C13 IMPOSE dans la gamme
        '**************************************************************************************************************
        If .ProchainNumPosteChargementSiAnodisationC13Impose > 0 Then  'ne traiter la s�quence qu'avec la pr�sence
            
            With .TInformationsPostesAnodisation(POSTES.P_C13)
                
                '--- affectation du num�ro de charge au poste de chargement pour C13 ---
                NumChargePosteChargementPourC13 = TEtatsPostes(TMoteurInference.ProchainNumPosteChargementSiAnodisationC13Impose).NumCharge
                
                If .Condamnation = True Then                                                'le poste est condamn� il ne faut pas
                                                                                                                'traiter la s�quence
                
                Else

                    '--- le poste d'anodisation est vide il faut v�rifier si une charge est d�j� dans la zone de pr�paration ---
                    If .NumCharge = 0 Then
                        
                        '--- affichage des informations sur les entr�es des charges ---
                        AfficheRenseignementsEntreesCharges VERT_4, "Pas de charge en C13" & vbCrLf
                        
                        Select Case NbrChargesEnPreparationAvecAnodisationImpose(POSTES.P_C13)
        
                            '----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
                            
                            Case 0
                                '--- PAS DE CHARGE EN PREPARATION ---
                                '--- affichage des informations sur les entr�es des charges ---
                                AfficheRenseignementsEntreesCharges VERT_4, "Plus de charge avec C13 IMPOSE en ZONE de pr�paration" & vbCrLf
                                
                                If NbrChargesEnPreparation < NBR_CHARGES_MAXI_EN_PREPARATION Then ' Or (NbrChargesEnPreparation = 1 And NbrChargesEnBrillantage = 1) Then
                                
                                    '--- affectation du num�ro de charge � lancer pour C13 ---
                                    NumChargeALancerPourC13 = NumChargePosteChargementPourC13
                                
                                    '--- affichage des informations sur les entr�es des charges ---
                                    AfficheRenseignementsEntreesCharges BLEU_3, "Plus de charge en pr�paration - N� de charge � lancer Anodisation VIDE = " & NumChargeALancerPourC13 & vbCrLf
                                
                                End If
                        
                            '----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
                            
                            Case 1
                                '--- UNE CHARGE EN PREPARATION ---
                                '--- affichage des informations sur les entr�es des charges ---
                                AfficheRenseignementsEntreesCharges VERT_4, "Plus de charge avec C13 IMPOSE en ZONE de pr�paration" & vbCrLf
                            
                                If NbrChargesEnBrillantage = 1 Then
                    
                                    '--- affectation de la gamme de la charge se trouvant au poste de chargement ---
                                    TGammesAnodisation = TEtatsCharges(NumChargePosteChargementPourC13).TGammesAnodisation
                    
                                    '--- calcul les temps principaux d'une gamme d'anodisation AVEC LES TEMPS DE MOUVEMENTS DES PONTS ---
                                    CalculTempsGammeAnodisationAvecPonts TGammesAnodisation, _
                                                                                                           TempsMouvementsAvantPostePrincipalSecondes, _
                                                                                                           TempsAvantPostePrincipalAvecPontsSecondes, _
                                                                                                           TempsPostePrincipalAvecPontsSecondes, _
                                                                                                           TempsMouvementsApresPostePrincipalSecondes, _
                                                                                                           TempsApresPostePrincipalAvecPontsSecondes, _
                                                                                                           TempsTotalPostesAvecPontsSecondes, _
                                                                                                           TempsTotalEgouttagesAvecPontsSecondes, _
                                                                                                           TempsTotalMouvementsSecondes, _
                                                                                                           TempsTotalGammeAvecPontsSecondes, _
                                                                                                           PassageZoneSpectrocoloration, _
                                                                                                           PassageZoneOr, _
                                                                                                           PassageZoneNoir

                                    '--- analyse du temps restant dans la pr�paration pour trouver la meilleure entr�e au chargement ---
                                    If IsNumeric(.DecompteDuTempsAuPosteReelSecondes) = True Then
                                        If CLng(.DecompteDuTempsAuPosteReelSecondes) < TempsAvantPostePrincipalAvecPontsSecondes Then
                                
                                            '--- affectation du num�ro de charge � lancer pour C13 ---
                                            NumChargeALancerPourC13 = NumChargePosteChargementPourC13
                                    
                                            '--- affichage des informations sur les entr�es des charges ---
                                            AfficheRenseignementsEntreesCharges BLEU_3, "Plus de charge en pr�paration - N� de charge � lancer Anodisation PLEIN = " & NumChargeALancerPourC13 & vbCrLf
                                
                                        End If
                                    End If
                        
                                End If
                        
                            Case Else
                        End Select
                   
                    Else
                   
                        '--- affichage des informations sur les entr�es des charges ---
                        AfficheRenseignementsEntreesCharges VERT_4, "Charge " & .NumCharge & " en C13" & vbCrLf
                        
                        If VerificationChargeEnPreparationAvecAnodisationImpose(POSTES.P_C13) = False Then
                        
                            '--- affichage des informations sur les entr�es des charges ---
                            AfficheRenseignementsEntreesCharges VERT_4, "Plus de charge avec C13 IMPOSE en ZONE de pr�paration" & vbCrLf
                            
                            If NbrChargesEnPreparation < NBR_CHARGES_MAXI_EN_PREPARATION Then ' Or (NbrChargesEnPreparation = 1 And NbrChargesEnBrillantage = 1) Then
                        
                                '--- affectation de la gamme de la charge se trouvant au poste de chargement ---
                                TGammesAnodisation = TEtatsCharges(NumChargePosteChargementPourC13).TGammesAnodisation
                        
                                '--- calcul les temps principaux d'une gamme d'anodisation AVEC LES TEMPS DE MOUVEMENTS DES PONTS ---
                                CalculTempsGammeAnodisationAvecPonts TGammesAnodisation, _
                                                                                                       TempsMouvementsAvantPostePrincipalSecondes, _
                                                                                                       TempsAvantPostePrincipalAvecPontsSecondes, _
                                                                                                       TempsPostePrincipalAvecPontsSecondes, _
                                                                                                       TempsMouvementsApresPostePrincipalSecondes, _
                                                                                                       TempsApresPostePrincipalAvecPontsSecondes, _
                                                                                                       TempsTotalPostesAvecPontsSecondes, _
                                                                                                       TempsTotalEgouttagesAvecPontsSecondes, _
                                                                                                       TempsTotalMouvementsSecondes, _
                                                                                                       TempsTotalGammeAvecPontsSecondes, _
                                                                                                       PassageZoneSpectrocoloration, _
                                                                                                       PassageZoneOr, _
                                                                                                       PassageZoneNoir
                        
                        
                                '--- analyse du temps restant dans l'anodisation pour trouver la meilleure entr�e au chargement ---
                                If IsNumeric(.DecompteDuTempsAuPosteReelSecondes) = True Then
                                    If CLng(.DecompteDuTempsAuPosteReelSecondes) < TempsAvantPostePrincipalAvecPontsSecondes Then
                                    
                                        '--- affectation du num�ro de charge � lancer pour C13 ---
                                        NumChargeALancerPourC13 = NumChargePosteChargementPourC13
                                        
                                        '--- affichage des informations sur les entr�es des charges ---
                                        AfficheRenseignementsEntreesCharges BLEU_3, "Plus de charge en pr�paration - N� de charge � lancer Anodisation PLEIN = " & NumChargeALancerPourC13 & vbCrLf
                                    
                                    End If
                                End If
                                                    
                            End If
                        
                        End If
                        
                    End If

                End If
            
            End With
        
        End If

        '----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

        '**************************************************************************************************************
        '                                              Analyse avec Anodisation C14 IMPOSE dans la gamme
        '**************************************************************************************************************
        If .ProchainNumPosteChargementSiAnodisationC14Impose > 0 Then  'ne traiter la s�quence qu'avec la pr�sence
            
            With .TInformationsPostesAnodisation(POSTES.P_C14)
                
                '--- affectation du num�ro de charge au poste de chargement pour C14 ---
                NumChargePosteChargementPourC14 = TEtatsPostes(TMoteurInference.ProchainNumPosteChargementSiAnodisationC14Impose).NumCharge
                
                If .Condamnation = True Then                                                'le poste est condamn� il ne faut pas
                                                                                                                'traiter la s�quence
                
                Else

                    '--- le poste d'anodisation est vide il faut v�rifier si une charge est d�j� dans la zone de pr�paration ---
                    If .NumCharge = 0 Then
                        
                        '--- affichage des informations sur les entr�es des charges ---
                        AfficheRenseignementsEntreesCharges VERT_4, "Pas de charge en C14" & vbCrLf
                        
                        Select Case NbrChargesEnPreparationAvecAnodisationImpose(POSTES.P_C14)
                        
                            '----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
                            
                            Case 0
                                '--- PAS DE CHARGE EN PREPARATION ---
                                '--- affichage des informations sur les entr�es des charges ---
                                AfficheRenseignementsEntreesCharges VERT_4, "Plus de charge avec C14 IMPOSE en ZONE de pr�paration" & vbCrLf
                                
                                If NbrChargesEnPreparation < NBR_CHARGES_MAXI_EN_PREPARATION Then ' Or (NbrChargesEnPreparation = 1 And NbrChargesEnBrillantage = 1) Then
                                
                                    '--- affectation du num�ro de charge � lancer pour C14 ---
                                    NumChargeALancerPourC14 = NumChargePosteChargementPourC14
                                
                                    '--- affichage des informations sur les entr�es des charges ---
                                    AfficheRenseignementsEntreesCharges BLEU_3, "Plus de charge en pr�paration - N� de charge � lancer Anodisation VIDE = " & NumChargeALancerPourC14 & vbCrLf
                                
                                End If
                        
                            '----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
                            
                            Case 1
                                '--- UNE CHARGE EN PREPARATION ---
                                '--- affichage des informations sur les entr�es des charges ---
                                AfficheRenseignementsEntreesCharges VERT_4, "Plus de charge avec C14 IMPOSE en ZONE de pr�paration" & vbCrLf
                            
                                If NbrChargesEnBrillantage = 1 Then
                    
                                    '--- affectation de la gamme de la charge se trouvant au poste de chargement ---
                                    TGammesAnodisation = TEtatsCharges(NumChargePosteChargementPourC14).TGammesAnodisation
                    
                                    '--- calcul les temps principaux d'une gamme d'anodisation AVEC LES TEMPS DE MOUVEMENTS DES PONTS ---
                                    CalculTempsGammeAnodisationAvecPonts TGammesAnodisation, _
                                                                                                           TempsMouvementsAvantPostePrincipalSecondes, _
                                                                                                           TempsAvantPostePrincipalAvecPontsSecondes, _
                                                                                                           TempsPostePrincipalAvecPontsSecondes, _
                                                                                                           TempsMouvementsApresPostePrincipalSecondes, _
                                                                                                           TempsApresPostePrincipalAvecPontsSecondes, _
                                                                                                           TempsTotalPostesAvecPontsSecondes, _
                                                                                                           TempsTotalEgouttagesAvecPontsSecondes, _
                                                                                                           TempsTotalMouvementsSecondes, _
                                                                                                           TempsTotalGammeAvecPontsSecondes, _
                                                                                                           PassageZoneSpectrocoloration, _
                                                                                                           PassageZoneOr, _
                                                                                                           PassageZoneNoir
                    
                                    '--- analyse du temps restant dans la pr�paration pour trouver la meilleure entr�e au chargement ---
                                    If IsNumeric(.DecompteDuTempsAuPosteReelSecondes) = True Then
                                        If CLng(.DecompteDuTempsAuPosteReelSecondes) < TempsAvantPostePrincipalAvecPontsSecondes Then
                                
                                            '--- affectation du num�ro de charge � lancer pour C14 ---
                                            NumChargeALancerPourC14 = NumChargePosteChargementPourC14
                                    
                                            '--- affichage des informations sur les entr�es des charges ---
                                            AfficheRenseignementsEntreesCharges BLEU_3, "Plus de charge en pr�paration - N� de charge � lancer Anodisation PLEIN = " & NumChargeALancerPourC14 & vbCrLf
                                
                                        End If
                                    End If
                        
                                End If
                        
                            Case Else
                        End Select
                   
                   Else
                   
                        '--- affichage des informations sur les entr�es des charges ---
                        AfficheRenseignementsEntreesCharges VERT_4, "Charge " & .NumCharge & " en C14" & vbCrLf
                        
                        If VerificationChargeEnPreparationAvecAnodisationImpose(POSTES.P_C14) = False Then
                        
                            '--- affichage des informations sur les entr�es des charges ---
                            AfficheRenseignementsEntreesCharges VERT_4, "Plus de charge avec C14 IMPOSE en ZONE de pr�paration" & vbCrLf
                            
                            If NbrChargesEnPreparation < NBR_CHARGES_MAXI_EN_PREPARATION Then ' Or (NbrChargesEnPreparation = 1 And NbrChargesEnBrillantage = 1) Then
                        
                                '--- affectation de la gamme de la charge se trouvant au poste de chargement ---
                                TGammesAnodisation = TEtatsCharges(NumChargePosteChargementPourC14).TGammesAnodisation
                        
                                '--- calcul les temps principaux d'une gamme d'anodisation AVEC LES TEMPS DE MOUVEMENTS DES PONTS ---
                                CalculTempsGammeAnodisationAvecPonts TGammesAnodisation, _
                                                                                                       TempsMouvementsAvantPostePrincipalSecondes, _
                                                                                                       TempsAvantPostePrincipalAvecPontsSecondes, _
                                                                                                       TempsPostePrincipalAvecPontsSecondes, _
                                                                                                       TempsMouvementsApresPostePrincipalSecondes, _
                                                                                                       TempsApresPostePrincipalAvecPontsSecondes, _
                                                                                                       TempsTotalPostesAvecPontsSecondes, _
                                                                                                       TempsTotalEgouttagesAvecPontsSecondes, _
                                                                                                       TempsTotalMouvementsSecondes, _
                                                                                                       TempsTotalGammeAvecPontsSecondes, _
                                                                                                       PassageZoneSpectrocoloration, _
                                                                                                       PassageZoneOr, _
                                                                                                       PassageZoneNoir
                        
                                '--- analyse du temps restant dans l'anodisation pour trouver la meilleure entr�e au chargement ---
                                If IsNumeric(.DecompteDuTempsAuPosteReelSecondes) = True Then
                                    If CLng(.DecompteDuTempsAuPosteReelSecondes) < TempsAvantPostePrincipalAvecPontsSecondes Then
                                    
                                        '--- affectation du num�ro de charge � lancer pour C14 ---
                                        NumChargeALancerPourC14 = NumChargePosteChargementPourC14
                                        
                                        '--- affichage des informations sur les entr�es des charges ---
                                        AfficheRenseignementsEntreesCharges BLEU_3, "Plus de charge en pr�paration - N� de charge � lancer Anodisation PLEIN = " & NumChargeALancerPourC14 & vbCrLf
                                    
                                    End If
                                End If
                                                    
                            End If
                        
                        End If
                        
                    End If

                End If
            
            End With
        
        End If
        
        '----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

        '**************************************************************************************************************
        '                                              Analyse avec Anodisation C15 IMPOSE dans la gamme
        '**************************************************************************************************************
        If .ProchainNumPosteChargementSiAnodisationC15Impose > 0 Then  'ne traiter la s�quence qu'avec la pr�sence
            
            With .TInformationsPostesAnodisation(POSTES.P_C15)
                
                '--- affectation du num�ro de charge au poste de chargement pour C15 ---
                NumChargePosteChargementPourC15 = TEtatsPostes(TMoteurInference.ProchainNumPosteChargementSiAnodisationC15Impose).NumCharge
                
                If .Condamnation = True Then                                                'le poste est condamn� il ne faut pas
                                                                                                                'traiter la s�quence
                
                Else

                    '--- le poste d'anodisation est vide il faut v�rifier si une charge est d�j� dans la zone de pr�paration ---
                    If .NumCharge = 0 Then
                        
                        '--- affichage des informations sur les entr�es des charges ---
                        AfficheRenseignementsEntreesCharges VERT_4, "Pas de charge en C15" & vbCrLf
                        
                        Select Case NbrChargesEnPreparationAvecAnodisationImpose(POSTES.P_C15)
                        
                            '----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
                            
                            Case 0
                                '--- PAS DE CHARGE EN PREPARATION ---
                                '--- affichage des informations sur les entr�es des charges ---
                                AfficheRenseignementsEntreesCharges VERT_4, "Plus de charge avec C15 IMPOSE en ZONE de pr�paration" & vbCrLf
                                
                                If NbrChargesEnPreparation < NBR_CHARGES_MAXI_EN_PREPARATION Then ' Or (NbrChargesEnPreparation = 1 And NbrChargesEnBrillantage = 1) Then
                                
                                    '--- affectation du num�ro de charge � lancer pour C15 ---
                                    NumChargeALancerPourC15 = NumChargePosteChargementPourC15
                                
                                    '--- affichage des informations sur les entr�es des charges ---
                                    AfficheRenseignementsEntreesCharges BLEU_3, "Plus de charge en pr�paration - N� de charge � lancer Anodisation VIDE = " & NumChargeALancerPourC15 & vbCrLf
                                
                                End If
                        
                            '----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
                            
                            Case 1
                                '--- UNE CHARGE EN PREPARATION ---
                                '--- affichage des informations sur les entr�es des charges ---
                                AfficheRenseignementsEntreesCharges VERT_4, "Plus de charge avec C15 IMPOSE en ZONE de pr�paration" & vbCrLf
                            
                                If NbrChargesEnBrillantage = 1 Then
                    
                                    '--- affectation de la gamme de la charge se trouvant au poste de chargement ---
                                    TGammesAnodisation = TEtatsCharges(NumChargePosteChargementPourC15).TGammesAnodisation
                    
                                    '--- calcul les temps principaux d'une gamme d'anodisation AVEC LES TEMPS DE MOUVEMENTS DES PONTS ---
                                    CalculTempsGammeAnodisationAvecPonts TGammesAnodisation, _
                                                                                                           TempsMouvementsAvantPostePrincipalSecondes, _
                                                                                                           TempsAvantPostePrincipalAvecPontsSecondes, _
                                                                                                           TempsPostePrincipalAvecPontsSecondes, _
                                                                                                           TempsMouvementsApresPostePrincipalSecondes, _
                                                                                                           TempsApresPostePrincipalAvecPontsSecondes, _
                                                                                                           TempsTotalPostesAvecPontsSecondes, _
                                                                                                           TempsTotalEgouttagesAvecPontsSecondes, _
                                                                                                           TempsTotalMouvementsSecondes, _
                                                                                                           TempsTotalGammeAvecPontsSecondes, _
                                                                                                           PassageZoneSpectrocoloration, _
                                                                                                           PassageZoneOr, _
                                                                                                           PassageZoneNoir
                    
                                    '--- analyse du temps restant dans la pr�paration pour trouver la meilleure entr�e au chargement ---
                                    If IsNumeric(.DecompteDuTempsAuPosteReelSecondes) = True Then
                                        If CLng(.DecompteDuTempsAuPosteReelSecondes) < TempsAvantPostePrincipalAvecPontsSecondes Then
                                
                                            '--- affectation du num�ro de charge � lancer pour C15 ---
                                            NumChargeALancerPourC15 = NumChargePosteChargementPourC15
                                    
                                            '--- affichage des informations sur les entr�es des charges ---
                                            AfficheRenseignementsEntreesCharges BLEU_3, "Plus de charge en pr�paration - N� de charge � lancer Anodisation PLEIN = " & NumChargeALancerPourC15 & vbCrLf
                                
                                        End If
                                    End If
                        
                                End If
                        
                            Case Else
                        End Select
                   
                   Else
                   
                        '--- affichage des informations sur les entr�es des charges ---
                        AfficheRenseignementsEntreesCharges VERT_4, "Charge " & .NumCharge & " en C15" & vbCrLf
                        
                        If VerificationChargeEnPreparationAvecAnodisationImpose(POSTES.P_C15) = False Then
                        
                            '--- affichage des informations sur les entr�es des charges ---
                            AfficheRenseignementsEntreesCharges VERT_4, "Plus de charge avec C15 IMPOSE en ZONE de pr�paration" & vbCrLf
                            
                            If NbrChargesEnPreparation < NBR_CHARGES_MAXI_EN_PREPARATION Then ' Or (NbrChargesEnPreparation = 1 And NbrChargesEnBrillantage = 1) Then
                        
                                '--- affectation de la gamme de la charge se trouvant au poste de chargement ---
                                TGammesAnodisation = TEtatsCharges(NumChargePosteChargementPourC15).TGammesAnodisation
                        
                                '--- calcul les temps principaux d'une gamme d'anodisation AVEC LES TEMPS DE MOUVEMENTS DES PONTS ---
                                CalculTempsGammeAnodisationAvecPonts TGammesAnodisation, _
                                                                                                       TempsMouvementsAvantPostePrincipalSecondes, _
                                                                                                       TempsAvantPostePrincipalAvecPontsSecondes, _
                                                                                                       TempsPostePrincipalAvecPontsSecondes, _
                                                                                                       TempsMouvementsApresPostePrincipalSecondes, _
                                                                                                       TempsApresPostePrincipalAvecPontsSecondes, _
                                                                                                       TempsTotalPostesAvecPontsSecondes, _
                                                                                                       TempsTotalEgouttagesAvecPontsSecondes, _
                                                                                                       TempsTotalMouvementsSecondes, _
                                                                                                       TempsTotalGammeAvecPontsSecondes, _
                                                                                                       PassageZoneSpectrocoloration, _
                                                                                                       PassageZoneOr, _
                                                                                                       PassageZoneNoir

                                '--- analyse du temps restant dans l'anodisation pour trouver la meilleure entr�e au chargement ---
                                If IsNumeric(.DecompteDuTempsAuPosteReelSecondes) = True Then
                                    If CLng(.DecompteDuTempsAuPosteReelSecondes) < TempsAvantPostePrincipalAvecPontsSecondes Then
                                    
                                        '--- affectation du num�ro de charge � lancer pour C15 ---
                                        NumChargeALancerPourC15 = NumChargePosteChargementPourC15
                                        
                                        '--- affichage des informations sur les entr�es des charges ---
                                        AfficheRenseignementsEntreesCharges BLEU_3, "Plus de charge en pr�paration - N� de charge � lancer Anodisation PLEIN = " & NumChargeALancerPourC15 & vbCrLf
                                    
                                    End If
                                End If
                                                    
                            End If
                        
                        End If
                        
                    End If

                End If
            
            End With
        
        End If
        
        '----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

        '**************************************************************************************************************
        ' Analyse avec anodisation AUTOMATIQUE dans la gamme et AU MOINS UN POSTE d'anodisation VIDE
        '**************************************************************************************************************
        If .ProchainNumPosteChargementSiAnodisationAutomatique > 0 Then
        
            '--- affectation du num�ro de charge au poste de chargement si le poste d'anodisation est automatique ---
            NumChargePosteChargementSiAnodisationAutomatique = TEtatsPostes(TMoteurInference.ProchainNumPosteChargementSiAnodisationAutomatique).NumCharge
        
            If NbrChargesEnPreparation < NBR_CHARGES_MAXI_EN_PREPARATION Then ' Or (NbrChargesEnPreparation = 1 And NbrChargesEnBrillantage = 1) Then
                
                For a = LBound(.TInformationsPostesAnodisation()) To UBound(.TInformationsPostesAnodisation())

                    Select Case a
            
                        Case POSTES.P_C13, POSTES.P_C14, POSTES.P_C15 ', POSTES.P_C16
                            '--- ne prendre que les postes d'anodisation ---
                            If .TInformationsPostesAnodisation(a).Condamnation = False Then
                                
                                If VerificationChargeEnPreparationAvecAnodisationImpose(a) = False Then
                                    
                                    If VerificationChargeAuChargementAvecAnodisationImpose(a) = False Then
                                    
                                        If TEtatsPostes(a).NumCharge = 0 Then
                                                    
                                            '--- indiquer la possibilit� d'entr�e une charge avec anodisation sur automatique ---
                                            EntreePossibleChargeAvecAnodisationAutomatique = True
                                            
                                            '--- sortie directe apr�s l'affectation du choix du poste d'anodisation ---
                                            Exit For
                                        
                                        End If
                                
                                    End If
                                
                                End If
                            
                            End If
                
                        Case Else
                    End Select
                
                Next a
                
            End If
        
        End If
        
        '----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
        
        '**************************************************************************************************************
        '    Analyse avec anodisation AUTOMATIQUE dans la gamme et PLUS DE POSTE d'anodisation VIDE
        '**************************************************************************************************************
        'If .ProchainNumPosteChargementSiAnodisationAutomatique > 0 Then
        
            '--- affectation du num�ro de charge au poste de chargement si le poste d'anodisation est automatique ---
            'NumChargePosteChargementSiAnodisationAutomatique = TEtatsPostes(TMoteurInference.ProchainNumPosteChargementSiAnodisationAutomatique).NumCharge
        
            'If NbrChargesEnPreparation < NBR_CHARGES_MAXI_EN_PREPARATION Then ' Or (NbrChargesEnPreparation = 1 And NbrChargesEnBrillantage = 1) Then
                
                'For a = LBound(.TInformationsPostesAnodisation()) To UBound(.TInformationsPostesAnodisation())

                    'Select Case a
            
                        'Case POSTES.P_C13, POSTES.P_C14, POSTES.P_C15 ', POSTES.P_C16
                            '--- ne prendre que les postes d'anodisation ---
                            'If .TInformationsPostesAnodisation(a).Condamnation = False Then
                                
                                'If VerificationChargeEnPreparationAvecAnodisationImpose(a) = False Then
                                    
                                    'If VerificationChargeAuChargementAvecAnodisationImpose(a) = False Then
                                    
                                        'If TEtatsPostes(a).NumCharge <> 0 Then
                                            
                                                    
                                            '--- sortie directe apr�s l'affectation du choix du poste d'anodisation ---
                                            'Exit For
                                        
                                            '--- recherche du temps le plus court dans l'anodisation ---
                                            'If a = POSTES.P_C13 Then
                                            
                                                '--- poste d'anodisation C13 ---
                                            
                                            'Else
                                            
                                            'Select Case a
                                            '    Case POSTES.P_C13
                                            '        'if TMoteurInference.TInformationsPostesAnodisation(a).DecompteDuTempsAuPosteReelSecondes
                                            '        TEtatsCharges(NumChargePosteSiAnodisationAutomatique).TGammesAnodisation.ChoixPosteAnodisation = C_C13_IMPOSE
                                            '    Case POSTES.P_C14
                                            '        TEtatsCharges(NumChargePosteSiAnodisationAutomatique).TGammesAnodisation.ChoixPosteAnodisation = C_C14_IMPOSE
                                            '    Case POSTES.P_C15
                                            '        TEtatsCharges(NumChargePosteSiAnodisationAutomatique).TGammesAnodisation.ChoixPosteAnodisation = C_C15_IMPOSE
                                            '    Case Else
                                            'End Select
                                
                                        'End If
                                
                                    'End If
                                
                                'End If
                            
                            'End If
                
                        'Case Else
                    'End Select
                
                'Next a
                
            'End If
        
        'End If
        
        '----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

        '**************************************************************************************************************
        '                                                             Lancement des charges
        '**************************************************************************************************************
        
        '--- lancement de la gamme pour C13 ---
        If NumChargeALancerPourC13 > 0 Then
            If TEtatsCharges(NumChargeALancerPourC13).PtrZoneGammeAnodisation = 0 Then
                
                '--- affectation du pointeur pour lancer la gamme ---
                With TEtatsCharges(NumChargeALancerPourC13)
                    If .PtrZoneGammeAnodisation = 0 Then
                        .PtrZoneGammeAnodisation = 1
                    End If
                End With
                
                '--- affichage des informations sur les entr�es des charges ---
                AfficheRenseignementsEntreesCharges ROUGE_3, "C13 IMPOSE - Gamme d'anodisation lancer pour la charge " & NumChargeALancerPourC13 & vbCrLf
            
            End If
        End If
        
        '----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
        
        '--- lancement de la gamme pour C14 ---
        If NumChargeALancerPourC14 > 0 Then
            If TEtatsCharges(NumChargeALancerPourC14).PtrZoneGammeAnodisation = 0 Then
                 
                '--- affectation du pointeur pour lancer la gamme ---
                With TEtatsCharges(NumChargeALancerPourC14)
                    If .PtrZoneGammeAnodisation = 0 Then
                        .PtrZoneGammeAnodisation = 1
                    End If
                End With
                
                '--- affichage des informations sur les entr�es des charges ---
                AfficheRenseignementsEntreesCharges ROUGE_3, "C14 IMPOSE - Gamme d'anodisation lancer pour la charge " & NumChargeALancerPourC14 & vbCrLf
            
            End If
        End If
        
        '----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
        
        '--- lancement de la gamme pour C15 ---
        If NumChargeALancerPourC15 > 0 Then
            If TEtatsCharges(NumChargeALancerPourC15).PtrZoneGammeAnodisation = 0 Then
                
                '--- affectation du pointeur pour lancer la gamme ---
                With TEtatsCharges(NumChargeALancerPourC15)
                    If .PtrZoneGammeAnodisation = 0 Then
                        .PtrZoneGammeAnodisation = 1
                    End If
                End With
                
                '--- affichage des informations sur les entr�es des charges ---
                AfficheRenseignementsEntreesCharges ROUGE_3, "C15 IMPOSE - Gamme d'anodisation lancer pour la charge " & NumChargeALancerPourC15 & vbCrLf
            
            End If
        End If
        
        '----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
        
        '--- lancement de la gamme pour C16 ---
        'If NumChargeALancerPourC16 > 0 Then
        '    If TEtatsCharges(NumChargeALancerPourC16).PtrZoneGammeAnodisation = 0 Then
        '
        '        '--- affectation du pointeur pour lancer la gamme ---
        '        with TEtatsCharges(NumChargeALancerPourC16)
        '            If .PtrZoneGammeAnodisation = 0 Then
        '                .PtrZoneGammeAnodisation = 1
        '            End If
        '       end with
        '
        '        '--- affichage des informations sur les entr�es des charges ---
        '        AfficheRenseignementsEntreesCharges ROUGE_3, "C16 IMPOSE - Gamme d'anodisation lancer pour la charge " & NumChargeALancerPourC16 & vbCrLf
        '
        '    End If
        'End If
        
        '----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
        
        '--- lancement de la gamme avec choix du poste d'anodisation sur automatique ---
        If EntreePossibleChargeAvecAnodisationAutomatique = True Then
            If NumChargePosteChargementSiAnodisationAutomatique > 0 Then
        
                '--- affectation du pointeur pour lancer la gamme ---
                With TEtatsCharges(NumChargePosteChargementSiAnodisationAutomatique)
                    If .PtrZoneGammeAnodisation = 0 Then
                        .PtrZoneGammeAnodisation = 1
                    End If
                End With
        
                '--- affichage des informations sur les entr�es des charges ---
                AfficheRenseignementsEntreesCharges ROUGE_3, "AUTOMATIQUE - Gamme d'anodisation lancer pour la charge " & NumChargePosteChargementSiAnodisationAutomatique & vbCrLf
        
            End If
        End If

    End With

End Sub

'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' R�le      : G�re le pr�visionnel afin de communiquer la meilleure entr�e des charges
' Entr�es :
' Retours :
' D�tails  :
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Sub GestionPrevisionnel()

    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next
    
    '--- constantes priv�es ---
    
    '--- d�claration ---
    Dim PassageZoneSpectrocoloration As Boolean                                                          'indique le passage dans la zone de spectrocoloration
    Dim PassageZoneOr As Boolean                                                                                   'indique le passage dans la zone d'or
    Dim PassageZoneNoir As Boolean                                                                                'indique le passage dans la zone de noir
    
    Dim a As Integer                                                                                                              'r�serv� pour les boucles FOR ... NEXT
    Dim NbrChargesPrevisionnel As Integer                                                                        'nombre de charges dans le pr�visionnel
    Dim TChoixIAChargesPrevisionnel(1 To NBR_LIGNES_PREVISIONNEL) As Integer     'tableau contenant le choix des charges pour le pr�visionnel
    Dim CptChoix As Integer                                                                                                  'compteur du choix
    Dim NumCharge As Integer                                                                                             'indique un num�ro de charge
    
    Dim NbrChargesGammeSpectrocoloration As Integer                                                    'nombre de charges avec une gamme spectrocoloration
    Dim NbrChargesGammeOr As Integer                                                                             'nombre de charges avec une gamme or
    Dim NbrChargesGammenoir As Integer                                                                          'nombre de charges avec une gamme noir
    
    Dim TempsMouvementsAvantPostePrincipalSecondes As Long                                  'temps des mouvements avant le poste principal en secondes
    Dim TempsAvantPostePrincipalAvecPontsSecondes As Long                                      'temps avant le poste principal avec les ponts en secondes
    Dim TempsPostePrincipalAvecPontsSecondes As Long                                               'temps au poste principal avec les ponts en secondes
    Dim TempsMouvementsApresPostePrincipalSecondes As Long                                  'temps des mouvements apr�s le poste principal en secondes
    Dim TempsApresPostePrincipalAvecPontsSecondes As Long                                      'temps apr�s le poste principal avec les ponts en secondes
    Dim TempsTotalPostesAvecPontsSecondes As Long                                                   'temps total des postes avec les ponts en secondes
    Dim TempsTotalEgouttagesAvecPontsSecondes As Long                                            'temps total des �gouttages avec les ponts en secondes
    Dim TempsTotalMouvementsSecondes As Long                                                          'temps total des mouvements en secondes
    Dim TempsTotalGammeAvecPontsSecondes As Long                                                  'temps total de la gamme avec les ponts en secondes

    Dim TGammesAnodisation As EnrGammesAnodisation                                               'repr�sente une gamme d'anodisation


    '--- affectation ---
    CptChoix = 1

    '--- analyse du pr�visionnel ---
    If TEtatsPonts(PONTS.P_1).NumCharge = CHARGES.PAS_DE_CHARGE Then

        '----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
        
        '--- analyse des gammes dans le pr�visionnel ---
        For a = 1 To NBR_LIGNES_PREVISIONNEL

            With TPrevisionnel(a)

                If .NumCommandeInterne <> "" And .CodeClient <> "" And .NumGammeAnodisation <> "" Then
            
                    '--- incr�mentation du nombre de charges dans le pr�visionnel ---
                    Inc NbrChargesPrevisionnel
                
                    If RechercheGammesAnodisation(.NumGammeAnodisation) = TROUVE Then
                       
                        '--- affectation de la gamme d'anodisation ---
                        TGammesAnodisation = TTempEnrGammesAnodisation                         'repr�sente une gamme d'anodisation
                        
                        '--- calcul les temps principaux d'une gamme d'anodisation AVEC LES TEMPS DE MOUVEMENTS DES PONTS ---
                        CalculTempsGammeAnodisationAvecPonts TGammesAnodisation, _
                                                                                               TempsMouvementsAvantPostePrincipalSecondes, _
                                                                                               TempsAvantPostePrincipalAvecPontsSecondes, _
                                                                                               TempsPostePrincipalAvecPontsSecondes, _
                                                                                               TempsMouvementsApresPostePrincipalSecondes, _
                                                                                               TempsApresPostePrincipalAvecPontsSecondes, _
                                                                                               TempsTotalPostesAvecPontsSecondes, _
                                                                                               TempsTotalEgouttagesAvecPontsSecondes, _
                                                                                               TempsTotalMouvementsSecondes, _
                                                                                               TempsTotalGammeAvecPontsSecondes, _
                                                                                               PassageZoneSpectrocoloration, _
                                                                                               PassageZoneOr, _
                                                                                               PassageZoneNoir
                    
                        '--- affectation du nombre de charges avec coloration ---
                        NbrChargesGammeSpectrocoloration = RechercheEnPreparationNbrChargesGammeSpectrocoloration()
                        NbrChargesGammeOr = RechercheEnPreparationNbrChargesGammeOr()
                        NbrChargesGammenoir = RechercheEnPreparationNbrChargesGammeNoir()
                        
                        '--- choix des charges avec coloration ---
                        If NbrChargesGammeSpectrocoloration = 0 And PassageZoneSpectrocoloration = True Then
                            TChoixIAChargesPrevisionnel(a) = CptChoix
                            Inc CptChoix
                        End If
                        If NbrChargesGammeOr = 0 And NbrChargesGammenoir = 0 And PassageZoneOr = True Then
                            TChoixIAChargesPrevisionnel(a) = CptChoix
                            Inc CptChoix
                        End If
                        If NbrChargesGammenoir = 0 And NbrChargesGammeOr = 0 And PassageZoneNoir = True Then
                            TChoixIAChargesPrevisionnel(a) = CptChoix
                            Inc CptChoix
                        End If
                    
                    End If
                
                Else
                
                    '--- sortie directe ---
                    Exit For
                
                End If

            End With


        Next a

        '----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
        
        '--- remplissage du tableau avec le choix ---
        If NbrChargesPrevisionnel > 0 Then
            For a = 1 To NBR_LIGNES_PREVISIONNEL
                TPrevisionnel(a).ChoixIA = TChoixIAChargesPrevisionnel(a)
            Next a
        End If
    
    End If

End Sub


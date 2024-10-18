Attribute VB_Name = "MIntelligenceArtificielle"
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' R�le                    : MODULE DE GESTION DE L'INTELLIGENCE ARTIFICIELLE
' Nom                    : MIntelligenceArtificielle.bas
' Date de cr�ation : 08/11/2000
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

'--- d�clarations obligatoires ---
Option Explicit

'--- options g�n�rales ---
Option Base 1
DefVar A-Z

'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' R�le      : Contr�le l'anti-collision entre les ponts avant tout d�placement � vide et transfert de charge
' Entr�es :              NumPontAControlerPont -> Num�ro du pont � contr�ler
'                 NumPosteDepartPontAControler -> Num�ro du poste de d�part du pont � contr�ler
'                NumPosteArriveePontAControler -> Num�ro du poste d'arriv�e du pont � contr�ler
' Retours :          ControleCollisionPossible -> Contient le message de la r�ponse
'                                             TypeCollision -> Fonction de l'�num�ration TYPES_COLLISION
'                                                                        0 = pas de risque de collision, les autres valeurs repr�sente le
'                                                                        num�ro du type de collision test�
'                                        NumPontOppose -> Num�ro du pont oppos�
'                       NumPosteAssurantSecurite -> Num�ro du poste assurant la s�curit� (poste ou doit se rendre le pont
'                                                                        oppos�
'                                        CouleurReponse -> Couleur de la r�ponse
' D�tails  :
'
'
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function ControleAntiCollision(ByVal NumPontAControler As Integer, _
                                                            ByVal NumPosteDepartPontAControler As Integer, _
                                                            ByVal NumPosteArriveePontAControler As Integer, _
                                                            ByRef TypeCollision As Integer, _
                                                            ByRef NumPontOppose As Integer, _
                                                            ByRef NumPosteAssurantSecurite As Integer, _
                                                            ByRef CouleurReponse As Long) As String
    
    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next
                                                                                                       
    '--- constantes priv�es ---
    'Const DISTANCE_SECURITE As Long = 6300
    'Call Log("Distance =" & DISTANCE_SECURITE)
    '--- d�claration ---
    Dim a As Integer                                                                 'r�serv� pour les boucles FOR ... NEXT
    Dim NumPosteDepartPontOppose As Integer                    'num�ro du poste de d�part du pont oppos�
    Dim NumPosteArriveePontOppose As Integer                   'num�ro du poste d'arriv�e du pont oppos�
    Dim NumPosteDebutAnalyse As Integer                            'num�ro de poste de d�but d'une analyse
    Dim NumPosteFinAnalyse As Integer                                 'num�ro de poste de fin d'une analyse
    
    Dim XPosteDepartPont1 As Long                                       'repr�sente les coordonn�es laser du poste de
                                                                                                'd�part du pont 1
    Dim XPosteArriveePont1 As Long                                      'repr�sente les coordonn�es laser du poste
                                                                                                'arriv�e du pont 1
    Dim XPosteDepartPont2 As Long                                       'repr�sente les coordonn�es laser du poste de
                                                                                                'd�part du pont 2
    Dim XPosteArriveePont2 As Long                                      'repr�sente les coordonn�es laser du poste
                                                                                                'arriv�e du pont 2
    
    Dim ReponseDeBase As String                                         'correpond � la r�ponse de base sans la cas pr�cis d'anti-collisoion
    Dim Reponse As String                                                      'correpond � la valeur de retour de la fonction

    '--- pour le d�boguage ---
    Dim Couleur As Long                                                          'repr�sente une couleur quelconque
    Dim Texte As String                                                            'repr�sente un texte quelconque

    '--- affectation par d�faut ---
    TypeCollision = TYPES_COLLISION.AUCUN_RISQUE      'type de collision par d�faut
    Reponse = "AUCUN RISQUE DE COLLISION"                   'texte de la r�ponse par d�faut pour l'anti-collision
    NumPosteAssurantSecurite = 0                                         'RAZ du poste de s�curit�
    ControleAntiCollision = ""                                                   'RAZ de la valeur de retour

    '***************************************************************************************************************
    '                                    Sortie directe de la fonction si un des ponts est condamn�
    '***************************************************************************************************************
    If TEtatsPonts(PONTS.P_1).Condamnation = True Or TEtatsPonts(PONTS.P_2).Condamnation = True Then
        Exit Function
    End If
    
    If (NumPontAControler = PONTS.P_1 Or NumPontAControler = PONTS.P_2) And _
        NumPosteDepartPontAControler >= POSTES.P_CHGT_1 And NumPosteDepartPontAControler <= DERNIER_POSTE And _
        NumPosteArriveePontAControler >= POSTES.P_CHGT_1 And NumPosteArriveePontAControler <= DERNIER_POSTE Then
        
        '***************************************************************************************************************
        '                                            Affectation du num�ro de pont oppos�
        '***************************************************************************************************************
        NumPontOppose = IIf(NumPontAControler = PONTS.P_1, PONTS.P_2, PONTS.P_1)
        
        '***************************************************************************************************************
        '                         Extraction des valeurs des axes de postes pour le PONT A CONTROLER
        '***************************************************************************************************************
        If NumPontAControler = PONTS.P_1 Then
            
            '--- le pont � contr�ler est le 1 ---
            XPosteDepartPont1 = TEtatsPostes(NumPosteDepartPontAControler).DefinitionPoste.XAxePosteLigne
            XPosteArriveePont1 = TEtatsPostes(NumPosteArriveePontAControler).DefinitionPoste.XAxePosteLigne
        
        Else
            
            '--- le pont � contr�ler est le 2 ---
            XPosteDepartPont2 = TEtatsPostes(NumPosteDepartPontAControler).DefinitionPoste.XAxePosteLigne
            XPosteArriveePont2 = TEtatsPostes(NumPosteArriveePontAControler).DefinitionPoste.XAxePosteLigne
        
        End If
        
        '***************************************************************************************************************
        '                        Extraction des valeurs des axes de postes pour le PONT OPPOSE
        '***************************************************************************************************************
        With TEtatsPonts(NumPontOppose)
        
            '--- affectation des n� de postes ---
            NumPosteDepartPontOppose = .TParametresCyclesPonts(CYCLES.C_ACTUEL).NumPosteDepart
            NumPosteArriveePontOppose = .TParametresCyclesPonts(CYCLES.C_ACTUEL).NumPosteArrivee
    
            '--- au lancement de programme il n'y a pas eu de d�placement ou de transfert ---
            'la valeur par d�faut est dans ce la valeur du poste actuel
            If NumPosteDepartPontOppose = 0 Or NumPosteArriveePontOppose = 0 Then
                NumPosteDepartPontOppose = .PosteActuel
                NumPosteArriveePontOppose = NumPosteDepartPontOppose
            End If
    
            '--- affectation des valeurs laser ---
            If NumPontOppose = PONTS.P_1 Then
                
                '--- le pont oppos� est le 1 ---
                XPosteDepartPont1 = TEtatsPostes(NumPosteDepartPontOppose).DefinitionPoste.XAxePosteLigne
                XPosteArriveePont1 = TEtatsPostes(NumPosteArriveePontOppose).DefinitionPoste.XAxePosteLigne
            
            Else
                
                '--- le pont oppos� est le 2 ---
                XPosteDepartPont2 = TEtatsPostes(NumPosteDepartPontOppose).DefinitionPoste.XAxePosteLigne
                XPosteArriveePont2 = TEtatsPostes(NumPosteArriveePontOppose).DefinitionPoste.XAxePosteLigne
            
            End If
    
        End With
        
        '----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
        
        '***************************************************************************************************************
        '                                                                     DEBOGUAGE
        '***************************************************************************************************************
        If DEBUG_MODE = True Then
        
            Texte = "---------------------------------------------------------"
            AfficheRenseignementsDebug Couleur, Texte & vbCrLf
            
            Couleur = COULEURS.VERT_4
            Texte = "CONTROLE SUR P" & NumPontAControler
            Texte = Texte & " D�part = " & NumPosteDepartPontAControler & " (" & TEtatsPostes(NumPosteDepartPontAControler).DefinitionPoste.NomPoste & ")"
            Texte = Texte & " Arriv�e = " & NumPosteArriveePontAControler & " (" & TEtatsPostes(NumPosteArriveePontAControler).DefinitionPoste.NomPoste & ")"
            AfficheRenseignementsDebug Couleur, Texte & vbCrLf
            
            Texte = "AUTRE PONT P" & NumPontOppose
            Texte = Texte & " D�part = " & NumPosteDepartPontOppose & " (" & TEtatsPostes(NumPosteDepartPontOppose).DefinitionPoste.NomPoste & ")"
            Texte = Texte & " Arriv�e = " & NumPosteArriveePontOppose & " (" & TEtatsPostes(NumPosteArriveePontOppose).DefinitionPoste.NomPoste & ")"
            AfficheRenseignementsDebug Couleur, Texte & vbCrLf
            
            
            Texte = "Position du pont 2: " & TEtatsPonts(PONTS.P_2).PositionActuelleLaserTrlPont
            AfficheRenseignementsDebug Couleur, Texte & vbCrLf
            
            Texte = "Position du pont 1: " & TEtatsPonts(PONTS.P_1).PositionActuelleLaserTrlPont
            AfficheRenseignementsDebug Couleur, Texte & vbCrLf
            
             Texte = "Poste Assurant Securite: " & TEtatsPostes(NumPosteAssurantSecurite).DefinitionPoste.NomPoste
            AfficheRenseignementsDebug Couleur, Texte & vbCrLf
            
            
            Dim diff As Integer
            diff = XPosteDepartPont2 - XPosteDepartPont1
            
            
            
            Texte = "XPosteDepartPont2 - XPosteDepartPont1: " & diff
            AfficheRenseignementsDebug Couleur, Texte & vbCrLf
            
            
            diff = XPosteDepartPont2 - TEtatsPonts(PONTS.P_1).PositionActuelleLaserTrlPont
            Texte = "XPosteDepartPont2 - PONTS.P_1.PositionActuelleLaserTrlPont: " & diff
            AfficheRenseignementsDebug Couleur, Texte & vbCrLf
              
         
        End If
        
        
        
        '----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
        
        '***************************************************************************************************************
        '                                          GESTION DES CAS DE RISQUES DE COLLISION
        '***************************************************************************************************************
            
        '********************************************************************************************************
        'le pont 2 avance dans la ligne et le pont 1 recule
        'cas de collision possible
        '             PONT 2             PONT 1           OU      PONT 2   A <------------- D
        '       A <------------- D     D -------------> A                                            D -------------> A   PONT 1
        '********************************************************************************************************
        If XPosteArriveePont1 <= XPosteDepartPont1 And XPosteArriveePont2 >= XPosteDepartPont2 Then
            
            If NumPontAControler = PONTS.P_1 Then

                '--- le pont � contr�ler est le PONT 1, il faut v�rifier les mouvements du PONT 2 ---
                '             PONT 2             PONT 1           OU      PONT 2   A <------------- D
                '       A <------------- D     D -------------> A                                            D -------------> A   PONT 1
                
                '--- affectation de la r�ponse de base ---
                ReponseDeBase = "ANTI-COLLISION - DEMANDE P1->AR, P2->AV - "
                AfficheRenseignementsDebug Couleur, "- PASSAGE - " & ReponseDeBase & vbCrLf
                
                With TEtatsPonts(PONTS.P_2)
                    Select Case .SensX
                        Case SENS_X.S_ARRIERE
                            '--- sens arri�re (le pont 2 recule pour aller � son poste de d�part) ---
                            If XPosteDepartPont2 - XPosteDepartPont1 < DISTANCE_SECURITE Then
                                TypeCollision = TYPES_COLLISION.RISQUE_DEM_P1_AR_P2_AV
                                Reponse = ReponseDeBase & "CAS 1"
                            End If
                        Case SENS_X.S_AVANT
                            '--- sens avant (le pont 2 avance et son laser de destination est le poste de destination) ---
                            If .PositionCibleLaserTrlPont = XPosteArriveePont2 Then
                                If XPosteArriveePont2 - XPosteDepartPont1 < DISTANCE_SECURITE Then
                                    TypeCollision = TYPES_COLLISION.RISQUE_DEM_P1_AR_P2_AV
                                    Reponse = ReponseDeBase & "CAS 2"
                                End If
                            Else
                                TypeCollision = TYPES_COLLISION.RISQUE_DEM_P1_AR_P2_AV
                                Reponse = ReponseDeBase & "CAS 3"
                            End If
                        Case Else
                            '--- arr�t au poste ---
                            If .PositionActuelleLaserTrlPont - XPosteDepartPont1 < DISTANCE_SECURITE Then
                                TypeCollision = TYPES_COLLISION.RISQUE_DEM_P1_AR_P2_AV
                                Reponse = ReponseDeBase & "CAS 4"
                            End If
                    End Select
                End With

            Else

                '--- le pont � contr�ler est le PONT 2, il faut v�rifier les mouvements du PONT 1 ---
                '             PONT 2             PONT 1           OU      PONT 2   A <------------- D
                '       A <------------- D     D -------------> A                                            D -------------> A   PONT 1
                
                '--- affectation de la r�ponse de base ---
                ReponseDeBase = "ANTI-COLLISION - DEMANDE P2->AV, P1->AR - "
                AfficheRenseignementsDebug Couleur, "- PASSAGE - " & ReponseDeBase & vbCrLf
                
                With TEtatsPonts(PONTS.P_1)
                    Select Case .SensX
                        Case SENS_X.S_AVANT
                            '--- sens avant (le pont 1 avance pour aller � son poste de d�part) ---
                            If XPosteDepartPont2 - XPosteDepartPont1 < DISTANCE_SECURITE Then
                                TypeCollision = TYPES_COLLISION.RISQUE_DEM_P2_AV_P1_AR
                                Reponse = ReponseDeBase & "CAS 1"
                            End If
                        Case SENS_X.S_ARRIERE
                            '--- sens avant (le pont 1 recule et son laser de destination est le poste de destination) ---
                            If .PositionCibleLaserTrlPont = XPosteArriveePont1 Then
                                If XPosteDepartPont2 - XPosteArriveePont1 < DISTANCE_SECURITE Then
                                    TypeCollision = TYPES_COLLISION.RISQUE_DEM_P2_AV_P1_AR
                                    Reponse = ReponseDeBase & "CAS 2"
                                End If
                            Else
                                TypeCollision = TYPES_COLLISION.RISQUE_DEM_P2_AV_P1_AR
                                Reponse = ReponseDeBase & "CAS 3"
                            End If
                        Case Else
                            '--- arr�t au poste ---
                            If XPosteDepartPont2 - .PositionActuelleLaserTrlPont < DISTANCE_SECURITE Then
                                TypeCollision = TYPES_COLLISION.RISQUE_DEM_P2_AV_P1_AR
                                Reponse = ReponseDeBase & "CAS 4"
                            End If
                    End Select
                End With
        
            End If
        
        End If
        
        
        '********************************************************************************************************
        'le pont 2 recule dans la ligne et le pont 1 avance
        'cas de collision possible
        '            PONT 2               PONT 1           OU      PONT 2   D -------------> A
        '       D -------------> A     A <------------- D                                          A <------------- D   PONT 1
        '********************************************************************************************************
        If XPosteArriveePont1 >= XPosteDepartPont1 And XPosteArriveePont2 <= XPosteDepartPont2 Then
            
            '--- affectation de la r�ponse de base ---
            ReponseDeBase = "ANTI-COLLISION - DEMANDE P1->AV, P2->AR - "
            AfficheRenseignementsDebug Couleur, "- PASSAGE - " & ReponseDeBase & vbCrLf
            
            '--- v�rifier la distance de s�curit� ---
            If XPosteArriveePont2 - XPosteArriveePont1 < DISTANCE_SECURITE Then
                
                '--- affectation ---
                If NumPontAControler = PONTS.P_1 Then
                    TypeCollision = TYPES_COLLISION.RISQUE_DEM_P1_AV_P2_AR
                    Reponse = ReponseDeBase & "CAS 1"
                Else
                    TypeCollision = TYPES_COLLISION.RISQUE_DEM_P2_AR_P1_AV
                    Reponse = ReponseDeBase & "CAS 2"
                End If
            
            End If
        
        End If
                
        '********************************************************************************************************
        'les 2 ponts avancent dans la ligne
        'cas de collision possible
        '             PONT 2              PONT 1           OU       PONT 2   A <------------- D
        '       A <------------- D     A <------------- D                                             A <------------- D   PONT 1
        '********************************************************************************************************
        If XPosteArriveePont1 >= XPosteDepartPont1 And XPosteArriveePont2 >= XPosteDepartPont2 Then
            
            If NumPontAControler = PONTS.P_1 Then
                
                '--- le pont � contr�ler est le PONT 1, il faut v�rifier les mouvements du PONT 2 ---
                '             PONT 2              PONT 1            OU       PONT 2   A <------------- D
                '       A <------------- D     A <------------- D                                             A <------------- D   PONT 1
                
                '--- affectation de la r�ponse de base ---
                ReponseDeBase = "ANTI-COLLISION - DEMANDE P1->AV, P2->AV - "
                AfficheRenseignementsDebug Couleur, "- PASSAGE - " & ReponseDeBase & vbCrLf
                
                With TEtatsPonts(PONTS.P_2)
                    Select Case .SensX
                        Case SENS_X.S_ARRIERE
                            '--- sens arri�re  (le pont 2 recule pour aller � son poste de d�part) ---
                            If XPosteDepartPont2 - XPosteArriveePont1 < DISTANCE_SECURITE Then
                                TypeCollision = TYPES_COLLISION.RISQUE_DEM_P1_AV_P2_AV
                                Reponse = ReponseDeBase & "CAS 1"
                            End If
                        Case SENS_X.S_AVANT
                            '--- sens avant (le pont 2 avance et son laser de destination est le poste de destination) ---
                            If .PositionCibleLaserTrlPont = XPosteArriveePont2 Then
                                If XPosteArriveePont2 - XPosteArriveePont1 < DISTANCE_SECURITE Then
                                    TypeCollision = TYPES_COLLISION.RISQUE_DEM_P1_AV_P2_AV
                                    Reponse = ReponseDeBase & "CAS 2"
                                End If
                            Else
                                TypeCollision = TYPES_COLLISION.RISQUE_DEM_P1_AV_P2_AV
                                Reponse = ReponseDeBase & "CAS 3"
                            End If
                        Case Else
                            '--- arr�t au poste ---
                            If .PositionActuelleLaserTrlPont - XPosteArriveePont1 < DISTANCE_SECURITE Then
                                TypeCollision = TYPES_COLLISION.RISQUE_DEM_P1_AV_P2_AV
                                Reponse = ReponseDeBase & "CAS 4"
                            End If
                    End Select
                End With
                
            Else
                
                '--- le pont � contr�ler est le PONT 2, il faut v�rifier les mouvements du PONT 1 ---
                '             PONT 2              PONT 1            OU       PONT 2   A <------------- D
                '       A <------------- D     A <------------- D                                             A <------------- D   PONT 1
                'dans ce cas le pont 2 ne pourra aller dans le segment du pont 1
                
                '--- affectation de la r�ponse de base ---
                ReponseDeBase = "ANTI-COLLISION - DEMANDE P2->AV, P1->AV - "
                AfficheRenseignementsDebug Couleur, "- PASSAGE - " & ReponseDeBase & vbCrLf
                
                If XPosteDepartPont2 - XPosteArriveePont1 < DISTANCE_SECURITE Then
                    TypeCollision = TYPES_COLLISION.RISQUE_DEM_P2_AV_P1_AV
                    Reponse = ReponseDeBase & "CAS 1"
                End If
            
            End If
        
        End If
    
        '********************************************************************************************************
        'les 2 ponts reculent dans la ligne
        'cas de collision possible
        '            PONT 2             PONT 1            OU      PONT 2   D -------------> A
        '       D -------------> A     D -------------> A                                          D -------------> A   PONT 1
        '********************************************************************************************************
        If XPosteArriveePont1 <= XPosteDepartPont1 And XPosteArriveePont2 <= XPosteDepartPont2 Then
                
            If NumPontAControler = PONTS.P_1 Then
                         
                '--- le pont � contr�ler est le PONT 1, il faut v�rifier les mouvements du PONT 2 ---
                '            PONT 2             PONT 1            OU      PONT 2   D -------------> A
                '       D -------------> A     D -------------> A                                          D -------------> A   PONT 1
                'dans ce le pont 1 ne pourra aller dans le segment du pont 2
                
                '--- affectation de la r�ponse de base ---
                ReponseDeBase = "ANTI-COLLISION - DEMANDE P1->AR, P2->AR - "
                AfficheRenseignementsDebug Couleur, "- PASSAGE - " & ReponseDeBase & vbCrLf
                
                If XPosteArriveePont2 - XPosteDepartPont1 < DISTANCE_SECURITE Then
                    TypeCollision = TYPES_COLLISION.RISQUE_DEM_P1_AR_P2_AR
                    Reponse = ReponseDeBase & "CAS 1"
                End If
            
            Else
                
                '--- le pont � contr�ler est le PONT 2, il faut v�rifier les mouvements du PONT 1 ---
                '            PONT 2             PONT 1            OU      PONT 2   D -------------> A
                '       D -------------> A     D -------------> A                                          D -------------> A   PONT 1
                
                '--- affectation de la r�ponse de base ---
                ReponseDeBase = "ANTI-COLLISION - DEMANDE P2->AR, P1->AR - "
                AfficheRenseignementsDebug Couleur, "- PASSAGE - " & ReponseDeBase & vbCrLf
                
                With TEtatsPonts(PONTS.P_1)
                    Select Case .SensX
                        Case SENS_X.S_AVANT
                            '--- sens avant (le pont 1 avance pour aller � son poste de d�part) ---
                            ' If XPosteArriveePont2 - XPosteDepartPont1 < DISTANCE_SECURITE Then
                            If XPosteArriveePont2 - XPosteDepartPont1 < DISTANCE_SECURITE Then
                                TypeCollision = TYPES_COLLISION.RISQUE_DEM_P2_AR_P1_AR
                                Reponse = ReponseDeBase & "CAS 1"
                            End If
                        Case SENS_X.S_ARRIERE
                            '--- sens arri�re (le pont 1 recule et son laser de destination est le poste de destination) ---
                            If .PositionCibleLaserTrlPont = XPosteArriveePont1 Then
                                If XPosteArriveePont2 - XPosteArriveePont1 < DISTANCE_SECURITE Then
                                    TypeCollision = TYPES_COLLISION.RISQUE_DEM_P2_AR_P1_AR
                                    Reponse = ReponseDeBase & "CAS 2"
                                End If
                            Else
                                TypeCollision = TYPES_COLLISION.RISQUE_DEM_P2_AR_P1_AR
                                Reponse = ReponseDeBase & "CAS 3"
                            End If
                        Case Else
                            '--- arr�t au poste ---
                            If XPosteArriveePont2 - .PositionActuelleLaserTrlPont < DISTANCE_SECURITE Then
                                TypeCollision = TYPES_COLLISION.RISQUE_DEM_P2_AR_P1_AR
                                Reponse = ReponseDeBase & "CAS 4"
                            End If
                    End Select
                End With

            End If
        
        End If
    
    
        '********************************************************************************************************
        '                                      RECHERCHE DU POSTE ASSURANT LA SECURITE
        '********************************************************************************************************
        If TypeCollision <> TYPES_COLLISION.AUCUN_RISQUE Then

            If TEtatsPonts(PONTS.P_1).PtrEtActionEnCoursAPI.PtrAction = 0 And _
               TEtatsPonts(PONTS.P_2).PtrEtActionEnCoursAPI.PtrAction = 0 Then

                If NumPontAControler = PONTS.P_1 Then

                    '--- le pont � contr�ler est le PONT 1, il faut DEPLACER LE PONT 2 (toujours en sens AVANT) ---
                    'recherche du poste le plus �loign�
                    'ATTENTION un pont qui AVANCE dans la ligne a sa VALEUR LASER QUI AUGMENTE
                    NumPosteDebutAnalyse = POSTES.P_CHGT_1
                    NumPosteFinAnalyse = DERNIER_POSTE

                    '--- lancement de l'analyse ---
                    For a = NumPosteDebutAnalyse To NumPosteFinAnalyse
                        With TEtatsPostes(a).DefinitionPoste
                            If XPosteArriveePont1 > XPosteDepartPont1 Then    'le pont s'est avanc� dans la ligne

                                '--- comparaison sur le poste D'ARRIVEE ---
                                If .XAxePosteLigne > XPosteArriveePont1 + DISTANCE_SECURITE Then
                                    NumPosteAssurantSecurite = a
                                    Exit For
                                End If

                            Else

                                '--- comparaison sur le poste de DEPART ---
                                If .XAxePosteLigne > XPosteDepartPont1 + DISTANCE_SECURITE Then
                                    NumPosteAssurantSecurite = a
                                    Exit For
                                End If

                            End If
                        End With
                    Next a

                Else

                    '--- le pont � contr�ler est le PONT 2, il faut DEPLACER LE PONT 1 (toujours en sens ARRIERE) ---
                    'recherche du poste le plus �loign�
                    'ATTENTION un pont qui AVANCE dans la ligne a sa VALEUR LASER QUI AUGMENTE
                    NumPosteDebutAnalyse = DERNIER_POSTE
                    NumPosteFinAnalyse = POSTES.P_CHGT_1

                    '--- lancement de l'analyse ---
                    For a = NumPosteDebutAnalyse To NumPosteFinAnalyse Step -1
                        With TEtatsPostes(a).DefinitionPoste
                            If XPosteArriveePont2 < XPosteDepartPont2 Then    'le pont a recul� dans la ligne

                                '--- comparaison sur le poste D'ARRIVEE ---
                                If .XAxePosteLigne < XPosteArriveePont2 - DISTANCE_SECURITE Then
                                    NumPosteAssurantSecurite = a
                                    Exit For
                                End If

                            Else

                                '--- comparaison sur le poste de DEPART ---
                                If .XAxePosteLigne < XPosteDepartPont2 - DISTANCE_SECURITE Then
                                    NumPosteAssurantSecurite = a
                                    Exit For
                                End If

                            End If
                        End With
                    Next a

                End If

            End If

        End If
    
    End If

    '----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    
    '***************************************************************************************************************
    '                                                                     DEBOGUAGE
    '***************************************************************************************************************
    Couleur = COULEURS.ROUGE_3
    Select Case TypeCollision
        Case 0: Texte = "AUCUN_RISQUE"
        Case 1: Texte = "RISQUE_DEM_P1_AR_P2_AV"
        Case 2: Texte = "RISQUE_DEM_P2_AV_P1_AR"
        Case 3: Texte = "RISQUE_DEM_P1_AV_P2_AR"
        Case 4: Texte = "RISQUE_DEM_P2_AR_P1_AV"
        Case 5: Texte = "RISQUE_DEM_P1_AV_P2_AV"
        Case 6: Texte = "RISQUE_DEM_P2_AV_P1_AV"
        Case 7: Texte = "RISQUE_DEM_P1_AR_P2_AR"
        Case 8: Texte = "RISQUE_DEM_P2_AR_P1_AR"
        Case Else
    End Select
    AfficheRenseignementsDebug Couleur, Texte & vbCrLf
    AfficheRenseignementsDebug Couleur, Reponse & vbCrLf
    
    
    'Call Log("controle antic collision = " & Texte)
    'Call Log("controle antic collision Reponse" & Reponse)
    '----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    
    '--- valeur de retour / couleur de la r�ponse ---
    ControleAntiCollision = Reponse
    If TypeCollision = TYPES_COLLISION.AUCUN_RISQUE Then
        CouleurReponse = COULEURS.BLEU_3
    Else
        CouleurReponse = COULEURS.ROUGE_3
        If NumPosteAssurantSecurite >= POSTES.P_CHGT_1 And NumPosteAssurantSecurite <= DERNIER_POSTE Then
            ControleAntiCollision = ControleAntiCollision & _
                                                  " (DEPLACEMENT DU PONT " & NumPontOppose & _
                                                  " en " & _
                                                  TEtatsPostes(NumPosteAssurantSecurite).DefinitionPoste.NomPoste & _
                                                  ")"
        End If
    End If
    
End Function

'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' R�le      : Recherche l'ordre de sortie des charges par les temps les plus courts restant dans les postes
' Entr�es :
' Retours : Les valeurs recherch�es se trouvent dans la tableau TOrdreSortieCharges qui se trouve
'                 lui-m�me dans le tableau TMoteurInference du moteur d'inf�rence
' D�tails  : - l'ordre de sortie des charges s'effectue par les temps les plus courts restant dans les postes
'                   � condition que les postes ne soient pas condamn�s
'                - les charges se trouvant dans les postes condamn�s se retrouvent en queue de liste
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Sub RechercheOrdreSortieCharges()

    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next
                                                                                                       
    '--- constantes priv�es ---

    '--- d�claration ---
    Dim a As Integer                                                                                               'r�serv� pour les boucles FOR ... NEXT
    Dim b As Integer                                                                                               'r�serv� pour les boucles FOR ... NEXT
    Dim NumCharge As Integer                                                                              'indique un num�ro de charge
    Dim CptPostes As Integer                                                                                 'compteur des postes pour pointer dans le tableau
                                                                                                                              'de l'ordre de sortie des charges
    Dim TCptOrdreSortiePonts(PONTS.P_1 To PONTS.P_2) As Integer                 'compteur des ordres de sortie par pont
    Dim PtrZoneGammeAnodisation As Integer                                                     'pointeur de la zone de la gamme d'anodisation
            
    Dim NumPont As Integer                                                                                  'num�ro de pont
    Dim NumZoneDepart As Integer                                                                      'num�ro de la zone de d�part
    Dim NumZoneArrivee As Integer                                                                     'num�ro de la zone d'arriv�e
                  
    Dim NumPosteArrivee As Integer                                                                    'num�ro du poste d'arriv�e
    Dim NumChargePosteArrivee As Integer                                                         'num�ro de charge du poste d'arriv�e
                  
                  '********** CORRESPOND AUX DETAILS DES GAMMES D'ANODISATION DES CHARGES **********
    
    Dim NumPosteReel As Integer                                                                        'N� de poste r�el utilis� dans la zone (cas des postes multiples)
                                                                                                        
    Dim DecompteDuTempsAuPosteReelSecondes As String                              'Repr�sente la diff�rence entre le temps th�orique
                                                                                                                              'au poste et le temps r�el pass� dans le poste
                                                                                                                              'un nombre n�gatif apparait si la charge est rest� plus
                                                                                                                              'longtemps dans le poste que le temps th�orique pr�vu
                                                                                                                              'ATTENTION variable du type String volontairement
                                                                                                                              'Si "" alors il n'y a pas eu de temps de d�compter
    
    Dim FicheOrdreSortieCharges As VarOrdreSortieCharges                              'fiche de l'ordre de sortie des charges
    Dim FicheVideOrdreSortieCharges As VarOrdreSortieCharges                       'fiche vide de l'ordre de sortie des charges
    
    '--- l'analyse se fait uniquement avec les bains ---
    For a = PREMIER_BAIN To DERNIER_POSTE
        
        If a <> POSTES.P_D1 And a <> POSTES.P_D2 Then
            With TEtatsPostes(a)

            '--- affectation du n� de charge ---
            NumCharge = .NumCharge
            
            '--- recherche du temps le plus court ---
            If NumCharge >= CHARGES.C_NUM_MINI And NumCharge <= CHARGES.C_NUM_MAXI Then
                
                '--- affectation du pointeur de la zone de la gamme d'anodisation ---
                PtrZoneGammeAnodisation = TEtatsCharges(NumCharge).PtrZoneGammeAnodisation
                
                If PtrZoneGammeAnodisation > 0 Then
                
                    With TEtatsCharges(NumCharge).TGammesAnodisation.TDetailsGammesAnodisation(PtrZoneGammeAnodisation)
                    
                        '--- affectation du n� du poste r�el ---
                        NumPosteReel = .NumPosteReel
                        
                        If a = .NumPosteReel Then 'v�rifier la concordance entre le poste scrut� et le poste r�el
                        
                            '--- affectation d�compte du temps au poste ---
                            DecompteDuTempsAuPosteReelSecondes = .DecompteDuTempsAuPosteReelSecondes
                    
                            '--- remplir le tableau avec le n� du poste ainsi que le temps de d�compte de celui-ci ---
                            If IsNumeric(DecompteDuTempsAuPosteReelSecondes) = True Then
                                
                                '--- incr�menter le compteur des postes ---
                                Inc CptPostes
                                
                                '--- affectation des zones de d�part et arriv�e ---
                                NumZoneDepart = TEtatsCharges(NumCharge).TGammesAnodisation.TDetailsGammesAnodisation(PtrZoneGammeAnodisation).NumZone
                                NumZoneArrivee = TEtatsCharges(NumCharge).TGammesAnodisation.TDetailsGammesAnodisation(PtrZoneGammeAnodisation + 1).NumZone
                            
                                '--- d�termination du num�ro du poste d'arriv�e et du num�ro de charge au poste d'arriv�e ---
                                NumPosteArrivee = ProchainNumTheoriquePosteArrivee(NumCharge, NumZoneArrivee)
                                
                                '--- affectation du num�ro de charge au poste d'arriv�e ---
                                NumChargePosteArrivee = 0
                                If NumPosteArrivee >= POSTES.P_CHGT_1 And NumPosteArrivee <= DERNIER_POSTE Then
                                    NumChargePosteArrivee = TEtatsPostes(NumPosteArrivee).NumCharge
                                End If
                                
                                '--- recherche le num�ro de pont choisi dans la pr�misse pour un d�placement entre 2 zones ---
                                NumPont = 0
                                If NumZoneDepart > 0 And NumZoneArrivee > 0 Then
                                    NumPont = RechercheNumPontChoisiDansPremisse(NumZoneDepart, NumZoneArrivee)
                                End If
                                
                                '--- remplir le tableau des ordres de sortie ---
                                With TMoteurInference.TOrdreSortieCharges(CptPostes)
                                    
                                    .NumPoste = a
                                    .NumCharge = NumCharge
                                    
                                    .NumPosteArrivee = NumPosteArrivee
                                    .NumChargePosteArrivee = NumChargePosteArrivee
                                    
                                    .DecompteDuTempsAuPosteReelSecondes = DecompteDuTempsAuPosteReelSecondes
                                    .Condamnation = TEtatsPostes(a).Condamnation
                                    .NumPont = NumPont

                                End With
                    
                            End If
                    
                        End If
                    
                    End With
                
                End If
                
            End If
        
        End With
        End If
        
    
    Next a

    '--- analyse en fonction du compteur des postes ---
    If CptPostes = 0 Then

        '--- effacement complet des tableaux ---
        Erase TMoteurInference.TOrdreSortieCharges()
        Erase TMoteurInference.TOrdreSortiePonts()
    
    Else

        '----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
        
        '--- vider le reste des fiches pour �liminer les anciennes fiches ---
        For a = CptPostes + 1 To UBound(TMoteurInference.TOrdreSortieCharges())
            TMoteurInference.TOrdreSortieCharges(a) = FicheVideOrdreSortieCharges
        Next a
        
        '--- tri du tableau du d�compte du temps au poste le plus petit au plus grand ---
        For a = 1 To CptPostes - 1
            For b = a + 1 To CptPostes
                
                If IsNumeric(TMoteurInference.TOrdreSortieCharges(a).DecompteDuTempsAuPosteReelSecondes) = True And IsNumeric(TMoteurInference.TOrdreSortieCharges(b).DecompteDuTempsAuPosteReelSecondes) = True And _
                    TMoteurInference.TOrdreSortieCharges(a).NumPosteArrivee > 0 And TMoteurInference.TOrdreSortieCharges(b).NumPosteArrivee > 0 Then
                
                    If (Val(TMoteurInference.TOrdreSortieCharges(a).DecompteDuTempsAuPosteReelSecondes) > Val(TMoteurInference.TOrdreSortieCharges(b).DecompteDuTempsAuPosteReelSecondes) And _
                       TMoteurInference.TOrdreSortieCharges(a).Condamnation = False And TMoteurInference.TOrdreSortieCharges(b).Condamnation = False) Or _
                       TMoteurInference.TOrdreSortieCharges(a).Condamnation = True Or _
                       TEtatsPostes(TMoteurInference.TOrdreSortieCharges(a).NumPosteArrivee).Condamnation = True Or _
                       TEtatsPostes(TMoteurInference.TOrdreSortieCharges(b).NumPosteArrivee).Condamnation = True Or _
                       TMoteurInference.TOrdreSortieCharges(a).NumChargePosteArrivee <> 0 Then
                    
                            '--- inversion des 2 fiches ---
                            FicheOrdreSortieCharges = TMoteurInference.TOrdreSortieCharges(a)
                            TMoteurInference.TOrdreSortieCharges(a) = TMoteurInference.TOrdreSortieCharges(b)
                            TMoteurInference.TOrdreSortieCharges(b) = FicheOrdreSortieCharges
                    
                    End If
            
                End If
            
            Next b
        Next a

        '----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
        
        '--- effacement complet du tableau des ordres de sortie pour les ponts  ---
        Erase TMoteurInference.TOrdreSortiePonts()
        
        '--- affectation pour les ponts ---
        For a = 1 To CptPostes
            
            '--- affectation du num�ro de pont ---
            NumPont = TMoteurInference.TOrdreSortieCharges(a).NumPont
            
            If NumPont >= PONTS.P_1 And NumPont <= PONTS.P_2 Then
                
                '--- incr�mentation du compteur ---
                TCptOrdreSortiePonts(NumPont) = TCptOrdreSortiePonts(NumPont) + 1
        
                '--- transfert de la fiche ---
                TMoteurInference.TOrdreSortiePonts(NumPont, TCptOrdreSortiePonts(NumPont)) = TMoteurInference.TOrdreSortieCharges(a)
        
            End If

        Next a

    End If

End Sub

'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' R�le      : Construit le prochain cycle d'un pont dans le tableau des �tats des ponts
'                 Les divers entr�es sont donn�es par le moteur d'inf�rence
'                 Cette fonction est utilis�e pour renseigner l'op�rateur dans l'�cran des cycles des ponts
' Entr�es :              ViderProchainCycle  -> FALSE = Le prochain cycle va �tre rempli par la pr�misse concern�e
'                                                                   TRUE = Le prochain cycle va �tre vider dans le tableau des �tats des
'                                                                                ponts
'                                             TypeCycle -> Type de cycle (d�placement ou transfert)
'                                                                   fonction de l'�num�ration TYPES_CYCLES
'                                               NumPont -> Num�ro du pont concern� par le prochain cycle
'                                  NumPosteDepart -> Num�ro du poste de d�part
'                                 NumPosteArrivee -> Num�ro du poste d'arriv�e
' Retours : ConstruitProchainCyclePont -> Contient le message du r�sultat du d�codage
'                                                                   OK = Construction du prochain cycle correctement effectu�
'                                                                   ""  = Mauvais poste de d�part ou arriv�e
'                                                                   PREMISSE_INEXISTANTE = la pr�misse n'existe pas
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function ConstruitProchainCyclePont(ByVal ViderProchainCycle As Boolean, _
                                                                        ByVal TypeCycle As TYPES_CYCLES, _
                                                                        ByVal NumPont As Integer, _
                                                                        ByVal NumPosteDepart As Integer, _
                                                                        ByVal NumPosteArrivee As Integer) As String

    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next
                                                                                                       
    '--- constantes priv�es ---

    '--- d�claration ---
    Dim a As Integer, _
           PtrLigne As Integer, _
           PtrAction As Integer, _
           NumAction As Integer, _
           TCyclePont(1 To NBR_LIGNES_CYCLES_PONTS) As Integer
    Dim PremisseDecodee As String
    Dim FicheVideCyclesPonts As CyclesPonts
    Dim TPremisseDecodee As Variant           'tableau de base contenant les actions apr�s d�codage
    
    '--- affectation ---
    ConstruitProchainCyclePont = ""
    
    If NumPont >= PONTS.P_1 And NumPont <= PONTS.P_2 Then

        If ViderProchainCycle = True Then
    
            '**********************************************************************************************************
            '                     Vidage du tableau du prochain cycle dans le tableau des �tats des ponts
            '**********************************************************************************************************
            For a = 1 To NBR_LIGNES_CYCLES_PONTS
                TEtatsPonts(NumPont).TCyclesPonts(CYCLES.C_PROCHAIN, a) = FicheVideCyclesPonts
            Next a
    
        Else
    
            '**********************************************************************************************************
            '                                                 Analyse en fonction du type de cycle
            '**********************************************************************************************************
            Select Case TypeCycle
            
                Case TYPES_CYCLES.TC_DEPLACEMENT_PONT
                    '*********************************************************************************************
                    '                                     Le cycle est un DEPLACEMENT DE PONT
                    '*********************************************************************************************
                    For a = 1 To NBR_LIGNES_CYCLES_PONTS
                        With TEtatsPonts(NumPont).TCyclesPonts(CYCLES.C_PROCHAIN, a)
                            Select Case a
                                Case 1
                                    .NumAction = NumPosteArrivee
                                    .Parametre = ""
                                    .EtatParametre = ""
                                Case 2
                                    .NumAction = NUM_ACTION_FCY
                                    .Parametre = ""
                                    .EtatParametre = ""
                                Case Else
                                    .NumAction = NUM_ACTION_NOP
                                    .Parametre = ""
                                    .EtatParametre = ""
                            End Select
                        End With
                    Next a
            
                    '*********************************************************************************************
                    '                                             Affectation des param�tres
                    '*********************************************************************************************
                    With TEtatsPonts(NumPont).TParametresCyclesPonts(CYCLES.C_PROCHAIN)
                        .TypeCycle = TYPES_CYCLES.TC_DEPLACEMENT_PONT
                        .NumPosteDepart = NumPosteDepart
                        .NumPosteArrivee = NumPosteArrivee
                        .DelaiSupStabilisationChargeSecondes = 0
                        .TempsEgouttageSecondes = 0
                    End With
            
                Case TYPES_CYCLES.TC_TRANSFERT_CHARGE
                    '*********************************************************************************************
                    '                                     Le cycle est un TRANSFERT DE CHARGE
                    '*********************************************************************************************
                    If NumPosteDepart >= POSTES.P_CHGT_1 And NumPosteDepart <= DERNIER_POSTE And _
                       NumPosteArrivee >= POSTES.P_CHGT_1 And NumPosteArrivee <= DERNIER_POSTE Then
            
                        '--- recherche de la pr�misse d�cod�e ---
                        PremisseDecodee = TPremisses(NumPosteDepart, NumPosteArrivee).PremisseDecodee
                                                
                        'Call Log("Prochain cycle pont, PremisseDecodee:" + PremisseDecodee)
                        
                        
                        If PremisseDecodee = "" Then
                
                            '***********************************************************************************
                            '                                        La pr�misse n'existe pas
                            '***********************************************************************************
                            ConstruitProchainCyclePont = PREMISSE_INEXISTANTE
                           
                    
                        Else
                        
                            '***********************************************************************************
                            '                                  Extraction de la pr�misse d�cod�e
                            '***********************************************************************************
                       
                            '--- construction du tableau de la pr�misse d�cod�e ---
                            TPremisseDecodee = Split(PremisseDecodee, SEPARATEUR_PREMISSES)
                        
                            '--- affectation ---
                            PtrLigne = 0
                    
                            '--- transfert dans le tableau du cycle du pont ---
                            For a = LBound(TCyclePont()) To UBound(TCyclePont())
                                TCyclePont(a) = TPremisseDecodee(PtrLigne)
                                Inc PtrLigne
                                If PtrLigne > UBound(TPremisseDecodee) Then Exit For
                            Next a
                
                            '***********************************************************************************
                            '                                Construction du PROCHAIN CYCLE
                            '***********************************************************************************
                        
                            '--- affectation ---
                            PtrAction = 1
            
                            '--- affectation des valeurs ---
                            For a = 1 To NBR_LIGNES_CYCLES_PONTS
        
                                '--- transfert des valeurs dans le tableau ---
                                NumAction = TCyclePont(a)
                                
                                'Call Log("numaction construit pont:" + NumAction)
                    
                                If NumAction >= NUM_ACTION_NOP And NumAction <= NUM_ACTION_FCY Then
        
                                    If TActions(NumAction).ParametreOuiNon = True And a < NBR_LIGNES_CYCLES_PONTS Then
                            
                                        '--- action avec un param�tre ---
                                        With TEtatsPonts(NumPont).TCyclesPonts(CYCLES.C_PROCHAIN, PtrAction)
                                            .NumAction = NumAction
                                            .Parametre = TCyclePont(Succ(a))
                                            .EtatParametre = ""
                                        End With
                                        Inc a   'd�calage de l'index car le param�tre est d�j� enregistr�
                        
                                    Else
                            
                                        '--- action sans param�tre ---
                                        With TEtatsPonts(NumPont).TCyclesPonts(CYCLES.C_PROCHAIN, PtrAction)
                                            .NumAction = NumAction
                                            .Parametre = ""
                                            .EtatParametre = ""
                                        End With
                    
                                    End If
                    
                                    '--- incr�mentation du pointeur de l'action ---
                                    Inc PtrAction
                    
                                End If
                        
                            Next a
        
                            '--- affectation ---
                            ConstruitProchainCyclePont = OK
                    
                        End If
                    
                    End If
                            
                    '*****************************************************************************************************
                    '                                             Affectation des param�tres
                    '*****************************************************************************************************
                    With TEtatsPonts(NumPont).TParametresCyclesPonts(CYCLES.C_PROCHAIN)
                        .TypeCycle = TYPES_CYCLES.TC_TRANSFERT_CHARGE
                        .NumPosteDepart = NumPosteDepart
                        .NumPosteArrivee = NumPosteArrivee
                        .DelaiSupStabilisationChargeSecondes = 0
                        .TempsEgouttageSecondes = 0
                    End With
            
                Case Else
            
            End Select
        
        End If

    End If

End Function

'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' R�le      : Converti une pr�misse d�cod�e en pr�misse cod�e
' Entr�es :                   PremisseDecodee -> Contient la pr�misse d�cod�e (exemple 100-101-102)
' Retours : PremisseDecodeeVersCodee -> Contient la premisse cod�e (exemple NOP-A3-NB)
' D�tails  :
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function PremisseDecodeeVersCodee(ByVal PremisseDecodee As String) As String
                                                                                                       
    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next
                                                                                                       
    '--- constantes priv�es ---

    '--- d�claration ---
    Dim ParametreOuiNon As Boolean
    Dim a As Integer, _
            NumAction As Integer
    Dim PremisseCodee As String
    Dim TPremisseDecodee As Variant

    '--- affectation ---
    PremisseCodee = ""

    If PremisseDecodee <> "" Then

        '--- construction du tableau ---
        TPremisseDecodee = Split(PremisseDecodee, SEPARATEUR_PREMISSES)
        
        '--- construction de la pr�misse cod�e ---
        For a = LBound(TPremisseDecodee) To UBound(TPremisseDecodee)
                    
            If TPremisseDecodee(a) <> "" Then
                                
                '--- num�ro de l'action ---
                NumAction = TPremisseDecodee(a)
                                
                If NumAction >= LBound(TActions()) And NumAction <= UBound(TActions()) Then
                                                        
                    If ParametreOuiNon = False Then
                        PremisseCodee = PremisseCodee & TActions(NumAction).CodeAction & SEPARATEUR_PREMISSES
                        ParametreOuiNon = TActions(NumAction).ParametreOuiNon
                    Else
                        'ATTENTION Le num�ro de l'action est dans ce cas le param�tre
                        PremisseCodee = PremisseCodee & NumAction & SEPARATEUR_PREMISSES
                        ParametreOuiNon = False
                    End If
                                
                End If
        
            End If
        
        Next a

        '--- �limination du dernier s�parateur de la pr�misse cod�e ---
        If PremisseCodee <> "" Then
            PremisseCodee = Mid(PremisseCodee, 1, Pred(Len(PremisseCodee)))
        End If

    End If

    '--- valeur de retour ---
    PremisseDecodeeVersCodee = PremisseCodee

End Function

'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' R�le      : Recherche le num�ro de pont choisi dans la pr�misse pour un d�placement entre 2 zones
' Entr�es :                                         NumZoneDepart -> Num�ro de la zone de d�part
'                                                        NumZoneArrivee -> Num�ro de la zone d'arriv�e
' Retours : RechercheNumPontChoisiDansPremisse -> 0 = d�termination impossible (pr�misse inexistente par exemple)
'                                                                                        sinon le num�ro du pont
' D�tails  :
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function RechercheNumPontChoisiDansPremisse(ByVal NumZoneDepart As Integer, _
                                                                                             ByVal NumZoneArrivee As Integer) As Integer

    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next

    '--- constantes priv�es ---

    '--- d�claration ---
    Dim a As Integer                                                    'pour les boucles FOR...NEXT
    Dim b As Integer                                                    'pour les boucles FOR...NEXT
    Dim NumPosteDepart As Integer                           'num�ro du poste de d�part
    Dim NumPosteArrivee As Integer                          'num�ro du poste d'arriv�e

    '--- affectation par d�faut ---
    RechercheNumPontChoisiDansPremisse = 0
    
    If NumZoneDepart > 0 And NumZoneArrivee > 0 Then

        '--- v�rification de l'existence des pr�misses ---
        For a = TZones(NumZoneDepart).NumPremierPoste To TZones(NumZoneDepart).NumDernierPoste
            For b = TZones(NumZoneArrivee).NumPremierPoste To TZones(NumZoneArrivee).NumDernierPoste
    
                '--- affectation ---
                NumPosteDepart = a
                NumPosteArrivee = b
            
                '--- contr�le ---
                With TPremisses(NumPosteDepart, NumPosteArrivee)
                    
                    If .PremisseCodee = "" Then
                                
                        '--- la pr�misse n'existe pas alors sortie directe ---
                        RechercheNumPontChoisiDansPremisse = 0
                        Exit Function
                                
                    Else
                    
                        '--- affectation du num�ro de pont ---
                        If TEtatsPonts(PONTS.P_1).Condamnation = True Then
                            RechercheNumPontChoisiDansPremisse = PONTS.P_2
                        ElseIf TEtatsPonts(PONTS.P_2).Condamnation = True Then
                            RechercheNumPontChoisiDansPremisse = PONTS.P_1
                        Else
                            RechercheNumPontChoisiDansPremisse = .NumPontIA
                        End If
                    
                    End If
                
                End With
                                
            Next b
        Next a
                                
    End If

End Function

'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' R�le      : Calcul automatique d'une pr�misse d�cod�e en fonction des param�tres
' Entr�es :                                 NumPosteDepart -> Num�ro du poste de d�part
'                                                NumPosteArrivee -> Num�ro du poste d'arriv�e
' Retours : CalculAutomatiquePremisseDecodee -> Contient la premisse d�cod�e (exemple 100-101-102)
' D�tails  :
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function CalculAutomatiquePremisseDecodee(ByVal NumPosteDepart As Integer, _
                                                                                       ByVal NumPosteArrivee As Integer) As String
    
    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next

    '--- constantes priv�es ---

    '--- d�claration ---
    Dim PremisseDecodee As String

    '--- affectation ---
    PremisseDecodee = ""
    
    If NumPosteDepart >= POSTES.P_CHGT_1 And NumPosteDepart <= DERNIER_POSTE And _
       NumPosteArrivee >= POSTES.P_CHGT_1 And NumPosteArrivee <= DERNIER_POSTE Then
        
        
        '****************************************************************************************************************
        '*                                                                      PARTIE PRISE
        '****************************************************************************************************************
        
        
        '--- cas de l'�tuve (ordre d'arr�t) ---
        'If NumPosteDepart = POSTES.P_C37 Then ' 3 Or NumPosteDepart = POSTES.P_C38 Then 'SZB 20180406
        '    PremisseDecodee = PremisseDecodee & NumActionParCode(CODE_ARRET_SECHOIR).NumAction & SEPARATEUR_PREMISSES
        '    PremisseDecodee = PremisseDecodee & NumActionParCode(CODE_TEMPO).NumAction & SEPARATEUR_PREMISSES
        '    PremisseDecodee = PremisseDecodee & "3" & SEPARATEUR_PREMISSES
        'End If
    
        '--- arr�t de l'agitation au poste de prise ---
        If TEtatsPostes(NumPosteDepart).DefinitionPoste.PresenceAgitationBain = True Then
            PremisseDecodee = PremisseDecodee & NumActionParCode(CODE_ARRET_AGITATION).NumAction & SEPARATEUR_PREMISSES
            PremisseDecodee = PremisseDecodee & NumPosteDepart & SEPARATEUR_PREMISSES
        End If
    
        '--- ouverture des couvercles au poste de prise ---
        If TEtatsPostes(NumPosteDepart).DefinitionPoste.PresenceCouvercles = True Then
            PremisseDecodee = PremisseDecodee & NumActionParCode(CODE_OUVERTURE_COUVERCLES).NumAction & SEPARATEUR_PREMISSES
            PremisseDecodee = PremisseDecodee & NumPosteDepart & SEPARATEUR_PREMISSES
        End If
    
        '--- translation au poste de prise ---
        PremisseDecodee = PremisseDecodee & CStr(NumPosteDepart) & SEPARATEUR_PREMISSES
        PremisseDecodee = PremisseDecodee & NumActionParCode(CODE_TEMPO).NumAction & SEPARATEUR_PREMISSES
        PremisseDecodee = PremisseDecodee & CStr(TEMPS_MINI_STABILISATION_A_VIDE) & SEPARATEUR_PREMISSES
        
        '--- contr�le de l'arr�t de l'agitation du bain ---
        If TEtatsPostes(NumPosteDepart).DefinitionPoste.PresenceAgitationBain = True Then
            PremisseDecodee = PremisseDecodee & NumActionParCode("CTRLARRETAGIT").NumAction & SEPARATEUR_PREMISSES
            PremisseDecodee = PremisseDecodee & NumPosteDepart & SEPARATEUR_PREMISSES
        End If
        
        '--- contr�le de l'ouverture des couvercles du poste de prise ---
        If TEtatsPostes(NumPosteDepart).DefinitionPoste.PresenceCouvercles = True Then
            PremisseDecodee = PremisseDecodee & NumActionParCode(CODE_CONTROLE_COUVERCLES_OUVERTS).NumAction & SEPARATEUR_PREMISSES
            PremisseDecodee = PremisseDecodee & NumPosteDepart & SEPARATEUR_PREMISSES
        End If
        
        '--- contr�le de l'arr�t du redresseur du poste de prise ---
        If TEtatsPostes(NumPosteDepart).DefinitionPoste.PresenceRedresseur = True Then
            PremisseDecodee = PremisseDecodee & NumActionParCode("AARRETRED").NumAction & SEPARATEUR_PREMISSES
            PremisseDecodee = PremisseDecodee & NumPosteDepart & SEPARATEUR_PREMISSES
        End If

        '--- descente des accroches ---
        PremisseDecodee = PremisseDecodee & NumActionParCode(CODE_DESCENTE_ACCROCHES).NumAction & SEPARATEUR_PREMISSES
    
        '--- mont�e au niveau haut ---
        PremisseDecodee = PremisseDecodee & NumActionParCode(CODE_NIVEAU_HAUT).NumAction & SEPARATEUR_PREMISSES
        
        '--- temporisation au niveau d'�gouttage (si poste � �gouttage) ---
        If TEtatsPostes(NumPosteDepart).DefinitionPoste.AvecEgouttage = True Then
            PremisseDecodee = PremisseDecodee & NumActionParCode(CODE_TEMPO_EGOUTTAGE).NumAction & SEPARATEUR_PREMISSES
            PremisseDecodee = PremisseDecodee & NumActionParCode("NOP").NumAction & SEPARATEUR_PREMISSES
        End If
        
        '--- fermeture des couvercles du poste de prise---
        If TEtatsPostes(NumPosteDepart).DefinitionPoste.PresenceCouvercles = True Then
            PremisseDecodee = PremisseDecodee & NumActionParCode(CODE_FERMETURE_COUVERCLES).NumAction & SEPARATEUR_PREMISSES
            PremisseDecodee = PremisseDecodee & NumPosteDepart & SEPARATEUR_PREMISSES
        End If
        
        
        '****************************************************************************************************************
        '*                                                                   PARTIE DEPOSE
        '****************************************************************************************************************
        
        
        '--- ouverture des couvercles au poste de d�pose ---
        If TEtatsPostes(NumPosteArrivee).DefinitionPoste.PresenceCouvercles = True Then
            PremisseDecodee = PremisseDecodee & NumActionParCode(CODE_OUVERTURE_COUVERCLES).NumAction & SEPARATEUR_PREMISSES
            PremisseDecodee = PremisseDecodee & NumPosteArrivee & SEPARATEUR_PREMISSES
        End If
    
        '--- translation au poste de d�pose ---
        If NumPosteDepart <> NumPosteArrivee Then
            PremisseDecodee = PremisseDecodee & CStr(NumPosteArrivee) & SEPARATEUR_PREMISSES
            PremisseDecodee = PremisseDecodee & NumActionParCode(CODE_TEMPO_STABILISATION).NumAction & SEPARATEUR_PREMISSES
            PremisseDecodee = PremisseDecodee & CStr(TEMPS_MINI_STABILISATION_AVEC_CHARGE) & SEPARATEUR_PREMISSES
        End If
        
        '--- contr�le de l'arr�t de l'agitation du bain ---
        If TEtatsPostes(NumPosteArrivee).DefinitionPoste.PresenceAgitationBain = True Then
            PremisseDecodee = PremisseDecodee & NumActionParCode("CTRLARRETAGIT").NumAction & SEPARATEUR_PREMISSES
            PremisseDecodee = PremisseDecodee & NumPosteArrivee & SEPARATEUR_PREMISSES
        End If
    
        '--- contr�le de l'ouverture des couvercles du poste de d�pose ---
        If TEtatsPostes(NumPosteArrivee).DefinitionPoste.PresenceCouvercles = True Then
            PremisseDecodee = PremisseDecodee & NumActionParCode(CODE_CONTROLE_COUVERCLES_OUVERTS).NumAction & SEPARATEUR_PREMISSES
            PremisseDecodee = PremisseDecodee & NumPosteArrivee & SEPARATEUR_PREMISSES
        End If
        
        '--- attente de l'arr�t du redresseur (contr�le de l'arr�t) du poste de d�pose ---
        If TEtatsPostes(NumPosteArrivee).DefinitionPoste.PresenceRedresseur = True Then
            PremisseDecodee = PremisseDecodee & NumActionParCode("AARRETRED").NumAction & SEPARATEUR_PREMISSES
            PremisseDecodee = PremisseDecodee & NumPosteArrivee & SEPARATEUR_PREMISSES
        End If
        
        '--- descente au niveau bas ---
        PremisseDecodee = PremisseDecodee & NumActionParCode(CODE_NIVEAU_BAS).NumAction & SEPARATEUR_PREMISSES
        
        '--- mont�e des accroches ---
        PremisseDecodee = PremisseDecodee & NumActionParCode(CODE_MONTEE_ACCROCHES).NumAction & SEPARATEUR_PREMISSES
        
        '--- mise en marche de l'agitation ---
        If TEtatsPostes(NumPosteArrivee).DefinitionPoste.PresenceAgitationBain = True Then
            PremisseDecodee = PremisseDecodee & NumActionParCode(CODE_MARCHE_AGITATION).NumAction & SEPARATEUR_PREMISSES
            PremisseDecodee = PremisseDecodee & NumPosteArrivee & SEPARATEUR_PREMISSES
        End If
        
        '--- fermeture des couvercles ---
        If TEtatsPostes(NumPosteArrivee).DefinitionPoste.PresenceCouvercles = True Then
            PremisseDecodee = PremisseDecodee & NumActionParCode(CODE_FERMETURE_COUVERCLES).NumAction & SEPARATEUR_PREMISSES
            PremisseDecodee = PremisseDecodee & NumPosteArrivee & SEPARATEUR_PREMISSES
        End If
        
        '--- cas de l'�tuve (ordre de marche) --- 20180406
        'If NumPosteDepart = POSTES.P_C37 Then '3 Or NumPosteDepart = POSTES.P_C34 Then
        '    PremisseDecodee = PremisseDecodee & NumActionParCode(CODE_TEMPO).NumAction & SEPARATEUR_PREMISSES
        '    PremisseDecodee = PremisseDecodee & "3" & SEPARATEUR_PREMISSES
        '    PremisseDecodee = PremisseDecodee & NumActionParCode(CODE_MARCHE_SECHOIR).NumAction & SEPARATEUR_PREMISSES
        'End If
        
        '--- fin de cycle ---
        PremisseDecodee = PremisseDecodee & NumActionParCode("FCY").NumAction & SEPARATEUR_PREMISSES
        
    Else
    
        '--- valeur de retour ---
        'une chaine vide indique une erreur dans la construction de la pr�misse
        CalculAutomatiquePremisseDecodee = ""
    
    End If

    '--- valeur de retour ---
    If PremisseDecodee <> "" Then
        PremisseDecodee = Mid(PremisseDecodee, 1, Pred(Len(PremisseDecodee)))      '�limination du dernier s�parateur
        CalculAutomatiquePremisseDecodee = PremisseDecodee
    End If

End Function

'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' R�le      : Calcul le temps total d'un cycle (total du temps de chaque action) d'une pr�misse
' Entr�es :                   NumPosteDepart -> Num�ro du poste de d�part
'                                  NumPosteArrivee -> Num�ro du poste d'arriv�e
' Retours :           TempsCycleSecondes -> Temps du cycle de la pr�misse calcul�e
'                 CalculTempsCyclePremisse -> Contient le message du r�sultat du calcul
'                                                                    OK = Calcul effectu� avec succ�s
'                                                                    MAUVAIS_POSTE_DEPART_ARRIVEE = Mauvais poste de d�part ou arriv�e
'                                                                    PREMISSE_INEXISTANTE = la pr�misse n'existe pas (calcul impossible)
' D�tails  : Le temps calcul� des actions ne tient pas compte des temps d'�gouttage et de d�placement au poste
'                 de prise, par-contre le temps de d�placement au poste de d�pose est calcul�
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function CalculTempsCyclePremisse(ByVal NumPosteDepart As Integer, _
                                                                        ByVal NumPosteArrivee As Integer, _
                                                                        ByRef TempsCycleSecondes As Long) As String
    
    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next

    '--- d�claration ---
    Dim a As Integer, _
           PtrLigne As Integer, _
           NumPontIA As Integer, _
           NumPostePrise As Integer, _
           NumPosteDepose As Integer, _
           NumActionNiveauActuel As Integer, _
           NumActionNiveauDestination As Integer
    Dim TempsTranslation As Single, _
           TempsLevage As Single
    Dim PremisseDecodee As String
    Dim TPremisseDecodee As Variant           'tableau de base contenant les actions apr�s d�codage
    Dim TDetailsPremisses(1 To NBR_LIGNES_DETAILS_PREMISSES) As VarPremissesTempsCycle
    
    '--- affectation ---
    TempsCycleSecondes = 0
    
    If NumPosteDepart >= POSTES.P_CHGT_1 And NumPosteDepart <= DERNIER_POSTE And _
        NumPosteArrivee >= POSTES.P_CHGT_1 And NumPosteArrivee <= DERNIER_POSTE Then
    
        '--- recherche de la pr�misse d�cod�e et du n� de pont par I.A. ---
        PremisseDecodee = TPremisses(NumPosteDepart, NumPosteArrivee).PremisseDecodee
        NumPontIA = TPremisses(NumPosteDepart, NumPosteArrivee).NumPontIA
    
        If PremisseDecodee = "" Then
        
            '--- la pr�misse n'existe pas ---
            CalculTempsCyclePremisse = PREMISSE_INEXISTANTE
        
        Else
                
            '--- construction du tableau ---
            TPremisseDecodee = Split(PremisseDecodee, SEPARATEUR_PREMISSES)
            
            '--- affectation du niveau actuel du pont ---
            NumActionNiveauActuel = NUM_ACTION_NIVEAU_BAS
            
            '--- transfert des donn�es dans le tableau ---
            PtrLigne = 1
            For a = LBound(TPremisseDecodee) To UBound(TPremisseDecodee)
                
                With TDetailsPremisses(PtrLigne)
                    
                    If TPremisseDecodee(a) <> "" Then
                            
                        '--- num�ro de l'action ---
                        .NumAction = TPremisseDecodee(a)
                            
                        If .NumAction >= LBound(TActions()) And .NumAction <= UBound(TActions()) Then
                                                    
                            '--- remplir l'action ---
                            .CodeAction = TActions(.NumAction).CodeAction
                            .LibelleAction = TActions(.NumAction).LibelleAction
                            .ParametreOuiNon = TActions(.NumAction).ParametreOuiNon
                            
                            '--- contr�le sur le param�tre ---
                            If .ParametreOuiNon = False Then
                        
                                '--- action sans param�tre ---
                                .Parametre = ""
                                                    
                            Else
                        
                                '--- action avec param�tre ---
                                Inc a
                                If a <= UBound(TPremisseDecodee) Then
                                    .Parametre = TPremisseDecodee(a)
                                End If
                        
                            End If
                            
                            '***************************************************************************************
                            '*                                  Contr�le du n� de pont donn� par l'I.A.
                            '***************************************************************************************
                            
                            If NumPontIA = PONTS.P_1 Or NumPontIA = PONTS.P_2 Then
                            Else
                                NumPontIA = PONTS.P_1     'fixer le pont 1 comme r�f�rence si le n� du pont
                                                                              'donn� par IA est � 0 (CAS NORMALLEMENT IMPOSSIBLE)
                                                                                                          
                            End If
                            
                            '***************************************************************************************
                            '*                                             Analyse de la translation
                            '***************************************************************************************
                            
                            '--- analyse du temps pour le poste de d�pose apr�s m�morisation du poste de prise ---
                            If NumPostePrise > 0 And .NumAction >= POSTES.P_CHGT_1 And .NumAction <= DERNIER_POSTE Then
                                
                                '--- affectation ---
                                NumPosteDepose = .NumAction
                                
                                '--- recherche du temps sur le pont I.A si le temps est � 0 alors recherche sur le pont oppos� ---
                                TempsTranslation = TEtatsPonts(NumPontIA).TTempsMouvements.TTempsTranslation(NumPostePrise, NumPosteDepose)
                                If TempsTranslation > 0 Then
                                    TempsCycleSecondes = TempsCycleSecondes + TempsTranslation
                                Else
                                    If NumPontIA = PONTS.P_1 Then           'regarder le temps sur le pont oppos�
                                        TempsTranslation = TEtatsPonts(PONTS.P_2).TTempsMouvements.TTempsTranslation(NumPostePrise, NumPosteDepose)
                                    Else
                                        TempsTranslation = TEtatsPonts(PONTS.P_1).TTempsMouvements.TTempsTranslation(NumPostePrise, NumPosteDepose)
                                    End If
                                    If TempsTranslation > 0 Then
                                        TempsCycleSecondes = TempsCycleSecondes + TempsTranslation
                                    End If
                                End If
                                
                            End If
                            
                            '--- m�moris� le poste de prise (premier poste d�fini dans le cycle) ---
                            If NumPostePrise = 0 And .NumAction >= POSTES.P_CHGT_1 And .NumAction <= DERNIER_POSTE Then
                                NumPostePrise = .NumAction
                            End If
                            
                            '***************************************************************************************
                            '*                                              Analyse des niveaux
                            '***************************************************************************************
                             
                            If .NumAction >= NUM_ACTION_NIVEAU_BAS And .NumAction <= NUM_ACTION_NIVEAU_HAUT Then
                                
                                '--- affectation du niveau de destination ---
                                NumActionNiveauDestination = .NumAction
 
                                '--- calcul du temps en fonction des niveaux pour la MONTEE ---
                                If NumActionNiveauActuel = NUM_ACTION_NIVEAU_BAS And _
                                   NumActionNiveauDestination = NUM_ACTION_NIVEAU_INTERMEDIAIRE Then
                                   TempsLevage = TEtatsPonts(NumPontIA).TTempsMouvements.TempsMonteeBasVersIntermediaire
                                End If
                                If NumActionNiveauActuel = NUM_ACTION_NIVEAU_BAS And _
                                   NumActionNiveauDestination = NUM_ACTION_NIVEAU_HAUT Then
                                   TempsLevage = TEtatsPonts(NumPontIA).TTempsMouvements.TempsMonteeBasVersHaut
                                End If
                                
                                '--- calcul du temps en fonction des niveaux pour la DESCENTE ---
                                If NumActionNiveauActuel = NUM_ACTION_NIVEAU_HAUT And _
                                   NumActionNiveauDestination = NUM_ACTION_NIVEAU_BAS Then
                                   TempsLevage = TEtatsPonts(NumPontIA).TTempsMouvements.TempsDescenteHautVersBas
                                End If
                                If NumActionNiveauActuel = NUM_ACTION_NIVEAU_INTERMEDIAIRE And _
                                   NumActionNiveauDestination = NUM_ACTION_NIVEAU_BAS Then
                                   TempsLevage = TEtatsPonts(NumPontIA).TTempsMouvements.TempsDescenteIntermediaireVersBas
                                End If
                                
                                '--- ajout du temps ---
                                If TempsLevage > 0 Then

                                    '--- affectation ---
                                    TempsCycleSecondes = TempsCycleSecondes + TempsLevage
                                    TempsLevage = 0
                                
                                End If
                            
                                '--- niveau actuel = niveau de destination ---
                                NumActionNiveauActuel = NumActionNiveauDestination
                            
                            End If
                            
                            '***************************************************************************************
                            '*                                  Ouverture / Fermeture des accroches
                            '***************************************************************************************
                            If .CodeAction = CODE_MONTEE_ACCROCHES Then
                                TempsCycleSecondes = TempsCycleSecondes + TEtatsPonts(NumPontIA).TTempsMouvements.TempsAccrochesChargeVersHaut
                            End If
                            If .CodeAction = CODE_DESCENTE_ACCROCHES Then
                                TempsCycleSecondes = TempsCycleSecondes + TEtatsPonts(NumPontIA).TTempsMouvements.TempsAccrochesChargeVersBas
                            End If
                            
                            '***************************************************************************************
                            '*                                                Temporisation fixe
                            '***************************************************************************************
                            If .CodeAction = CODE_TEMPO And IsNumeric(.Parametre) = True Then
                                TempsCycleSecondes = TempsCycleSecondes + CSng(.Parametre)
                            End If
                            
                            '***************************************************************************************
                            '*                                      Temporisation de stabilisation
                            '***************************************************************************************
                            'dans le cas de la temporisation de stabilisation, on ne prend que le temps mini
                            'de stabilisation de la charge, le temps suppl�mentaire �tant ajout� au moment
                            'des calculs des temps de gamme en r�el sur la ligne
                            If .CodeAction = CODE_TEMPO_STABILISATION Then
                                TempsCycleSecondes = TempsCycleSecondes + TEMPS_MINI_STABILISATION_AVEC_CHARGE
                            End If
                            
                            '***************************************************************************************
                            '*                                                  Fin de cycle
                            '***************************************************************************************
                            If .CodeAction = CODE_FIN_DE_CYCLE Then
                                Exit For
                            End If
                            
                            '--- incr�ment de la ligne ---
                            Inc PtrLigne
                        
                        End If
                
                    End If
                    
                End With
            
            Next a
        
            '--- affectation ---
            ' � ce stade la calcul peut �tre consid�r� comme bon
            CalculTempsCyclePremisse = OK
        
        End If
    
    Else
    
        '--- affectation en mauvais poste de d�part ou d'arriv�e ---
        CalculTempsCyclePremisse = MAUVAIS_POSTE_DEPART_ARRIVEE
    
    End If
    
End Function

'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' R�le      : Recherche un poste par son nom
' Entr�es : NomPoste       -> Nom du poste
' Retours : PosteParNom -> Details du poste selon le type EnrPostes
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function PosteParNom(ByVal NomPoste As String) As EnrPostes
    
    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next

    '--- d�claration ---
    Dim a As Integer
    Dim PosteVide As EnrPostes
    
    '--- affectation ---
    PosteParNom = PosteVide
    
    '--- recherche du poste ---
    For a = LBound(TEtatsPostes()) To UBound(TEtatsPostes())
        If TEtatsPostes(a).DefinitionPoste.NomPoste = NomPoste Then
            PosteParNom = TEtatsPostes(a).DefinitionPoste
            Exit For
        End If
    Next a

End Function

'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' R�le      : Recherche une action par son code
' Entr�es :              CodeAction -> Code de l'action
' Retours : NumActionParCode -> Details de l'action selon le type EnrActions
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function NumActionParCode(ByVal CodeAction As String) As EnrActions
    
    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next

    '--- d�claration ---
    Dim a As Integer
    Dim ActionVide As EnrActions
    
    '--- affectation ---
    NumActionParCode = ActionVide
    
    '--- recherche du poste ---
    For a = LBound(TActions()) To UBound(TActions())
        If TActions(a).CodeAction = CodeAction Then
            NumActionParCode = TActions(a)
            Exit For
        End If
    Next a

End Function

'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' R�le      : Recherche des postes par un ensemble de noms
' Entr�es : NomsPostes       -> Noms des postes s�par�s par des virgules
' Retours : PostesParNoms -> Tableau contenant les d�tails des postes selon le type EnrPostes
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function PostesParNoms(ByVal NomsPostes As String) As EnrPostes()
    
    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next

    '--- d�claration ---
    Dim a As Integer
    Dim FicheVide As EnrPostes, _
            TPostes() As EnrPostes
    Dim TNomsPostes As Variant
    
    '--- construction du tableau des noms de postes ---
    TNomsPostes = Split(NomsPostes, SEPARATEUR_POSTES)
    
    If UBound(TNomsPostes) > 0 Then
    
        '--- red�claration du tableau des postes ---
        ReDim TPostes(Succ(LBound(TNomsPostes)) To Succ(UBound(TNomsPostes))) As EnrPostes
    
        '--- recherche des postes ---
        For a = LBound(TNomsPostes) To UBound(TNomsPostes)
            TPostes(Succ(a)) = PosteParNom(TNomsPostes(a))
        Next a

    Else
    
        '--- red�claration d'un poste vide ---
        ReDim TPostes(1) As EnrPostes
        TPostes(1) = FicheVide
    
    End If

    '--- tableau de retour ---
    PostesParNoms = TPostes()

End Function

'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' R�le      : Effectue la correspondance d'un poste avec le num�ro de cuve g�r� par l'API
' Entr�es :                                      NumPoste -> Num�ro du poste de recherche de la correspondance
' Retours : CorrespondancePostesCuvesAPI -> Contient le num�ro de la cuve g�r� par l'API
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function CorrespondancePostesCuvesAPI(ByVal NumPoste As Long) As Long
    
    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next

    '--- d�claration ---
    
    '--- affectation ---
    CorrespondancePostesCuvesAPI = 0
    
    '--- correspondance ---
    Select Case NumPoste
        Case POSTES.P_C00: CorrespondancePostesCuvesAPI = CUVES_REGULATION.C_C00
        Case POSTES.P_DEC: CorrespondancePostesCuvesAPI = CUVES_REGULATION.C_DEC
        'Case POSTES.P_SAT: CorrespondancePostesCuvesAPI = CUVES_REGULATION.C_SAT
        'Case POSTES.P_C02: CorrespondancePostesCuvesAPI = CUVES_REGULATION.C_C02
        'Case POSTES.P_C03: CorrespondancePostesCuvesAPI = CUVES_REGULATION.C_C03
        'Case POSTES.P_C05: CorrespondancePostesCuvesAPI = CUVES_REGULATION.C_C05
        'Case POSTES.P_C06: CorrespondancePostesCuvesAPI = CUVES_REGULATION.C_C06
        Case POSTES.P_C07: CorrespondancePostesCuvesAPI = CUVES_REGULATION.C_C07
        Case POSTES.P_C13: CorrespondancePostesCuvesAPI = CUVES_REGULATION.C_C13
        Case POSTES.P_C14: CorrespondancePostesCuvesAPI = CUVES_REGULATION.C_C14
        Case POSTES.P_C15: CorrespondancePostesCuvesAPI = CUVES_REGULATION.C_C15
        'Case POSTES.P_C16: CorrespondancePostesCuvesAPI = CUVES_REGULATION.C_C16
        'Case POSTES.P_C19: CorrespondancePostesCuvesAPI = CUVES_REGULATION.C_C19
        Case POSTES.P_C22: CorrespondancePostesCuvesAPI = CUVES_REGULATION.C_C22
        Case POSTES.P_C27: CorrespondancePostesCuvesAPI = CUVES_REGULATION.C_C27
        Case POSTES.P_C28: CorrespondancePostesCuvesAPI = CUVES_REGULATION.C_C28
        Case POSTES.P_C31: CorrespondancePostesCuvesAPI = CUVES_REGULATION.C_C31
        Case POSTES.P_C32: CorrespondancePostesCuvesAPI = CUVES_REGULATION.C_C32
        'Case POSTES.P_C33: CorrespondancePostesCuvesAPI = CUVES_REGULATION.C_C33
        'Case POSTES.P_C34: CorrespondancePostesCuvesAPI = CUVES_REGULATION.C_C34
        'Case POSTES.P_C35: CorrespondancePostesCuvesAPI = CUVES_API.C_C35
        'Case POSTES.P_C36: CorrespondancePostesCuvesAPI = CUVES_API.C_C36
        'Case POSTES.P_C37: CorrespondancePostesCuvesAPI = CUVES_REGULATION.C_C37
        'Case POSTES.P_C38: CorrespondancePostesCuvesAPI = CUVES_API.C_MAX
        Case Else
    End Select
    
End Function

'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' R�le      : Effectue la correspondance d'une cuve g�r� par l'API avec le num�ro de poste
' Entr�es :                                      NumCuve -> Num�ro de la cuve de recherche de la correspondance
' Retours : CorrespondanceCuvesAPIPostes -> Contient le num�ro du poste
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function CorrespondanceCuvesAPIPostes(ByVal NumCuve As Long) As Long
    
    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next

    '--- d�claration ---
    
    '--- affectation ---
    CorrespondanceCuvesAPIPostes = 0
    
    '--- correspondance ---
    If NumCuve >= CUVES_REGULATION.C_C00 And NumCuve <= DERNIERE_CUV_REGULATION Then
        
        Select Case NumCuve
            
            Case CUVES_REGULATION.C_C00: CorrespondanceCuvesAPIPostes = POSTES.P_C00                 'd�graissage
            'Case CUVES_API.C_SAT: CorrespondanceCuvesAPIPostes = POSTES.P_DEC                 'satinage S201
            Case CUVES_REGULATION.C_DEC: CorrespondanceCuvesAPIPostes = POSTES.P_SAT                 'futur d�capage
            'Case CUVES_API.C_C02: CorrespondanceCuvesAPIPostes = POSTES.P_C02                 'Reserve
            'Case CUVES_API.C_C03: CorrespondanceCuvesAPIPostes = POSTES.P_C03                 'rin�age soude
            'Case CUVES_API.C_C05: CorrespondanceCuvesAPIPostes = POSTES.P_C05                 'brillantage n�1
            'Case CUVES_API.C_C06: CorrespondanceCuvesAPIPostes = POSTES.P_C06                 'rin�age Mt brillantage
            Case CUVES_REGULATION.C_C07: CorrespondanceCuvesAPIPostes = POSTES.P_C07                 'brillantage n�2
            Case CUVES_REGULATION.C_C13: CorrespondanceCuvesAPIPostes = POSTES.P_C13                 'anodisation
            Case CUVES_REGULATION.C_C14: CorrespondanceCuvesAPIPostes = POSTES.P_C14                 'anodisation
            Case CUVES_REGULATION.C_C15: CorrespondanceCuvesAPIPostes = POSTES.P_C15                 'anodisation
            'Case CUVES_API.C_C16: CorrespondanceCuvesAPIPostes = POSTES.P_C16                 'anodisation
            'Case CUVES_API.C_C19: CorrespondanceCuvesAPIPostes = POSTES.P_C19                 'spectrocoloration
            Case CUVES_REGULATION.C_C22: CorrespondanceCuvesAPIPostes = POSTES.P_C22                 'coloration or
            Case CUVES_REGULATION.C_C27: CorrespondanceCuvesAPIPostes = POSTES.P_C27                 'impr�gnation � froid
            Case CUVES_REGULATION.C_C28: CorrespondanceCuvesAPIPostes = POSTES.P_C28                 'coloration noire
            Case CUVES_REGULATION.C_C31: CorrespondanceCuvesAPIPostes = POSTES.P_C31                 'colmatage chaud
            Case CUVES_REGULATION.C_C32: CorrespondanceCuvesAPIPostes = POSTES.P_C32                 'colmatage chaud
            'Case CUVES_API.C_C33: CorrespondanceCuvesAPIPostes = POSTES.P_C33        's�choir - poste 1
            'Case CUVES_API.C_MAX: CorrespondanceCuvesAPIPostes = POSTES.P_C34        '
            'Case CUVES_API.C_C35: CorrespondanceCuvesAPIPostes = POSTES.P_C35
            'Case CUVES_API.C_C36: CorrespondanceCuvesAPIPostes = POSTES.P_C36        '
            'Case CUVES_API.C_C37: CorrespondanceCuvesAPIPostes = POSTES.P_C37        '
            'Case CUVES_API.C_MAX: CorrespondanceCuvesAPIPostes = POSTES.P_C38        '


            Case Else
        End Select
    
    End If
    
End Function

'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' R�le      : Effectue la correspondance d'un redresseur avec le num�ro de cuves g�r� par l'API
' Entr�es :                                     NumRedresseur -> Num�ro du redresseur
' Retours : CorrespondanceRedresseursCuvesAPI -> Contient le num�ro de cuve
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function CorrespondanceRedresseursCuvesAPI(ByVal NumRedresseur As Long) As Long
    
    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next

    '--- d�claration ---
    
    '--- affectation ---
    CorrespondanceRedresseursCuvesAPI = 0
    
    '--- correspondance ---
    If NumRedresseur >= REDRESSEURS.R_C13 And NumRedresseur <= REDRESSEURS.R_C16 Then
        
        Select Case NumRedresseur
            
            Case REDRESSEURS.R_C13: CorrespondanceRedresseursCuvesAPI = CUVES_REGULATION.C_C13
            Case REDRESSEURS.R_C14: CorrespondanceRedresseursCuvesAPI = CUVES_REGULATION.C_C14
            Case REDRESSEURS.R_C15: CorrespondanceRedresseursCuvesAPI = CUVES_REGULATION.C_C15
            'Case REDRESSEURS.R_C16: CorrespondanceRedresseursCuvesAPI = CUVES_API.C_C16

            Case Else
        End Select
    
    End If
    
End Function

'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' R�le      : Effectue la correspondance d'un redresseur g�r� par l'API avec le num�ro de poste
' Entr�es :                                       NumRedresseur -> Num�ro de la cuve de recherche de la correspondance
' Retours : CorrespondanceRedresseursAPIPostes -> Contient le num�ro du poste
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function CorrespondanceRedresseursAPIPostes(ByVal NumRedresseur As Long) As Long
    
    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next

    '--- d�claration ---
    
    '--- affectation ---
    CorrespondanceRedresseursAPIPostes = 0
    
    '--- correspondance ---
    If NumRedresseur >= REDRESSEURS.R_C13 And NumRedresseur <= REDRESSEURS.R_C16 Then
        Select Case NumRedresseur
            Case REDRESSEURS.R_C13: CorrespondanceRedresseursAPIPostes = POSTES.P_C13
            Case REDRESSEURS.R_C14: CorrespondanceRedresseursAPIPostes = POSTES.P_C14
            Case REDRESSEURS.R_C15: CorrespondanceRedresseursAPIPostes = POSTES.P_C15
            Case REDRESSEURS.R_C16: CorrespondanceRedresseursAPIPostes = POSTES.P_C16
            Case Else
        End Select
    End If
    
End Function

'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' R�le      : Extrait une pr�misse d�cod�e dans un tableau compatible avec un cycle de pont en vue de la
'                 transmettre � l'automate
' Entr�es :              NumPosteDepart -> Num�ro du poste de d�part
'                             NumPosteArrivee -> Num�ro du poste d'arriv�e
'                                          NumPont -> Num�ro du pont donn�e par les diagrammes en cyclique
'                                       NumPontIA -> Num�ro du pont donn� par IA (la validation de la pr�misse � la cr�ation
'                                                              force la variable NumPontIA � la valeur par d�faut de la variable NumPont
'                     TempsCycleSecondes -> Temps de la pr�misse par apprentissage
'                                     TCyclePont() -> Tableau contenant le cycle du pont (pr�misse d�cod�e)
' Retours : ExtraitPremisseDecodee -> Contient le message du r�sultat du d�codage
'                                                               ""  = Mauvais poste de d�part ou arriv�e
'                                                              OK = Pr�misse d�cod�e correctement et transmise dans le tableau de cycle
'                                                              PREMISSE_INEXISTANTE = la pr�misse n'existe pas
' D�tails  :
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function ExtraitPremisseDecodee(ByVal NumPosteDepart As Integer, _
                                                                   ByVal NumPosteArrivee As Integer, _
                                                                   ByRef NumPont As Integer, _
                                                                   ByRef NumPontIA As Integer, _
                                                                   ByRef TempsCycleSecondes As Long, _
                                                                   ByRef TCyclePont() As Integer) As String
                                                                   
    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next

    '--- d�claration ---
    Dim a As Integer, _
           PtrLigne As Integer
    Dim PremisseDecodee As String
    Dim TPremisseDecodee As Variant           'tableau de base contenant les actions apr�s d�codage
    
    '--- contr�le ---
    If NumPosteDepart >= POSTES.P_CHGT_1 And NumPosteDepart <= DERNIER_POSTE And _
        NumPosteArrivee >= POSTES.P_CHGT_1 And NumPosteArrivee <= DERNIER_POSTE Then
    
        '--- recherche de la pr�misse et du num�ro de pont ---
        With TPremisses(NumPosteDepart, NumPosteArrivee)
            PremisseDecodee = .PremisseDecodee
            NumPontIA = .NumPontIA
            TempsCycleSecondes = .TempsCycleSecondes
        End With
    
        If PremisseDecodee <> "" Then
                
            '--- construction du tableau ---
            TPremisseDecodee = Split(PremisseDecodee, SEPARATEUR_PREMISSES)
    
            '--- affectation ---
            PtrLigne = 0
            
            '--- transfert dans le tableau du cycle du pont ---
            For a = LBound(TCyclePont()) To UBound(TCyclePont())
                TCyclePont(a) = TPremisseDecodee(PtrLigne)
                Inc PtrLigne
                If PtrLigne > UBound(TPremisseDecodee) Then Exit For
            Next a
        
            '--- affectation ---
            ExtraitPremisseDecodee = OK
        
        Else
        
            '--- affectation ---
            ExtraitPremisseDecodee = PREMISSE_INEXISTANTE
    
        End If
    
    Else
    
        '--- affectation ---
        ExtraitPremisseDecodee = ""
    
    End If
    
End Function




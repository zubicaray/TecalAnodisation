Attribute VB_Name = "MSynoptique"
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Rôle                    : MODULE AIDANT A LA GESTION DU SYNOPTIQUE
' Nom                    : MSynoptique.bas
' Date de création : 17/12/2001
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

'--- déclarations obligatoires ---
Option Explicit

'--- options générales ---
Option Base 1
DefVar A-Z

'--- énumérations privées ---

'--- constantes publiques ---

'constantes communes pour la copie de l'image du synoptique dans les charges en ligne
Public Const LONGUEUR_IMAGE_LIGNE As Integer = 1877
Public Const HAUTEUR_IMAGE_LIGNE As Integer = 200

'--- énumérations publiques ---

'--- dimensions des ponts en coupe ---
Public Enum DIMENSIONS_ANIMATIONS
    
    D_LONG_ENSEMBLE_PONTS = 342
    D_LONG_PONT = 57
    D_HAUT_PONT = 72
    D_AXE_PONT = 28
    
    D_LONG_ENSEMBLE_PALONNIERS = 84
    D_LONG_PALONNIER = 21
    D_HAUT_PALONNIER = 17
    D_AXE_PALONNIER = 10
    
    D_LONG_ENSEMBLE_ACCROCHES = 36
    D_LONG_ACCROCHE = 9
    D_HAUT_ACCROCHE = 12
    D_AXE_ACCROCHE = 4
    
    D_LONG_ENSEMBLE_CHARGES = 230
    D_HAUT_ENSEMBLE_CHARGES = 2592
    D_NBR_COLONNES_ENSEMBLE_CHARGES = 10
    
    D_LONG_CHARGE = 23
    D_HAUT_CHARGE = 48
    D_AXE_CHARGE = 11
    
    D_LONG_ENSEMBLE_COUVERCLES = 31
    D_HAUT_ENSEMBLE_COUVERCLES = 65
    
    D_LONG_COUVERCLES = 31
    D_HAUT_COUVERCLES = 13
    D_AXE_COUVERCLES = 15
    
    D_LONG_CHARIOT = 29
    D_HAUT_CHARIOT = 55
    D_AXE_CHARIOT = 14

    D_LONG_ENSEMBLE_LIBELLES = 129
    D_HAUT_ENSEMBLE_LIBELLES = 855
    D_NBR_COLONNES_ENSEMBLE_LIBELLES = 3

    D_LONG_1_LIBELLE = 43                                   'il ya plusieurs longueurs de libellés
    D_LONG_2_LIBELLE = 30
    D_LONG_3_LIBELLE = 21
    D_HAUT_LIBELLE = 19                                        'hauteur des libellés constante

End Enum

'--- variables publiques ---

'variables communes pour la copie de l'image du synoptique dans les charges en ligne
Public ObjDDSImageTampon As DirectDrawSurface7                        'objet de l'image du tampon
Public DDSDImageTampon  As DDSURFACEDESC2                          'description de la surface de l'image du tampon
Public RImageTampon As RECT                                                         'coordonnées du rectangle de l'image du tampon

'--- tableaux publiques ---
Public TXPonts(PONTS.P_1 To PONTS.P_2) As Single                        'X des ponts
Public TDerniersXPonts(PONTS.P_1 To PONTS.P_2) As Single          'derniers X des ponts
Public TYPonts(PONTS.P_1 To PONTS.P_2) As Single                        'Y des ponts
Public TDerniersYPonts(PONTS.P_1 To PONTS.P_2) As Single          'derniers Y des ponts

'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Rôle      : Visualise les libellés de tous les états de la ligne
' Détails  :
' Retours :
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Sub VisualisationLibellesEtatsLigne()
    
    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next

    '--- constantes privées ---
        
    '--- déclaration ---
    Dim a As Integer
    
    With OccFSynoptique
    
        '----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
        
        '--- libellés des postes ---
        For a = POSTES.P_CHGT_1 To DERNIER_POSTE
            
            '--- noms de postes ---
            With .LNomsPostes(a)
                .Caption = TEtatsPostes(a).DefinitionPoste.NomPoste
                .Refresh
            End With
            
            Select Case a
            
                Case POSTES.P_CHGT_1 To POSTES.P_CHGT_2, POSTES.P_D1 To POSTES.P_D2
                    '--- postes de chargement et déchargement ---
                
                Case Else
                    '--- les autres postes ---
                    With .LLibellesPostes(a)
                        .Caption = UN_ESPACE & UCase(TEtatsPostes(a).DefinitionPoste.LibellePoste)
                        .Refresh
                    End With
            
            End Select
        
        Next a
    
        '----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
        
        '--- libellés des annexes ---
        'With .LLibellesAnnexes(INDEX_CHAMPS.IDX_CHAMP_VENTILATION)
        '    .Caption = " VENTILATION"
        '    .Refresh
        'End With
        'With .LLibellesAnnexes(INDEX_CHAMPS.IDX_CHAMP_VOLET_COMPENSATION)
        '    .Caption = " VOLET de COMPENSATION"
        '    .Refresh
        'End With
        'With .LLibellesAnnexes(INDEX_CHAMPS.IDX_CHAMP_AIR_COMPRIME)
        '    .Caption = " AIR COMPRIME"
        '    .Refresh
        'End With
        'With .LLibellesAnnexes(INDEX_CHAMPS.IDX_CHAMP_SURPRESSEUR_AIR)
        '    .Caption = " SURPRESSEUR d'AIR"
        '    .Refresh
        'End With
        'With .LLibellesAnnexes(INDEX_CHAMPS.IDX_CHAMP_ROTATION_TONNEAU_CUVES)
        '    .Caption = " ROTATION TONNEAU dans CUVES"
        '    .Refresh
        'End With
    
    End With
    
End Sub

'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Rôle      : Construction d'un cadre en 3D pour l'esthétique de l'écran principal
' Détails  :
' Retours :
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Sub ConstructionCadre3D()
    
    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next

    '--- constantes privées ---
        
    '--- déclaration ---

    With OccFSynoptique
    
        .AutoRedraw = True
        
        OccFSynoptique.Line (0, 0)-(.ScaleWidth, 0), COULEURS.GRIS_3
        OccFSynoptique.Line (.ScaleWidth - 1, 1)-(.ScaleWidth - 1, .ScaleHeight), COULEURS.NOIR
        OccFSynoptique.Line (.ScaleWidth - 1, .ScaleHeight - 1)-(0, .ScaleHeight - 1), COULEURS.NOIR
        OccFSynoptique.Line (0, .ScaleHeight - 1)-(0, 0), COULEURS.GRIS_3
        
        OccFSynoptique.Line (1, 1)-(.ScaleWidth - 1, 1), COULEURS.BLANC
        OccFSynoptique.Line (.ScaleWidth - 2, 2)-(.ScaleWidth - 2, .ScaleHeight - 1), COULEURS.GRIS_3
        OccFSynoptique.Line (.ScaleWidth - 2, .ScaleHeight - 2)-(1, .ScaleHeight - 2), COULEURS.GRIS_3
        OccFSynoptique.Line (1, .ScaleHeight - 2)-(1, 1), COULEURS.BLANC
        
        OccFSynoptique.Line (2, 2)-(.ScaleWidth - 4, .ScaleHeight - 4), COULEURS.GRIS_2, B
        
        .AutoRedraw = False
    
    End With
    
End Sub

'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Rôle      : Permet le condamnation d'un pont
' Détails  : NumPont -> Numéro du pont fonction de l'énumération PONTS
' Retours :
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Sub CondamnationPont(ByVal NumPont As PONTS, _
                                                  ByVal TitreMessages As String)

    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next

    '--- demande de confirmation avant la condamnation ---
    With TEtatsPonts(NumPont)
        
        If .Condamnation = False Then
            
            If AppelFenetre(F_MESSAGE, _
                                    TitreMessages, _
                                    vbCrLf & _
                                    "Cette action bloquera tous les mouvements du pont " & NumPont & vbCrLf & _
                                    "en AUTOMATIQUE." & vbCrLf & vbCrLf & _
                                    "c|Voulez-vous réellement CONDAMNE le PONT " & NumPont & " ?", _
                                    TYPES_MESSAGES.T_ATTENTION, _
                                    TYPES_BOUTONS.T_OUI_NON, _
                                    EMPLACEMENT_FOCUS.E_SUR_NON) = vbYes Then
                
                '--- condamnation du pont ---
                .Condamnation = True
            
            End If
                        
         Else
            
            If AppelFenetre(F_MESSAGE, _
                                    TitreMessages, _
                                    vbCrLf & vbCrLf & vbCrLf & _
                                    "c|Voulez-vous réellement DECONDAMNE le PONT " & NumPont & " ?", _
                                    TYPES_MESSAGES.T_AVERTISSEMENT, _
                                    TYPES_BOUTONS.T_OUI_NON, _
                                    EMPLACEMENT_FOCUS.E_SUR_NON) = vbYes Then
                
                '--- rétablissement  du pont ---
                .Condamnation = False
            
            End If
    
         End If

    End With

End Sub

'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Rôle      : Permet le condamnation d'un poste
' Détails  : NumPont -> Numéro du pont fonction de l'énumération POSTES
' Retours :
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Sub CondamnationPoste(ByVal NumPoste As POSTES, _
                                                    ByVal TitreMessages As String)

    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next

    '--- déclaration ---
    Dim NomPoste As String
    
    '--- demande de confirmation avant la condamnation ---
    With TEtatsPostes(NumPoste)
        
        '--- affectation ---
        NomPoste = TEtatsPostes(NumPoste).DefinitionPoste.NomPoste
        
        If .Condamnation = False Then
            
            If AppelFenetre(F_MESSAGE, _
                                    TitreMessages, _
                                    vbCrLf & _
                                    "Cette action bloquera tout les accès au poste " & NomPoste & vbCrLf & _
                                    "en AUTOMATIQUE." & vbCrLf & vbCrLf & _
                                    "c|Voulez-vous réellement CONDAMNE le POSTE " & NomPoste & " ?", _
                                    TYPES_MESSAGES.T_ATTENTION, _
                                    TYPES_BOUTONS.T_OUI_NON, _
                                    EMPLACEMENT_FOCUS.E_SUR_NON) = vbYes Then
                
                '--- condamnation du poste ---
                .Condamnation = True
            
                '--- écriture dans l'automate ---
                Bidon = APICondamnationDecondamnationPoste(NumPoste:=NumPoste, _
                                                                                               EtatSouhaite:=True)
            
            End If
                        
         Else
            
            If AppelFenetre(F_MESSAGE, _
                                    TitreMessages, _
                                    vbCrLf & vbCrLf & vbCrLf & _
                                    "c|Voulez-vous réellement DECONDAMNE le POSTE " & NomPoste & " ?", _
                                    TYPES_MESSAGES.T_AVERTISSEMENT, _
                                    TYPES_BOUTONS.T_OUI_NON, _
                                    EMPLACEMENT_FOCUS.E_SUR_NON) = vbYes Then
                
                '--- rétablissement  du poste ---
                .Condamnation = False
                
                '--- écriture dans l'automate ---
                Bidon = APICondamnationDecondamnationPoste(NumPoste:=NumPoste, _
                                                                                               EtatSouhaite:=False)
            
            End If
    
         End If

    End With

End Sub

'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Rôle      : Affiche les renseignements uniquement pour le mode renseignements
' Entrées : CouleurTexteRenseignements -> Couleur du texte des renseignements (énumération COULEURS)
'                             TexteRenseignements -> Texte à afficher
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Sub AfficheRenseignements(ByVal CouleurTexteRenseignements As COULEURS, _
                                                         ByVal TexteRenseignements As String)
    
    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next
    
    '--- sortie directe si mode différent des renseignements ---
    If OccFSynoptique.ModeDialoguesEnCours = MODES_DIALOGUES.M_RENSEIGNEMENTS Then
        Call OccFSynoptique.AfficheDialogues(CouleurTexteRenseignements, TexteRenseignements)
    End If

End Sub

'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Rôle      : Affiche uniquement les informations sur les entrées des charges
' Entrées : CouleurTexteRenseignements -> Couleur du texte des renseignements (énumération COULEURS)
'                             TexteRenseignements -> Texte à afficher
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Sub AfficheRenseignementsEntreesCharges(ByVal CouleurTexteRenseignements As COULEURS, _
                                                                                   ByVal TexteRenseignements As String)
    
    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next
    
    '--- sortie directe si mode différent de l'entrées des charges ---
    If OccFSynoptique.ModeDialoguesEnCours = MODES_DIALOGUES.M_ENTREE_CHARGES Then
        Call OccFSynoptique.AfficheDialogues(CouleurTexteRenseignements, TexteRenseignements)
    End If

End Sub

'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Rôle      : Affiche uniquement les informations sur le prévisionnel
' Entrées : CouleurTexteRenseignements -> Couleur du texte des renseignements (énumération COULEURS)
'                             TexteRenseignements -> Texte à afficher
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Sub AfficheRenseignementsPrevisionnel(ByVal CouleurTexteRenseignements As COULEURS, _
                                                                              ByVal TexteRenseignements As String)
    
    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next
    
    '--- sortie directe si mode différent du prévisonnel ---
    If OccFSynoptique.ModeDialoguesEnCours = MODES_DIALOGUES.M_PREVISIONNEL Then
        Call OccFSynoptique.AfficheDialogues(CouleurTexteRenseignements, TexteRenseignements)
    End If

End Sub

'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Rôle      : Affiche uniquement les informations de déboguage sur les entrées des charges
' Entrées : CouleurTexteRenseignements -> Couleur du texte des renseignements (énumération COULEURS)
'                             TexteRenseignements -> Texte à afficher
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Sub AfficheRenseignementsDebug(ByVal CouleurTexteRenseignements As COULEURS, _
                                                                    ByVal TexteRenseignements As String)
    
    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next
    'Call Log(TexteRenseignements)
    '--- sortie directe si mode questions réponses ---
    If OccFSynoptique.ModeDialoguesEnCours = MODES_DIALOGUES.M_ENTREE_CHARGES Then
        Call OccFSynoptique.AfficheDialogues(CouleurTexteRenseignements, TexteRenseignements)
    End If

End Sub

'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Rôle      : Afficher les réponses en mode questions réponses
' Entrées : CouleurTexteReponses -> Couleur du texte des réponses (énumération COULEURS)
'                             TexteReponses -> Texte à afficher
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Sub AfficheReponses(ByVal CouleurTexteReponses As COULEURS, _
                                               ByVal TexteReponses As String)
    
    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next
    
    '--- sortie directe si mode questions réponses ---
    If OccFSynoptique.ModeDialoguesEnCours = MODES_DIALOGUES.M_QUESTIONS_REPONSES Then
        Call OccFSynoptique.AfficheDialogues(CouleurTexteReponses, TexteReponses)
    End If

End Sub




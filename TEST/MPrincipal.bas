Attribute VB_Name = "MPrincipal"
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' R�le                    : MODULE PRINCIPAL
' Nom                    : MPrincipal.bas
' Date de cr�ation : 23/03/1999
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

'--- d�clarations obligatoires ---
Option Explicit

'--- options g�n�rales ---
Option Base 1
DefVar A-Z

'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' R�le     : Permet l'appel de la calculatrice de Windows
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Sub AppelCalculatrice()
    
    '--- aiguillage en cas d'erreurs ---
    On Error GoTo GestionErreurs
    
    '--- d�claration ---
    Static Identificateur As Variant
    
    '--- activation de l'application ---
    AppActivate Identificateur, True 'False
    
    Exit Sub

'--- gestion des erreurs ---
GestionErreurs:
    On Error Resume Next
    Identificateur = Shell("C:\WINDOWS\CALC.EXE", vbMaximizedFocus) 'vbHide) 'vbNormalNoFocus)

End Sub

'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' R�le      : Permet l'appel de la calculatrice de Windows
' Entr�es :  TexteAuFormatRTF -> Texte format� au format RTF
' Retours : ExtraitTexteSurRTF -> Tetxte sans formatage RTF
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function ExtraitTexteSurRTF(ByVal TexteAuFormatRTF As String) As String
    
    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next
    
    '--- extraction du texte ---
    With OccFPrincipale.RTBTampon
        .TextRTF = TexteAuFormatRTF
        ExtraitTexteSurRTF = .Text
    End With

End Function

'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' R�le      : Effectue le copier ou coller sp�cial
' Entr�es : CopierOuColler -> FALSE = Copie sp�ciale
'                                               TRUE = Collage sp�ciale
' D�tails  :
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Sub CopieCollageSpecial(ByVal CopieOuCollage As Boolean)

    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next

    '--- d�claration ---
    Dim NumfenetreActive As Long
    Dim ChaineTampon As String
    
    '--- affectation ---
    NumfenetreActive = OccFPrincipale.ActiveForm.NumFenetre
    If CopieOuCollage = False Then
        NumFenetreEnCopie = NumfenetreActive
    End If
    
    '--- appel de la fen�tre ---
    Select Case NumfenetreActive

        Case FENETRES.F_PREMISSES
            '--- pr�misses ---
            With OccFPrincipale.ActiveForm
                If CopieOuCollage = False Then
                    With .ADODCPremisses.Recordset
                        If Not .BOF And Not .EOF Then
                            CleDeCopie = .Fields("PremisseDecodee").Value
                        End If
                     End With
                Else
                    Call .GestionCopierCollerSpecial
                End If
            End With

        Case Else
            '--- aucune fen�tre active ---
            If CopieOuCollage = False Then
                NumFenetreEnCopie = 0
            End If

    End Select

End Sub

'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' R�le     : Proc�dure de d�marrage
' D�tails :
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Sub Main()
    
    '--- aiguillage en cas d'erreur ---
    On Error Resume Next
    
    'Dim d As Date
    'd = CDate("25/01/2022 3:6")
    'Dim heureIn As String
   
    'heureIn = Format(CStr(d), "hhnn")
    'heureIn = Format(d, "hhmm")
    
    
    '--- fen�tre d'acceuil ---
    If PROGRAMME_TERMINE = True Then
        FAcceuil.Show vbModal
        Unload FAcceuil
        Set FAcceuil = Nothing
    End If
    
    '--- fen�tre principal ---
    OccFPrincipale.Show vbModeless

End Sub

'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' R�le      : Initialisation des variables de l'ensemble du programme
' Retours :
' D�tails  :
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Sub InitialisationVariables()
    
    '--- aiguillage en cas d'erreur ---
    On Error Resume Next

    '--- d�claration ---
    
    '--- forcer la variable d'arr�t g�n�rale � VRAI de la ligne en attendant
    '    la premi�re communication avec l'automate
    TEtatsLigne.ArretGeneral = True
    
    '--- nom de l'ordinateur et de l'utilisateur ---
    NOM_ORDINATEUR = API_NomOrdinateur()
    NOM_UTILISATEUR = API_NomUtilisateur()
    NOM_ORDINATEUR_UTILISATEUR = NOM_ORDINATEUR & "|" & NOM_UTILISATEUR
    
    '--- affectation ---
    CARACTERE_PHI = Chr$(CODE_ASCII_PHI)
    CARACTERE_FRANC = "F"
    CARACTERE_EURO = Chr$(CODE_ASCII_EURO)

End Sub

'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' R�le      : Initialisation des images de l'ensemble du programme
' Retours :
' D�tails  :
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function InitialisationImages() As String
    
    '--- aiguillage en cas d'erreur ---
    On Error GoTo GestionErreurs
    
    '--- fonds de fentre ---
    Set ImgFondDeFenetre = LoadPicture(RepImagesAnodisation & "Fond de fen�tre.JPG")
    Set ImgFondDeFenetreXP = LoadPicture(RepImagesAnodisation & "Fond de fen�tre XP.JPG")
    Set ImgFondEspace = LoadPicture(RepImagesAnodisation & "Fond de l'espace.JPG")
    
    Set ImgFondOrange1 = LoadPicture(RepImagesAnodisation & "Fond orange 1.JPG")
    Set ImgFondOrange2 = LoadPicture(RepImagesAnodisation & "Fond orange 2.JPG")
    
    Set ImgFondBleu1 = LoadPicture(RepImagesAnodisation & "Fond bleu 1.JPG")
    Set ImgFondBleu2 = LoadPicture(RepImagesAnodisation & "Fond bleu 2.JPG")
    
    Set ImgFondVert1 = LoadPicture(RepImagesAnodisation & "Fond vert 1.JPG")
    Set ImgFondVert2 = LoadPicture(RepImagesAnodisation & "Fond vert 2.JPG")
    
    Set ImgFondGris1 = LoadPicture(RepImagesAnodisation & "Fond gris 1.JPG")
    Set ImgFondGris2 = LoadPicture(RepImagesAnodisation & "Fond gris 2.JPG")
    
    Set ImgFondDesBoutons = LoadPicture(RepImagesAnodisation & "Fond des boutons.JPG")
    
    '--- chargement des �chelles 24 heures ---
    Set TImgEchelles24H(ECHELLES_24H.E_CHAUFFAGE) = LoadPicture(RepImagesAnodisation & "Echelle 24 heures chauffage.BMP")
    Set TImgEchelles24H(ECHELLES_24H.E_POMPE_CHAUFFAGE) = LoadPicture(RepImagesAnodisation & "Echelle 24 heures pompe et chauffage.BMP")
    Set TImgEchelles24H(ECHELLES_24H.E_VENTILATION_CHAUFFAGE) = LoadPicture(RepImagesAnodisation & "Echelle 24 heures ventilation et chauffage.BMP")
    
    '--- chargement des images des redresseurs ---
    Set TRedresseursBMP(IMAGES_REDRESSEURS.I_BAS_REDRESSEUR_VERT) = LoadPicture(RepImagesAnodisation & "Bas redresseur vert.BMP")
    Set TRedresseursBMP(IMAGES_REDRESSEURS.I_BAS_REDRESSEUR_ORANGE) = LoadPicture(RepImagesAnodisation & "Bas redresseur orange.BMP")
    Set TRedresseursBMP(IMAGES_REDRESSEURS.I_BAS_REDRESSEUR_BLANC) = LoadPicture(RepImagesAnodisation & "Bas redresseur blanc.BMP")
    Set TRedresseursBMP(IMAGES_REDRESSEURS.I_BAS_REDRESSEUR_ROUGE) = LoadPicture(RepImagesAnodisation & "Bas redresseur rouge.BMP")
    Set TRedresseursBMP(IMAGES_REDRESSEURS.I_BAS_REDRESSEUR_EXCLUS) = LoadPicture(RepImagesAnodisation & "Bas redresseur exclus.BMP")

    '--- chargement des images des redresseurs en mode zoom ---
    Set TRedresseursZoomBMP(IMAGES_REDRESSEURS.I_BAS_REDRESSEUR_VERT) = LoadPicture(RepImagesAnodisation & "Bas redresseur vert zoom.BMP")
    Set TRedresseursZoomBMP(IMAGES_REDRESSEURS.I_BAS_REDRESSEUR_ORANGE) = LoadPicture(RepImagesAnodisation & "Bas redresseur orange zoom.BMP")
    Set TRedresseursZoomBMP(IMAGES_REDRESSEURS.I_BAS_REDRESSEUR_BLANC) = LoadPicture(RepImagesAnodisation & "Bas redresseur blanc zoom.BMP")
    Set TRedresseursZoomBMP(IMAGES_REDRESSEURS.I_BAS_REDRESSEUR_ROUGE) = LoadPicture(RepImagesAnodisation & "Bas redresseur rouge zoom.BMP")
    Set TRedresseursZoomBMP(IMAGES_REDRESSEURS.I_BAS_REDRESSEUR_EXCLUS) = LoadPicture(RepImagesAnodisation & "Bas redresseur exclus zoom.BMP")
    
    Exit Function

'--- gestion des erreurs ---
GestionErreurs:
    InitialisationImages = CStr(Err.Number)

End Function

Public Function InIDE() As Boolean
  On Error Resume Next
  Debug.Print 0 / 0
  InIDE = Err.Number <> 0
End Function

'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' R�le      : Affectation des chemins
' Retours : "" indique aucun incident sinon le num�ro de l'erreur est renvoy�
' D�tails  :
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function AffectationChemins() As String
    
    '--- aiguillage en cas d'erreurs ---
    On Error GoTo GestionErreurs
  
    '--- affectation ---
    AffectationChemins = ""
    
    '--- affectation du m�me chemin dans tous les cas pour les images ---
    RepImagesAnodisation = App.Path & "\Images\"
    
    RepFicAnodisation = App.Path & "\Fichiers communs\"
    RepGraphesProductionLocal = App.Path & "\Graphes de production\"
    RepGraphesProductionServeur = App.Path & "\Graphes de production\"

   
   
    
    If Environ("ANODISATION_TEST") = 1 Then
        MsgBox ("ATTENTION: Mode TEST")
        RepFicAnodisation = App.Path & "\Fichiers communs\"
        RepGraphesProductionLocal = App.Path & "\Graphes de production\"
        RepGraphesProductionServeur = App.Path & "\Graphes de production\"
               
       
    Else
        RepFicAnodisation = "D:\Fichiers communs de l'ANODISATION\"
        RepGraphesProductionLocal = "D:\Graphes de production\"
        RepGraphesProductionServeur = "C:\Anodisation\Graphes de production\"
       

    End If
    
    If FolderExists(RepGraphesProductionServeur) = False Then
        AffectationChemins = "Le dossier " & RepGraphesProductionServeur & " n'existe pas."
    End If
    
    If FolderExists(RepFicAnodisation) = False Then
        AffectationChemins = "Le dossier " & RepFicAnodisation & " n'existe pas."
    End If
    
    If FolderExists(RepGraphesProductionLocal) = False Then
        AffectationChemins = "Le dossier " & RepGraphesProductionLocal & " n'existe pas."
    End If
    
    If FolderExists(RepImagesAnodisation) = False Then
        AffectationChemins = "Le dossier " & RepImagesAnodisation & " n'existe pas."
    End If
    
    


    Exit Function

GestionErreurs:
    AffectationChemins = CStr(Err.Number)

End Function

'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' R�le      : Affiche la date et l'heure sur la barres des t�ches
' D�tails :
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Sub AfficheDateHeure()
    
    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next

    '--- affectation ---
    HeureSysteme = FormatDateTime(Time, vbLongTime)
    DateSysteme = StrConv(FormatDateTime(Date, vbLongDate), vbProperCase)
    
    '--- affichage de la date et heure ---
    If HeureSysteme <> MemHeureSysteme Or DateSysteme <> MemDateSysteme Then
        
        '--- affichage de la date et heure ---
        With OccFPrincipale.LDateHeureSysteme
            .Caption = DateSysteme & " - " & HeureSysteme
            .Refresh
        End With
        
        '--- affectation ---
        MemDateSysteme = DateSysteme
        MemHeureSysteme = HeureSysteme
    
    End If

End Sub

'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' R�le        : Affiche le type de t�che en cours
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Sub AfficheTypeTache(ByVal LibelleTache As String)
    
    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next
   
    '--- d�claration ---
    Static MemLibelleTache As String
   
    '--- affichage ---
    If LibelleTache <> MemLibelleTache Then
        If OccFPrincipale.LMessages.BackColor <> ROUGE_DEFAUT Then
            OccFPrincipale.LMessages.Caption = LibelleTache
            OccFPrincipale.LMessages.Refresh
            MemLibelleTache = LibelleTache
        End If
    End If

End Sub

'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' R�le      : Permet l'appel d'une fen�tre n'importe o� dans le programme
' Entr�es : NumFenetreActivation -> num�ro de la fenetre lorsqu'elle deviendra active
'                     ParametresFenetre -> Param�tres � transmettre � la fenetre
' Retours :               AppelFenetre -> Repr�sente une valeur retourn�e pour certaines fenetres
' D�tails  :
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function AppelFenetre(ByVal NumFenetreActivation As FENETRES, _
                                                Optional ByVal Parametre1 As Variant, _
                                                Optional ByVal Parametre2 As Variant, _
                                                Optional ByVal Parametre3 As Variant, _
                                                Optional ByVal Parametre4 As Variant, _
                                                Optional ByVal Parametre5 As Variant, _
                                                Optional ByVal Parametre6 As Variant, _
                                                Optional ByVal Parametre7 As Variant, _
                                                Optional ByVal Parametre8 As Variant, _
                                                Optional ByVal Parametre9 As Variant) As Variant

    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next
    
    '--- d�claration ---
    Dim FenetreTrouvee As Boolean
    Dim a As Integer
    Dim OccForm As Form, _
            OccFenetreAppele As Form
    Dim MemParametre1 As Variant, _
           MemParametre2 As Variant, _
           MemParametre3 As Variant, _
           MemParametre4 As Variant, _
           MemParametre5 As Variant, _
           MemParametre6 As Variant, _
           MemParametre7 As Variant, _
           MemParametre8 As Variant, _
           MemParametre9 As Variant

    '--- curseur de la souris ---
    Screen.MousePointer = vbHourglass
    
    '--- appel de la fen�tre ---
    Select Case NumFenetreActivation
        
        '-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
        
        Case FENETRES.F_SYNOPTIQUE
            '--- synoptique ---
            Load OccFSynoptique
            With OccFSynoptique
                .Left = 0
                .Top = 0
                .NumFenetre = NumFenetreActivation
                Call .InitialisationFenetre
                .Show vbModeless
                .SetFocus
            End With
            Screen.MousePointer = vbDefault
        
        '-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
        
        Case FENETRES.F_ORGANISATION_LIGNE
            '--- organisation de la ligne ---
            Load OccFOrganisationLigne
            With OccFOrganisationLigne
                .NumFenetre = NumFenetreActivation
                Call .InitialisationFenetre
                'Call .Parametragefenetre   'pas de param�trage de la fen�tre
                .Show vbModeless
                .SetFocus
            End With
            Screen.MousePointer = vbDefault
        
        '-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
        
        Case FENETRES.F_MOTEUR_INFERENCE
            '--- moteur d'inf�rence ---
            Load OccFMoteurInference
            With OccFMoteurInference
                .NumFenetre = NumFenetreActivation
                Call .InitialisationFenetre
                Call .ParametrageFenetre
                .Show vbModeless
                .SetFocus
            End With
            Screen.MousePointer = vbDefault
        
        '-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
        
        Case FENETRES.F_PREMISSES
            '--- pr�misses ---
            Load OccFPremisses
            With OccFPremisses
                .NumFenetre = NumFenetreActivation
                Call .InitialisationFenetre
                MemParametre1 = IIf(IsMissing(Parametre1) = True, 0, Parametre1)
                MemParametre2 = IIf(IsMissing(Parametre2) = True, 0, Parametre2)
                Call .ParametrageFenetre(NumPosteDepart:=MemParametre1, _
                                                         NumPosteArrivee:=MemParametre2)

                .Show vbModeless
                .SetFocus
            End With
            Screen.MousePointer = vbDefault
        
        '-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
        
        Case FENETRES.F_TEMPS_MOUVEMENTS
            '--- temps de mouvements ---
            Load OccFTempsMouvements
            With OccFTempsMouvements
                .NumFenetre = NumFenetreActivation
                Call .InitialisationFenetre
                Call .ParametrageFenetre
                .Show vbModeless
                .SetFocus
            End With
            Screen.MousePointer = vbDefault
        
        '-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
        
        Case FENETRES.F_MODE_CYCLIQUE
            '--- mode cyclique ---
            Load OccFModeCyclique
            With OccFModeCyclique
                .NumFenetre = NumFenetreActivation
                Call .InitialisationFenetre
                Call .ParametrageFenetre
                .Show vbModeless
                .SetFocus
            End With
            Screen.MousePointer = vbDefault
        
        '-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
        
        Case FENETRES.F_GAMMES_ANODISATION
            '--- gammes d'anodisation ---
            Load OccFGammesAnodisation
            With OccFGammesAnodisation
                .NumFenetre = NumFenetreActivation
                Call .InitialisationFenetre
                MemParametre1 = IIf(IsMissing(Parametre1) = True, False, Parametre1)
                MemParametre2 = IIf(IsMissing(Parametre2) = True, 1, Parametre2)
                MemParametre3 = IIf(IsMissing(Parametre3) = True, "", Parametre3)
                MemParametre4 = IIf(IsMissing(Parametre4) = True, "", Parametre4)
                Call .ParametrageFenetre(TravailSurGrille:=MemParametre1, _
                                                         RechercherPar:=MemParametre2, _
                                                         CommencantPar:=MemParametre3, _
                                                         Contenant:=MemParametre4)
                .Show vbModeless
                .SetFocus
            End With
            Screen.MousePointer = vbDefault
        
        '-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
        
        Case FENETRES.F_TRACABILITE_PRODUCTION
            '--- tra�abilit� de la production ---
            Load OccFTra�abiliteProduction
            With OccFTra�abiliteProduction
                .NumFenetre = NumFenetreActivation
                Call .InitialisationFenetre
                MemParametre1 = IIf(IsMissing(Parametre1) = True, False, Parametre1)
                MemParametre2 = IIf(IsMissing(Parametre2) = True, 1, Parametre2)
                MemParametre3 = IIf(IsMissing(Parametre3) = True, "", Parametre3)
                MemParametre4 = IIf(IsMissing(Parametre4) = True, "", Parametre4)
                Call .ParametrageFenetre(TravailSurGrille:=MemParametre1, _
                                                         RechercherPar:=MemParametre2, _
                                                         CommencantPar:=MemParametre3, _
                                                         Contenant:=MemParametre4)
                .Show vbModeless
                .SetFocus
            End With
            Screen.MousePointer = vbDefault
        
        '-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
        
        Case FENETRES.F_CHARGES_EN_LIGNE
            '--- charges en ligne ---
            Load OccFChargesEnLigne
            With OccFChargesEnLigne
                .NumFenetre = NumFenetreActivation
                Call .InitialisationFenetre
                MemParametre1 = IIf(IsMissing(Parametre1) = True, 0, Parametre1)
                Call .ParametrageFenetre(NumCharge:=MemParametre1)
                .Show vbModeless
                .SetFocus
            End With
            Screen.MousePointer = vbDefault
        
        '-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
        
        Case FENETRES.F_CYCLES_PONTS
            '--- cycles des ponts ---
            Load OccFCyclesPonts
            With OccFCyclesPonts
                .NumFenetre = NumFenetreActivation
                Call .InitialisationFenetre
                MemParametre1 = IIf(IsMissing(Parametre1) = True, FORMES_CYCLES_PONTS.F_CYCLES_PONTS_1_ET_2, Parametre1)
                Call .ParametrageFenetre(FormeCyclesPonts_:=MemParametre1)
                .Show vbModeless
                .SetFocus
            End With
            Screen.MousePointer = vbDefault
         
        '-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
         
         Case FENETRES.F_CHARGEMENT_PREVISIONNEL
            '--- chargement et pr�visonnel ---
            Load OccFChargementPrevisionnel
            With OccFChargementPrevisionnel
                .NumFenetre = NumFenetreActivation
                Call .InitialisationFenetre
                MemParametre1 = IIf(IsMissing(Parametre1) = True, ONGLETS_CHARGEMENT_PREVISIONNEL.O_CHARGEMENT, Parametre1)
                Call .ParametrageFenetre(OngletChoisie:=MemParametre1)
                .Show vbModeless
                .SetFocus
            End With
            Screen.MousePointer = vbDefault
        
        '-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
        
        Case FENETRES.F_GESTION_REDRESSEURS
            '--- gestion des redresseurs ---
            Load OccFGestionRedresseurs
            With OccFGestionRedresseurs
                .NumFenetre = NumFenetreActivation
                Call .InitialisationFenetre
                Call .ParametrageFenetre
                .Show vbModeless
                .SetFocus
            End With
            Screen.MousePointer = vbDefault
        
        '-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
        
        Case FENETRES.F_GESTION_CUVES
            '--- gestion des cuves ---
            Load OccFGestionCuves
            With OccFGestionCuves
                .NumFenetre = NumFenetreActivation
                Call .InitialisationFenetre
                MemParametre1 = IIf(IsMissing(Parametre1) = True, 1, Parametre1)
                Call .ParametrageFenetre(NumCuve:=MemParametre1)
                .Show vbModeless
                .SetFocus
            End With
            Screen.MousePointer = vbDefault
        
        '-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
        
        Case FENETRES.F_GESTION_REGULATION
            '--- gestion de la r�gulation ---
            Load OccFGestionRegulation
            With OccFGestionRegulation
                .NumFenetre = NumFenetreActivation
                Call .InitialisationFenetre
                MemParametre1 = IIf(IsMissing(Parametre1) = True, 1, Parametre1)
                Call .ParametrageFenetre(NumCuve:=MemParametre1)
                .Show vbModeless
                .SetFocus
            End With
            Screen.MousePointer = vbDefault
        
        '-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
        
        Case FENETRES.F_PROGRAMMATEUR_CYCLIQUE
            '--- programmateur cyclique ---
            Load OccFProgrammateurCyclique
            With OccFProgrammateurCyclique
                .NumFenetre = NumFenetreActivation
                Call .InitialisationFenetre
                'Call .Parametragefenetre   'pas de param�trage de la fen�tre
                .Show vbModeless
                .SetFocus
            End With
            Screen.MousePointer = vbDefault
        
        '-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
        
        Case FENETRES.F_ANNEXES
            '--- annexes ---
            Load OccFAnnexes
            With OccFAnnexes
                .NumFenetre = NumFenetreActivation
                Call .InitialisationFenetre
                'Call .Parametragefenetre   'pas de param�trage de la fen�tre
                .Show vbModeless
                .SetFocus
            End With
            Screen.MousePointer = vbDefault
        
        '-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
        
        Case FENETRES.F_VISUALISATION_GRAPHES_PRODUCTION
            '--- visualisation des graphes de production ---
            Load OccFVisualisationGraphesProduction
            With OccFVisualisationGraphesProduction
                .NumFenetre = NumFenetreActivation
                Call .InitialisationFenetre
                .Show vbModeless
                .SetFocus
            End With
            Screen.MousePointer = vbDefault
        
        '-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
        
        Case FENETRES.F_NETTOYAGE_GRAPHES_PRODUCTION
            '--- nettoyage des graphes de production ---
            Load OccFNettoyageGraphesProduction
            With OccFNettoyageGraphesProduction
                .NumFenetre = NumFenetreActivation
                Call .InitialisationFenetre
                .Show vbModeless
                .SetFocus
            End With
            Screen.MousePointer = vbDefault
        
        '-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
        
        'Case FENETRES.F_ALARMES
            '--- alarmes ---
            'Load OccFAlarmes
            'With OccFAlarmes
            '    .NumFenetre = NumFenetreActivation
            '    Call .InitialisationFenetre
            '    Call .ParametrageFenetre
            '    .Show vbModeless
            '    .SetFocus
            'End With
            'Screen.MousePointer = vbDefault
        
        '-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
        
        Case FENETRES.F_MAINTENANCE
            '--- maintenance ---
            Load OccFMaintenance
            With OccFMaintenance
                .NumFenetre = NumFenetreActivation
                Call .InitialisationFenetre
                'Call .Parametragefenetre   'pas de param�trage de la fen�tre
                .Show vbModeless
                .SetFocus
            End With
            Screen.MousePointer = vbDefault
        
        '-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
        
        Case FENETRES.F_INFORMATIONS_DEFAUTS_VARIATEURS
            '--- informations sur les d�fauts des variateurs ---
            Load OccFInformationsDefautsVariateurs
            With OccFInformationsDefautsVariateurs
                .NumFenetre = NumFenetreActivation
                Call .InitialisationFenetre
                Call .ParametrageFenetre
                .Show vbModeless
                .SetFocus
            End With
            Screen.MousePointer = vbDefault
        
        '-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
        
        Case FENETRES.F_INFORMATIONS_DEFAUTS_COMMUNICATION_AUTOMATE
            '--- informations sur les d�fauts de communication avec un automate ---
            Load OccFInformationsDefautsCommunicationAutomate
            With OccFInformationsDefautsCommunicationAutomate
                .NumFenetre = NumFenetreActivation
                Call .InitialisationFenetre
                Call .ParametrageFenetre
                .Show vbModeless
                .SetFocus
            End With
            Screen.MousePointer = vbDefault
        
        '-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
        
        'Case FENETRES.F_FIN_DE_JOURNEE
            '--- fin de journ�e ---
            'Load OccFFinDeJournee
            'With OccFFinDeJournee
                '.NumFenetre = NumFenetreActivation
                'Call .InitialisationFenetre
                'Call .Parametragefenetre   'pas de param�trage de la fen�tre
                '.Show vbModeless
                '.SetFocus
            'End With
            'Screen.MousePointer = vbDefault
        
        '-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
        
        Case FENETRES.F_ESSAIS
            '--- pour les essais ---
            Load OccFEssais
            With OccFEssais
                .NumFenetre = NumFenetreActivation
                Call .InitialisationFenetre
                MemParametre1 = IIf(IsMissing(Parametre1) = True, 0, Parametre1)
                Call .ParametrageFenetre
                .Show vbModeless
                .SetFocus
            End With
            Screen.MousePointer = vbDefault
        
        '-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
        
        Case FENETRES.F_APROPOS
            '--- � propos ---
            Load OccFAPropos
            With OccFAPropos
                .NumFenetre = NumFenetreActivation
                .Show vbModeless
                .SetFocus
            End With
            Screen.MousePointer = vbDefault
        
        '-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
        
        Case FENETRES.F_LISTE_DEFAUTS
            '--- liste des alarmes ---
            Load OccFListeDefauts
            With OccFListeDefauts
                .NumFenetre = NumFenetreActivation
                Call .InitialisationFenetre
                'Call .Parametragefenetre   'pas de param�trage de la fen�tre
                .Show vbModeless
                .SetFocus
            End With
            Screen.MousePointer = vbDefault
        
        '-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
        
        Case FENETRES.F_TRACABILITE_ALARMES
            '--- tra�abilit� des alarmes ---
            Load OccFTra�abiliteAlarmes
            With OccFTra�abiliteAlarmes
                .NumFenetre = NumFenetreActivation
                Call .InitialisationFenetre
                MemParametre1 = IIf(IsMissing(Parametre1) = True, 1, Parametre1)
                MemParametre2 = IIf(IsMissing(Parametre2) = True, "", Parametre2)
                MemParametre3 = IIf(IsMissing(Parametre3) = True, "", Parametre3)
                Call .ParametrageFenetre(RechercherPar:=MemParametre1, _
                                                         CommencantPar:=MemParametre2, _
                                                         Contenant:=MemParametre3)
                .Show vbModeless
                .SetFocus
            End With
            Screen.MousePointer = vbDefault
         
        '-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
        
        Case FENETRES.F_MODIFICATION_OPTIONS_CHARGE
            '--- modification des options d'une charge ---
            Load FModificationOptionsCharge
            With FModificationOptionsCharge
                .NumFenetre = NumFenetreActivation
                Call .InitialisationFenetre
                MemParametre1 = IIf(IsMissing(Parametre1) = True, "", Parametre1)
                Call .ParametrageFenetre(NumCharge:=MemParametre1)
                .Show vbModal
            End With
         
        '-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
         
        Case FENETRES.F_MESSAGE
            '--- message ---
            Load FMessage
            With FMessage
                .NumFenetre = NumFenetreActivation
                MemParametre1 = IIf(IsMissing(Parametre1) = True, "", Parametre1)
                MemParametre2 = IIf(IsMissing(Parametre2) = True, "", Parametre2)
                MemParametre3 = IIf(IsMissing(Parametre3) = True, 0, Parametre3)
                MemParametre4 = IIf(IsMissing(Parametre4) = True, 0, Parametre4)
                MemParametre5 = IIf(IsMissing(Parametre5) = True, 0, Parametre5)
                Call .ParametrageFenetre(TitreMessage:=MemParametre1, _
                                                         LibelleMessage:=MemParametre2, _
                                                         TypeMessage:=MemParametre3, _
                                                         TypesBoutons_:=MemParametre4, _
                                                         ChoixFocus_:=MemParametre5)
                .Show vbModal
                AppelFenetre = .VariableRetourneefenetre
            End With
            Unload FMessage
            Set FMessage = Nothing
        
        '-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
        
        Case FENETRES.F_MODIFICATIONS_AVANT_IMPRESSION
            '--- modifications avant impression ---
        
        Case Else

    End Select

End Function

'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' R�le     : Renseigne la fen�tre principale sur l'activit� des fen�tres filles
' D�tails :
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Sub RenseigneFPrincipale()

    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next
    
    '--- d�claration ---
    Dim NumfenetreActive As Long

    '--- curseur de la souris ---
    Screen.MousePointer = vbDefault
    
    If PremierPassageNoyauCentral = True Then
    
        With OccFPrincipale
            
            '--- num�ro de la fen�tre active ---
            NumfenetreActive = .ActiveForm.NumFenetre
            
            '--- appel de la Fenetre ---
            Select Case NumfenetreActive
               
                Case 0, FENETRES.F_SYNOPTIQUE
                    '--- outils du menu principal ---
                    AfficheOutilsMenuPrincipal O_STANDARD
                    AfficheOutilsMenuPrincipal O_MODE_IA_CYCLIQUE
                    AfficheOutilsMenuPrincipal O_PRODUCTION
                    
                    '--- indique un affichage complet des outils de la fen�tre principale ---
                    AffichageCompletOutilsFPrincipale = True
                
                Case FENETRES.F_GAMMES_ANODISATION
                    '--- op�rations � effectuer ---
                    'GereOutils .TOBOutilsPrincipaux, True, OUTILS_FENETRE_PRINCIPALE.O_APERCU_AVANT_IMPRESSION
                    'GereOutils .TOBOutilsPrincipaux, True, OUTILS_FENETRE_PRINCIPALE.O_IMPRIMER
                    'GereOutils .TOBOutilsPrincipaux, True, OUTILS_FENETRE_PRINCIPALE.O_IMPRIMER_FENETRE_ACTIVE
                    'GereOutils .TOBOutilsIA, True
                    'GereOutils .TOBOutilsProduction, True
            
                Case FENETRES.F_TRACABILITE_PRODUCTION
                    '--- op�rations � effectuer ---
                    'GereOutils .TOBOutilsPrincipaux, True, OUTILS_FENETRE_PRINCIPALE.O_APERCU_AVANT_IMPRESSION
                    'GereOutils .TOBOutilsPrincipaux, True, OUTILS_FENETRE_PRINCIPALE.O_IMPRIMER
                    'GereOutils .TOBOutilsPrincipaux, True, OUTILS_FENETRE_PRINCIPALE.O_IMPRIMER_FENETRE_ACTIVE
                    'GereOutils .TOBOutilsIA, True
                    'GereOutils .TOBOutilsProduction, True
            
                Case FENETRES.F_PREMISSES
                    '--- op�rations � effectuer ---
                    'GereOutils .TOBOutilsPrincipaux, False, OUTILS_FENETRE_PRINCIPALE.O_APERCU_AVANT_IMPRESSION
                    'GereOutils .TOBOutilsPrincipaux, False, OUTILS_FENETRE_PRINCIPALE.O_IMPRIMER
                    'GereOutils .TOBOutilsPrincipaux, True, OUTILS_FENETRE_PRINCIPALE.O_IMPRIMER_FENETRE_ACTIVE
                    'GereOutils .TOBOutilsIA, True
                    'GereOutils .TOBOutilsProduction, True
                
                Case FENETRES.F_APROPOS
                    '--- op�rations � effectuer ---
                    'GereOutils .TOBOutilsPrincipaux, False, OUTILS_FENETRE_PRINCIPALE.O_APERCU_AVANT_IMPRESSION
                    'GereOutils .TOBOutilsPrincipaux, False, OUTILS_FENETRE_PRINCIPALE.O_IMPRIMER
                    'GereOutils .TOBOutilsPrincipaux, False, OUTILS_FENETRE_PRINCIPALE.O_IMPRIMER_FENETRE_ACTIVE
                    'GereOutils .TOBOutilsIA, True
                    'GereOutils .TOBOutilsProduction, True
            
                Case Else
                    '--- aucune fen�tre active ---
                    'GereOutils .TOBOutilsPrincipaux, False, OUTILS_FENETRE_PRINCIPALE.O_APERCU_AVANT_IMPRESSION
                    'GereOutils .TOBOutilsPrincipaux, False, OUTILS_FENETRE_PRINCIPALE.O_IMPRIMER
                    'GereOutils .TOBOutilsPrincipaux, False, OUTILS_FENETRE_PRINCIPALE.O_IMPRIMER_FENETRE_ACTIVE
                    'GereOutils .TOBOutilsIA, True
                    'GereOutils .TOBOutilsProduction, True
    
            End Select
    
        End With

    End If

End Sub
    
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' R�le      : G�re des outils dans une barre d'outils
' Entr�es :   Barreoutils -> Barre d'outils de la fen�tre concern�e
'                EtatSouhaite -> FALSE = d�sactivation
'                                         TRUE = activation
'                IdxOutil         -> Index de l'outil (si pas d'index l'ensemble des outils sera activ� ou d�sactiv�)
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Sub GereOutils(ByRef BarreOutils As Toolbar, _
                                    ByVal EtatSouhaite As Boolean, _
                                    Optional ByVal IdxOutil As Variant)
    
    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next
 
    '--- d�claration ---
    Dim OccBouton As Button
     
    '--- activation ---
    If IsMissing(IdxOutil) = True Then
        For Each OccBouton In BarreOutils.buttons
            OccBouton.Enabled = EtatSouhaite
        Next
    Else
        
        '--- affectation ---
        BarreOutils.buttons(IdxOutil).Enabled = EtatSouhaite

        With OccFPrincipale
        
            '--- changement du menu ---
            Select Case BarreOutils.Name
        
                Case Else
    
            End Select
    
        End With
    
    End If
 
End Sub

'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' R�le      : Formater la date des messages
' Retours : DateMessages -> Date format�e
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function DateMessages() As String
    
    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next

    '--- affectation ---
    DateMessages = "(" & StrConv(Format(Date, "Long Date"), vbProperCase) & _
                                 " - " & Format(Time, "HH:MM:SS") & ")"

End Function

'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' R�le      : Effectue l'analyse des cycles des ponts
' Entr�es : NumPont -> N� du pont
' Retours :
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Sub AnalyseCyclesPonts(ByVal NumPont As PONTS)
    
    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next

    '--- constantes priv�es ---
    
    '--- d�claration ---
    Dim a As Integer, _
           PtrAction As Integer, _
           NumAction As Integer
    
    '--- affectation ---
    PtrAction = 1
    
    '--- affectation des valeurs ---
    For a = 1 To NBR_LIGNES_CYCLES_PONTS


        'Call Log("AnalyseCyclesPonts, ptrAction=" & PtrAction & " , numaction=" & NumAction)
        '--- transfert des valeurs dans le tableau ---
        NumAction = TImageAPICyclesPonts(NumPont, a)
            
        If NumAction >= NUM_ACTION_NOP And NumAction <= NUM_ACTION_FCY Then

            If TActions(NumAction).ParametreOuiNon = True And a < NBR_LIGNES_CYCLES_PONTS Then
                
                FSynoptique.TextInfo = "Action n�" & NumAction
                '--- action avec un param�tre ---
                With TEtatsPonts(NumPont).TCyclesPonts(CYCLES.C_ACTUEL, PtrAction)
                    .NumAction = NumAction
                    .Parametre = TImageAPICyclesPonts(NumPont, Succ(a))
                    .EtatParametre = ""
                End With
                Inc a   'd�calage de l'index car le param�tre est d�j� enregistr�
                
            Else
                FSynoptique.TextInfo = "Action 2 n�" & NumAction
                '--- action sans param�tre ---
                With TEtatsPonts(NumPont).TCyclesPonts(CYCLES.C_ACTUEL, PtrAction)
                    .NumAction = NumAction
                    .Parametre = ""
                    .EtatParametre = ""
                End With
            
            End If
            
            '--- incr�mentation du pointeur de l'action ---
            Inc PtrAction
            
        End If
                
    Next a

End Sub

'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' R�le      : Affiche les outils du menu principal
' Entr�es : OutilsChoisisMenuPrincipal -> Fonction de l'�num�ration OUTILS_MENU_PRINCIPAL
' Retours :
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Sub AfficheOutilsMenuPrincipal(ByVal OutilsChoisisMenuPrincipal As OUTILS_MENU_PRINCIPAL)

    '--- aiguillage en cas d'erreur ---
    On Error Resume Next

    With OccFPrincipale
    
        Select Case OutilsChoisisMenuPrincipal
    
            Case OUTILS_MENU_PRINCIPAL.O_STANDARD
                '--- outils standard ---
                AfficheBoutonOutils .TOBOutils(OutilsChoisisMenuPrincipal), 1, B_APERCU_AVANT_IMPRESSION
                AfficheBoutonOutils .TOBOutils(OutilsChoisisMenuPrincipal), 2, B_SEPARATEUR
            
                AfficheBoutonOutils .TOBOutils(OutilsChoisisMenuPrincipal), 3, B_CALCULATRICE
                AfficheBoutonOutils .TOBOutils(OutilsChoisisMenuPrincipal), 4, B_SEPARATEUR
            
            Case OUTILS_MENU_PRINCIPAL.O_MODE_IA_CYCLIQUE
                '--- outils pour la gestion du mode I.A. et du mode cyclique ---
                AfficheBoutonOutils .TOBOutils(OutilsChoisisMenuPrincipal), 1, B_ORGANISATION_LIGNE
                AfficheBoutonOutils .TOBOutils(OutilsChoisisMenuPrincipal), 2, B_SEPARATEUR
                
                AfficheBoutonOutils .TOBOutils(OutilsChoisisMenuPrincipal), 3, B_MOTEUR_INFERENCE
                AfficheBoutonOutils .TOBOutils(OutilsChoisisMenuPrincipal), 4, B_SEPARATEUR
                
                AfficheBoutonOutils .TOBOutils(OutilsChoisisMenuPrincipal), 5, B_MODE_CYCLIQUE
                AfficheBoutonOutils .TOBOutils(OutilsChoisisMenuPrincipal), 6, B_SEPARATEUR
            
            Case OUTILS_MENU_PRINCIPAL.O_PRODUCTION
                '--- outils de production ---
                AfficheBoutonOutils .TOBOutils(OutilsChoisisMenuPrincipal), 1, B_GAMMES_PRODUCTION
                AfficheBoutonOutils .TOBOutils(OutilsChoisisMenuPrincipal), 2, B_SEPARATEUR
                
                AfficheBoutonOutils .TOBOutils(OutilsChoisisMenuPrincipal), 3, B_TRACABILITE_PRODUCTION
                AfficheBoutonOutils .TOBOutils(OutilsChoisisMenuPrincipal), 4, B_SEPARATEUR
                
                AfficheBoutonOutils .TOBOutils(OutilsChoisisMenuPrincipal), 5, B_CHARGES_EN_LIGNE
                AfficheBoutonOutils .TOBOutils(OutilsChoisisMenuPrincipal), 6, B_SEPARATEUR
                
                AfficheBoutonOutils .TOBOutils(OutilsChoisisMenuPrincipal), 7, B_CYCLES_PONTS
                AfficheBoutonOutils .TOBOutils(OutilsChoisisMenuPrincipal), 8, B_SEPARATEUR
                
                AfficheBoutonOutils .TOBOutils(OutilsChoisisMenuPrincipal), 9, B_CHARGEMENT_PREVISIONNEL
                AfficheBoutonOutils .TOBOutils(OutilsChoisisMenuPrincipal), 10, B_SEPARATEUR
                
                AfficheBoutonOutils .TOBOutils(OutilsChoisisMenuPrincipal), 11, B_REDRESSEURS
                AfficheBoutonOutils .TOBOutils(OutilsChoisisMenuPrincipal), 12, B_SEPARATEUR
                
                AfficheBoutonOutils .TOBOutils(OutilsChoisisMenuPrincipal), 13, B_CUVES
                AfficheBoutonOutils .TOBOutils(OutilsChoisisMenuPrincipal), 14, B_SEPARATEUR
                
                AfficheBoutonOutils .TOBOutils(OutilsChoisisMenuPrincipal), 15, B_REGULATION
                AfficheBoutonOutils .TOBOutils(OutilsChoisisMenuPrincipal), 16, B_SEPARATEUR
                
                AfficheBoutonOutils .TOBOutils(OutilsChoisisMenuPrincipal), 17, B_PROGRAMMATEUR_CYCLIQUE
                AfficheBoutonOutils .TOBOutils(OutilsChoisisMenuPrincipal), 18, B_SEPARATEUR
                
                AfficheBoutonOutils .TOBOutils(OutilsChoisisMenuPrincipal), 19, B_ANNEXES
                AfficheBoutonOutils .TOBOutils(OutilsChoisisMenuPrincipal), 20, B_SEPARATEUR
                
                AfficheBoutonOutils .TOBOutils(OutilsChoisisMenuPrincipal), 21, B_LISTE_DEFAUTS
                AfficheBoutonOutils .TOBOutils(OutilsChoisisMenuPrincipal), 22, B_SEPARATEUR
                
                AfficheBoutonOutils .TOBOutils(OutilsChoisisMenuPrincipal), 23, B_MAINTENANCE
                AfficheBoutonOutils .TOBOutils(OutilsChoisisMenuPrincipal), 24, B_SEPARATEUR
                
                AfficheBoutonOutils .TOBOutils(OutilsChoisisMenuPrincipal), 25, B_FERMER_TOUT
                AfficheBoutonOutils .TOBOutils(OutilsChoisisMenuPrincipal), 26, B_SEPARATEUR
                
            Case Else
    
        End Select

    End With

End Sub

'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' R�le      : Affiche un bouton pr�d�fini dans une barres d'outils
' Entr�es :            BarreOutils -> Barre d'outils
'                            IdxBouton -> Index du bouton dans la barre d'outils
'               TypeBoutonOutils -> Fonction de l'�num�ration TYPES_BOUTONS_OUTILS
' Retours :
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Sub AfficheBoutonOutils(ByRef BarreOutils As Toolbar, _
                                                   ByVal IdxBouton As Integer, _
                                                   ByVal TypeBoutonOutils As TYPES_BOUTONS_OUTILS)

    '--- aiguillage en cas d'erreur ---
    On Error Resume Next
 
    '--- d�claration ---
    With BarreOutils.buttons(IdxBouton)
    
        Select Case TypeBoutonOutils
    
            Case TYPES_BOUTONS_OUTILS.B_VIDE
                '--- vide la barre d'outil � partir de l'index ---
                .Style = tbrDefault
                .Image = 0
                .Key = ""
                .Caption = ""
                .ToolTipText = ""
                .Visible = False
            
            Case TYPES_BOUTONS_OUTILS.B_SEPARATEUR
                '--- s�parateur ---
                .Style = tbrSeparator
                .Image = 0
                .Key = ""
                .Caption = ""
                .ToolTipText = ""
                .Visible = True
    
            '**********************************************************************************************************************************************
            '                                                                                                  MENU STANDARD
            '**********************************************************************************************************************************************
            
            Case TYPES_BOUTONS_OUTILS.B_APERCU_AVANT_IMPRESSION
                '--- aper�u avant impression ---
                .Style = tbrDropdown
                .Image = OccFPrincipale.ILOutils.ListImages("ApercuAvantImpression").Index
                .Key = "Aper�uAvantImpression"
                .Caption = "Aper�u"
                .ToolTipText = " Appel de l'�cran d'aper�u avant impression "
                
                '--- gestion des menus du bouton ---
                'menu 1 du bouton
                With .ButtonMenus(1)
                    .Key = "ImprimerDirectement"
                    .Text = "Imprimer directement"
                End With
                
                '--- gestion des menus du bouton ---
                'menu 2 du bouton -> s�parateur
                With .ButtonMenus(2)
                    .Key = ""
                    .Text = "-"
                End With
                
                '--- gestion des menus du bouton ---
                'menu 3 du bouton -> impression de la fen�tre active
                With .ButtonMenus(3)
                    .Key = "ImprimerFenetreActive"
                    .Text = "Imprimer la fen�tre active"
                End With
                
                .Visible = True
            
            Case TYPES_BOUTONS_OUTILS.B_CALCULATRICE
                '--- calculatrice ---
                .Style = tbrDefault
                .Image = OccFPrincipale.ILOutils.ListImages("Calculatrice").Index
                .Key = "Calculatrice"
                .Caption = "Calculatrice"
                .ToolTipText = " Appel de la calculatrice "
                .Visible = True
            
            '**********************************************************************************************************************************************
            '                                                                                MENU MODE I.A. ET MODE CYCLIQUE
            '**********************************************************************************************************************************************
            
            Case TYPES_BOUTONS_OUTILS.B_ORGANISATION_LIGNE
                '--- organisation de la ligne ---
                .Style = tbrDefault
                .Image = OccFPrincipale.ILOutils.ListImages("CaracteristiquesLigne").Index
                .Key = "OrganisationLigne"
                .Caption = "Organisation..."
                .ToolTipText = " Visualisation de l'organisation de la ligne "
                .Visible = True
            
            Case TYPES_BOUTONS_OUTILS.B_MOTEUR_INFERENCE
                '--- moteur d'inf�rence ---
                .Style = tbrDropdown
                .Image = OccFPrincipale.ILOutils.ListImages("MoteurInference").Index
                .Key = "MoteurInference"
                .Caption = "Moteur d'inf�rence"
                .ToolTipText = " Visualisation du moteur d'inf�rence "
                
                '--- gestion des menus du bouton ---
                'menu 1 du bouton -> pr�misses
                With .ButtonMenus(1)
                    .Key = "Premisses"
                    .Text = "Pr�misses"
                End With
                
                '--- gestion des menus du bouton ---
                'menu 2 du bouton -> s�parateur
                With .ButtonMenus(2)
                    .Key = ""
                    .Text = "-"
                End With
                
                '--- gestion des menus du bouton ---
                'menu 3 du bouton -> temps des mouvements
                With .ButtonMenus(3)
                    .Key = "TempsMouvements"
                    .Text = "Temps des mouvements"
                End With
                
                .Visible = True
            
            Case TYPES_BOUTONS_OUTILS.B_MODE_CYCLIQUE
                '--- mode cyclique ---
                .Style = tbrDefault
                .Image = OccFPrincipale.ILOutils.ListImages("ModeCyclique").Index
                .Key = "ModeCyclique"
                .Caption = "F2=Mode cyclique"
                .ToolTipText = " Visualisation du mode cyclique "
                .Visible = True
            
            '**********************************************************************************************************************************************
            '                                                                                          MENU PRODUCTION
            '**********************************************************************************************************************************************
            
            Case TYPES_BOUTONS_OUTILS.B_GAMMES_PRODUCTION
                '--- gammes de production ---
                .Style = tbrDefault
                .Image = OccFPrincipale.ILOutils.ListImages("GammesAnodisation").Index
                .Key = "GammesAnodisation"
                .Caption = "F3=Gammes"
                .ToolTipText = " Modification des gammes d'anodisation "
                .Visible = True
            
            Case TYPES_BOUTONS_OUTILS.B_TRACABILITE_PRODUCTION
                '--- tra�abilit� de production ---
                .Style = tbrDefault
                .Image = OccFPrincipale.ILOutils.ListImages("TracabiliteDeProduction").Index
                .Key = "Tracabilite"
                .Caption = "F4=Tra�abilit�"
                .ToolTipText = " Affiche la tra�abilit� des charges d�j� trait�es "
                .Visible = True
            
            Case TYPES_BOUTONS_OUTILS.B_CHARGES_EN_LIGNE
                '--- charges en ligne ---
                .Style = tbrDefault
                .Image = OccFPrincipale.ILOutils.ListImages("ChargesEnLigne").Index
                .Key = "ChargesEnLigne"
                .Caption = "F5=Charges..."
                .ToolTipText = " Visualise la totalit� des charges en ligne "
                .Visible = True
                
            Case TYPES_BOUTONS_OUTILS.B_CYCLES_PONTS
                '--- cycles des ponts ---
                .Style = tbrDefault
                .Image = OccFPrincipale.ILOutils.ListImages("CyclesPonts").Index
                .Key = "CyclesPonts"
                .Caption = "F6=Cycles..."
                .ToolTipText = " Visualisation des cycles des ponts"
                .Visible = True
            
            Case TYPES_BOUTONS_OUTILS.B_CHARGEMENT_PREVISIONNEL
                '--- chargement / pr�visionnel ---
                .Style = tbrDefault
                .Image = OccFPrincipale.ILOutils.ListImages("ChargementPrevisionnel").Index
                .Key = "ChargementPrevisionnel"
                .Caption = "F7=Chargement"
                .ToolTipText = " G�re l'entr�e des charges en ligne "
                .Visible = True
            
            Case TYPES_BOUTONS_OUTILS.B_REDRESSEURS
                '--- redresseurs ---
                .Style = tbrDefault
                .Image = OccFPrincipale.ILOutils.ListImages("Redresseurs").Index
                .Key = "Redresseurs"
                .Caption = "F8=Redresseurs"
                .ToolTipText = " G�re l'ensemble des redresseurs "
                .Visible = True
            
            Case TYPES_BOUTONS_OUTILS.B_CUVES
                '--- cuves ---
                .Style = tbrDefault
                .Image = OccFPrincipale.ILOutils.ListImages("Cuves").Index
                .Key = "Cuves"
                .Caption = "F9=Cuves"
                .ToolTipText = " G�re les cuves (niveaux, pompes, etc...) "
                .Visible = True
            
            Case TYPES_BOUTONS_OUTILS.B_REGULATION
                '--- r�gulation ---
                .Style = tbrDefault
                .Image = OccFPrincipale.ILOutils.ListImages("Regulation").Index
                .Key = "Regulation"
                .Caption = "F10=R�gulation"
                .ToolTipText = " G�re les temp�ratures des cuves "
                .Visible = True
            
            Case TYPES_BOUTONS_OUTILS.B_PROGRAMMATEUR_CYCLIQUE
                '--- programmateur cyclique ---
                .Style = tbrDefault
                .Image = OccFPrincipale.ILOutils.ListImages("ProgrammateurCyclique").Index
                .Key = "ProgrammateurCyclique"
                .Caption = "F11=Prog. cyclique"
                .ToolTipText = " G�re la programmation horaire des bains "
                .Visible = True
            
            Case TYPES_BOUTONS_OUTILS.B_ANNEXES
                '--- annexes ---
                .Style = tbrDefault
                .Image = OccFPrincipale.ILOutils.ListImages("Annexes").Index
                .Key = "Annexes"
                .Caption = "F12=Annexes"
                .ToolTipText = " Visualisation de l'�tat des annexes de la ligne "
                .Visible = True
            
            Case TYPES_BOUTONS_OUTILS.B_LISTE_DEFAUTS
                '--- liste des d�fauts ---
                .Style = tbrDefault
                .Image = OccFPrincipale.ILOutils.ListImages("Defauts").Index
                .Key = "Defauts"
                .Caption = "D�fauts"
                .ToolTipText = " Gestion des d�fauts "
                .Visible = True
            
            Case TYPES_BOUTONS_OUTILS.B_MAINTENANCE
                '--- maintenance ---
                .Style = tbrDefault
                .Image = OccFPrincipale.ILOutils.ListImages("Maintenance").Index
                .Key = "Maintenance"
                .Caption = "Maintenance"
                .ToolTipText = " Gestion de la maintenance "
                .Visible = True
            
            Case TYPES_BOUTONS_OUTILS.B_FERMER_TOUT
                '--- fermer toutes les fen�tres ---
                .Style = tbrDefault
                .Image = OccFPrincipale.ILOutils.ListImages("General").Index
                .Key = "Fermer tout"
                .Caption = "Fermer tout"
                .ToolTipText = " Fermeture de toutes les fen�tres "
                .Visible = True
            
            Case Else
        End Select
    
    End With
                
    '--- rendre visible le s�parateur ---
    If BarreOutils.buttons.Count > IdxBouton Then
        If TypeBoutonOutils <> B_VIDE Then
            BarreOutils.buttons(IdxBouton + 1).Visible = True
        Else
            BarreOutils.buttons(IdxBouton + 1).Visible = False
        End If
    End If
    
End Sub



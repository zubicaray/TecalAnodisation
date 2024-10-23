VERSION 5.00
Begin VB.Form FEssais 
   Caption         =   " ESSAIS"
   ClientHeight    =   12675
   ClientLeft      =   2130
   ClientTop       =   2280
   ClientWidth     =   23760
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   845
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   1584
   WindowState     =   2  'Maximized
   Begin VB.PictureBox PBBoutons 
      Align           =   2  'Align Bottom
      DrawStyle       =   2  'Dot
      DrawWidth       =   16891
      Height          =   1095
      Left            =   0
      ScaleHeight     =   1035
      ScaleWidth      =   23700
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   11580
      Width           =   23760
      Begin VB.PictureBox PBOutilsDeplacementFenetre 
         BackColor       =   &H00E0E0E0&
         Height          =   1035
         Left            =   0
         ScaleHeight     =   975
         ScaleWidth      =   1155
         TabIndex        =   6
         Top             =   0
         Visible         =   0   'False
         Width           =   1215
         Begin VB.HScrollBar HSDeplacementFenetre 
            Height          =   255
            LargeChange     =   300
            Left            =   0
            SmallChange     =   100
            TabIndex        =   9
            Top             =   720
            Width           =   915
         End
         Begin VB.VScrollBar VSDeplacementFenetre 
            Height          =   975
            LargeChange     =   300
            Left            =   900
            SmallChange     =   100
            TabIndex        =   8
            Top             =   0
            Width           =   255
         End
         Begin VB.CommandButton CBAgrandirFenetre 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Agrandir"
            DownPicture     =   "FEssais.frx":0000
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   720
            Left            =   0
            MaskColor       =   &H00FF00FF&
            Picture         =   "FEssais.frx":01AA
            Style           =   1  'Graphical
            TabIndex        =   7
            ToolTipText     =   " Agrandissement de la fenêtre "
            Top             =   0
            UseMaskColor    =   -1  'True
            Width           =   900
         End
      End
      Begin VB.Timer TimerEssais 
         Enabled         =   0   'False
         Interval        =   100
         Left            =   12120
         Top             =   120
      End
      Begin VB.CommandButton CBQuitter 
         BackColor       =   &H00FFFFFF&
         Cancel          =   -1  'True
         Caption         =   "&Quitter"
         DownPicture     =   "FEssais.frx":0354
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   840
         Left            =   13260
         MaskColor       =   &H00FF00FF&
         Picture         =   "FEssais.frx":0A56
         Style           =   1  'Graphical
         TabIndex        =   1
         ToolTipText     =   " Quitter cette fenêtre "
         Top             =   120
         UseMaskColor    =   -1  'True
         Width           =   1515
      End
      Begin VB.Shape SFocus 
         BorderColor     =   &H000000FF&
         BorderWidth     =   5
         Height          =   405
         Left            =   12660
         Top             =   120
         Visible         =   0   'False
         Width           =   420
      End
   End
   Begin VB.PictureBox PBRenseignementsFenetre 
      Align           =   1  'Align Top
      BackColor       =   &H00FF0000&
      Height          =   375
      Left            =   0
      Picture         =   "FEssais.frx":1158
      ScaleHeight     =   315
      ScaleWidth      =   23700
      TabIndex        =   4
      Top             =   0
      Width           =   23760
      Begin VB.Label LRenseignementsFenetre 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00008000&
         BackStyle       =   0  'Transparent
         Caption         =   "ESSAIS"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   240
         TabIndex        =   5
         Top             =   0
         Width           =   11415
         WordWrap        =   -1  'True
      End
   End
   Begin VB.PictureBox PBDeplacementFenetre 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   8355
      Index           =   0
      Left            =   0
      ScaleHeight     =   557
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   1584
      TabIndex        =   2
      Top             =   375
      Width           =   23760
      Begin VB.PictureBox PBDeplacementFenetre 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   6855
         Index           =   1
         Left            =   0
         ScaleHeight     =   457
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   1252
         TabIndex        =   3
         Top             =   0
         Width           =   18780
         Begin VB.PictureBox PBImageLigne 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   3615
            Left            =   600
            ScaleHeight     =   241
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   1141
            TabIndex        =   10
            Top             =   420
            Width           =   17115
         End
      End
   End
End
Attribute VB_Name = "FEssais"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Rôle                    : Fenêtre pour effectuer des essais
' Nom                    : FEssais.frm
' Date de création : 23/05/2011
' Détails                :
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

'--- déclarations obligatoires ---
Option Explicit

'--- options générales ---
Option Base 1
DefVar A-Z
    
'--- constantes privées ---
Private Const LONGUEUR_IMAGE_LIGNE As Integer = 1877
Private Const HAUTEUR_IMAGE_LIGNE As Integer = 200

Private Const TITRE_FENETRE As String = "ESSAIS"
Private Const TITRE_MESSAGES As String = TITRE_FENETRE

'--- énumérations privées ---

'--- variables privées ---
Private PremiereActivation As Boolean
Private MemDernierBouton As Long                'mémoire du dernier bouton

'--- variables et tableaux privées DIRECTX 7.0 ---
Private ObjDX As New DirectX7                                                          'objet DirectX
Private ObjDD As DirectDraw7                                                            'objet DirectDraw
        
Private ObjDDSEcran As DirectDrawSurface7                                     'objet de la surface de l'écran
Private DDSDEcran As DDSURFACEDESC2                                        'description de la surface de l'écran

Private ObjDDClip As DirectDrawClipper                                              'objet du clipper

Private ObjDDSImageLigne As DirectDrawSurface7                            'objet de la surface de l'image de la ligne
Private DDSDImageLigne As DDSURFACEDESC2                                'description de la surface de l'image de la ligne
Private RImageLigne As RECT                                                              'coordonnées du rectangle de l'image de la ligne


'--- tableaux privés ---

'--- variables publiques ---
Public NumFenetre As Long                               'numéro de la fenêtre lorsqu'elle devient active

'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Rôle      : Initialise la fenêtre (chargement ou en vue de la rendre visible)
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Sub InitialisationFenetre()
    
    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next

    '--- déclaration ---

    '--- affectation ---
  
    '--- divers sur la fenêtre ---
    With Me
        .Caption = TITRE_FENETRE
        .WindowState = vbMaximized
    End With
    PBDeplacementFenetre(ZONES_DEPLACEMENT_FENETRE.Z_FILLE).Picture = ImgFondOrange2
    PBBoutons.Picture = ImgFondDesBoutons
    
    '--- affectation ---
    
    '--- préparation de l'animation de la ligne ---
    InitialisationDirectX                          'initialisation de DirectX
    InitialisationSurfaces                        'Initialisation des surfaces
    'PremieresPositionsAnimations        'premières positions des animations
    
    '--- gestion de l'états des boutons ---
    GestionBoutons E_CHARGEMENT_FENETRE
    
End Sub

'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Rôle      : Effectue le paramètrage de la fenêtre
' Entrées : NumCharge -> Numéro de charge souhaité
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Sub ParametrageFenetre()
    
    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next

    '--- lancement du timer ---
    TimerEssais.Enabled = True

End Sub

'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Rôle      : Gére l'états des boutons après une action de l'opèrateur
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Private Sub GestionBoutons(ByVal Situation As ETATS_BOUTONS)
    
    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next

    Select Case Situation
        
        Case ETATS_BOUTONS.E_CHARGEMENT_FENETRE
            '--- au chargement de la fenetre ---
            CBQuitter.Enabled = True
        
        Case ETATS_BOUTONS.E_DECHARGEMENT_FENETRE
            '--- au déchargement de la fenêtre ---
        
        Case ETATS_BOUTONS.E_AVANT_VALIDER
            '--- avant valider ---
        
        Case ETATS_BOUTONS.E_APRES_VALIDER
            '--- après valider ---
            CBQuitter.Enabled = True

        Case ETATS_BOUTONS.E_AVANT_ANNULER
            '--- avant annuler ---
        
        Case ETATS_BOUTONS.E_APRES_ANNULER
            '--- après annuler ---
            CBQuitter.Enabled = True

        Case ETATS_BOUTONS.E_AVANT_ACTUALISER
            '--- avant actualiser ---
        
        Case ETATS_BOUTONS.E_APRES_ACTUALISER
            '--- après actualiser ---
            CBQuitter.Enabled = True
        
        Case ETATS_BOUTONS.E_MODIFICATION_EN_COURS
            '--- après modifier (à ne pas traiter si nouvel enregistrement) ---
            CBQuitter.Enabled = True

        Case ETATS_BOUTONS.E_AVANT_NOUVEAU
            '--- avant nouveau ---
        
        Case ETATS_BOUTONS.E_APRES_NOUVEAU
            '--- après nouveau ---
            CBQuitter.Enabled = True
        
        Case ETATS_BOUTONS.E_AVANT_SUPPRIMER
            '--- avant supprimer ---
        
        Case ETATS_BOUTONS.E_APRES_SUPPRIMER
            '--- après supprimer ---
            CBQuitter.Enabled = True
        
        Case Else
    
    End Select

    '--- affectation ---
    MemDernierBouton = Situation

End Sub

Private Sub CBAgrandirFENETRE_Click()
    On Error Resume Next
    Me.WindowState = vbMaximized
End Sub

Private Sub CBQuitter_Click()
    On Error Resume Next
    DechargeFenetre
End Sub

Private Sub CBQuitter_GotFocus()
    
    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next

    '--- déplacement du focus sur le bouton ---
    With SFocus
        .Left = ActiveControl.Left
        .Top = ActiveControl.Top
        .Height = ActiveControl.Height
        .Width = ActiveControl.Width
        .Visible = True
    End With

End Sub

Private Sub CBQuitter_LostFocus()
    On Error Resume Next
    SFocus.Visible = False
End Sub

Private Sub Form_Resize()
    
    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next
    
    '--- zone mére et fille du déplacement de la fenetre ---
    PBDeplacementFenetre(ZONES_DEPLACEMENT_FENETRE.Z_MERE).Height = Abs(Me.ScaleHeight - PBRenseignementsFenetre.Height - PBBoutons.Height)
    If Me.WindowState = vbMaximized Or Me.WindowState = vbMinimized Then
        
        '--- outils de déplacement invisible ---
        PBOutilsDeplacementFenetre.Visible = False
        
    Else
        
        '--- outils de déplacement visible ---
        With PBOutilsDeplacementFenetre
            .Left = 0
            .Top = 0
            .Height = Me.PBBoutons.ScaleHeight
            .Visible = True
        End With
    
    End If

End Sub

Private Sub HSDeplacementFenetre_Change()
    On Error Resume Next
    PBDeplacementFenetre(ZONES_DEPLACEMENT_FENETRE.Z_FILLE).Left = -HSDeplacementFenetre.Value
End Sub

Private Sub PBBoutons_Resize()
    
    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next
    
    '--- déclaration ---
    
    '--- calculs de l'emplacement des boutons ---
    CBQuitter.Left = PBBoutons.ScaleWidth - MARGES.M_BORD_DROIT - CBQuitter.Width
    
    '--- recalcul du focus après déplacement ---
    With SFocus
        If .Visible = True Then
            .Left = ActiveControl.Left
            .Top = ActiveControl.Top
            .Height = ActiveControl.Height
            .Width = ActiveControl.Width
        End If
    End With

End Sub

'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Rôle      : Décharge la fenêtre
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Private Sub DechargeFenetre()
    
    '--- aiguillage en cas d'erreur ---
    On Error Resume Next
    
    '--- affectation ---
    PremiereActivation = False

    '--- curseur souris par défaut ---
    SourisEnAttente False
    
    '--- neutralisation du timer ---
    With TimerEssais
        .Enabled = False
        .Interval = 0
    End With

    '--- déchargement de la fenêtre ---
    Me.Visible = False
    Unload Me
    Set OccFEssais = Nothing

End Sub

'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Rôle      : Change le curseur de la souris en fonction de l'attente
' Entrées : AttenteOuiNon -> TRUE   = Curseur en forme de sablier
'                                             FALSE = Curseur par défaut
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Sub SourisEnAttente(ByVal AttenteOuiNon As Boolean)
    
    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next

    '--- changement du curseur ---
    If AttenteOuiNon = True Then
        Me.MousePointer = vbHourglass
    Else
        Me.MousePointer = vbDefault
    End If

End Sub

Private Sub PBDeplacementFenetre_Resize(Index As Integer)
    
    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next
        
    If Index = ZONES_DEPLACEMENT_FENETRE.Z_MERE Then

        If Me.WindowState = vbMaximized Then
            
            '--- agrandir la zone fille ---
            With PBDeplacementFenetre(ZONES_DEPLACEMENT_FENETRE.Z_FILLE)
                
                .Left = PBDeplacementFenetre(ZONES_DEPLACEMENT_FENETRE.Z_MERE).ScaleLeft
                .Top = PBDeplacementFenetre(ZONES_DEPLACEMENT_FENETRE.Z_MERE).ScaleTop
                .Height = PBDeplacementFenetre(ZONES_DEPLACEMENT_FENETRE.Z_MERE).ScaleHeight
                .Width = PBDeplacementFenetre(ZONES_DEPLACEMENT_FENETRE.Z_MERE).ScaleWidth
            
            End With
                   
        End If

    Else

    End If
            
    '--- valeur des curseurs ---
    If Me.WindowState <> vbMaximized And Me.WindowState <> vbMinimized Then
        HSDeplacementFenetre.Max = PBDeplacementFenetre(ZONES_DEPLACEMENT_FENETRE.Z_FILLE).Width - _
                                                         PBDeplacementFenetre(ZONES_DEPLACEMENT_FENETRE.Z_MERE).Width
        VSDeplacementFenetre.Max = PBDeplacementFenetre(ZONES_DEPLACEMENT_FENETRE.Z_FILLE).Height - _
                                                         PBDeplacementFenetre(ZONES_DEPLACEMENT_FENETRE.Z_MERE).Height
    End If

End Sub

Private Sub PBRenseignementsFenetre_Resize()
    
    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next
    
    '--- déclaration ---
    
    '--- calculs des emplacements ---
    With PBRenseignementsFenetre
        LRenseignementsFenetre.Left = .ScaleLeft
        LRenseignementsFenetre.Top = .ScaleTop + 30
        LRenseignementsFenetre.Width = .ScaleWidth
        LRenseignementsFenetre.Height = .ScaleHeight
    End With

End Sub

Private Sub TimerEssais_Timer()

    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next
    
    '--- rafraichissement du synoptique ---
    TimerEssais.Enabled = False
    If OccFSynoptique.ArretTachesRapides = False Then
        GestionImageTampon True
        TimerEssais.Enabled = True
    End If

End Sub

Private Sub VSDeplacementFENETRE_Change()
    On Error Resume Next
    PBDeplacementFenetre(ZONES_DEPLACEMENT_FENETRE.Z_FILLE).Top = -VSDeplacementFenetre.Value
End Sub

'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Rôle      : Initialisation de DirectX
' Détails  :
' Retours :
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Private Sub InitialisationDirectX()
        
    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next
    
    '--- déclaration ---
            
            
            
            
            
            With PBImageLigne
                .Left = 0
                .Top = 0
                .Width = LONGUEUR_IMAGE_LIGNE
                .Height = HAUTEUR_IMAGE_LIGNE
            End With
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    '--- création de l'objet direct draw ---
    Set ObjDD = ObjDX.DirectDrawCreate("")
    
    '--- niveau de coopération avec l'écran ---
    Call ObjDD.SetCooperativeLevel(Me.hWnd, DDSCL_NORMAL)
    
    '--- description de la surface de l'écran ---
    With DDSDEcran
        .lFlags = DDSD_CAPS
        .ddsCaps.lCaps = DDSCAPS_PRIMARYSURFACE
    End With
    
    '--- création de la surface ---
    Set ObjDDSEcran = ObjDD.CreateSurface(DDSDEcran)
    
    '--- création de l'objet clipper pour utiliser que certaines régions de l'écran ---
    Set ObjDDClip = ObjDD.CreateClipper(0)
    
    '--- association de l'image à l'objet clipper ---
    ObjDDClip.SetHWnd PBImageLigne.hWnd
    
    '--- attachement du clipping à l'écran ---
    ObjDDSEcran.SetClipper ObjDDClip
    
    
    
    
    
    
    '--- description de l'image tampon (surface invisible dans la mémoire système) ---
    'With DDSDImageTampon
    '    .lFlags = DDSD_CAPS Or DDSD_HEIGHT Or DDSD_WIDTH                                      'Indicate that we want to specify the ddscaps height and width The format of the surface (bits per pixel) will be the same as the primary
    '    .ddsCaps.lCaps = DDSCAPS_OFFSCREENPLAIN Or DDSCAPS_SYSTEMMEMORY 'Indicate that we want a surface that is not visible and that we want it in system memory wich is plentiful as opposed to video memory
    '    .lWidth = PBImageLigne.Width                                                                                   'Specify the height and width of the surface to be the same as the picture box (note unit are in pixels)
    '    .lHeight = PBImageLigne.Height
   ' End With
    
    '--- création de l'image tampon (surface invisible dans la mémoire système) ---
    'Set ObjDDSImageTampon = ObjDD.CreateSurface(DDSDImageTampon)
   
    '--- coordonnées du rectangle de l'image tampon ---
   ' With RImageTampon
   '     .Left = 0
   '     .Top = 0
   '     .Right = DDSDImageTampon.lWidth
   '     .Bottom = DDSDImageTampon.lHeight
   ' End With
    
End Sub

'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Rôle      : Initialisation des surfaces
' Détails  :
' Retours :
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Private Sub InitialisationSurfaces()
    
    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next
    
    '--- déclaration ---
    Dim a As Integer
    Dim CouleurCle As DDCOLORKEY
    Dim DDFormatEnPixels As DDPIXELFORMAT
    
    
    
    
    
    
    '-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    
    '--- description de l'image de la ligne ---
    With DDSDImageLigne
        .lFlags = DDSD_CAPS Or DDSD_WIDTH Or DDSD_HEIGHT
        .ddsCaps.lCaps = DDSCAPS_OFFSCREENPLAIN
        .lWidth = PBImageLigne.Width
        .lHeight = PBImageLigne.Height
    End With
    
    '--- création de la surface et chargement de l'image de la ligne ---
    Set ObjDDSImageLigne = ObjDD.CreateSurfaceFromFile(RepImagesAnodisation & "Synoptique.bmp", DDSDImageLigne)
    
    '--- coordonnées du rectangle du synoptique ---
    With RImageLigne
        .Left = 0
        .Top = 0
        .Right = DDSDImageLigne.lWidth
        .Bottom = DDSDImageLigne.lHeight
    End With
                                                                        
                                                                        
                                                                        
                                                                        
                                                                        
    '--- reconstruction de l'image tampon en mémoire ---
   ' GestionImageTampon False
        
        
End Sub

'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Rôle      : Gére l'image tampon
' Détails  : ModeChoisi -> FALSE = Reconstruit l'image tampon dans la mémoire (il n'y a pas d'affichage)
'                                         TRUE  = Affichage de l'image tampon à l'écran
' Retours :
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Sub GestionImageTampon(ByVal ModeChoisi As Boolean)
    
    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next
    
    '--- déclaration ---
    Dim RDestination As RECT  'coordonnées du rectangle de destination
    

    'Set ObjDDSImageTampon = ObjDDSImageTampon
    'DDSDImageTampon = DDSDImageTampon
    'RImageTampon = RImageTampon



    
    
    If ModeChoisi = False Then
    
        '--- reconstruction ---
        Call ObjDDSImageTampon.BltFast(0, 0, ObjDDSImageLigne, RImageLigne, DDBLTFAST_WAIT)
    
    Else
    
        '--- récupération des coordonnées écran de l'image de la ligne ---
        Call ObjDX.GetWindowRect(PBImageLigne.hWnd, RDestination)
    
        '--- transfert de l'image tampon à l'écran ---
        Call ObjDDSEcran.Blt(RDestination, ObjDDSImageTampon, RImageTampon, DDBLT_WAIT)
    
    End If
    
End Sub


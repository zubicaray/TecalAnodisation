VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{38911DA0-E448-11D0-84A3-00DD01104159}#1.1#0"; "COMCT332.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form FCyclesPonts 
   ClientHeight    =   8535
   ClientLeft      =   2865
   ClientTop       =   5040
   ClientWidth     =   14955
   Icon            =   "FCyclesPonts.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8535
   ScaleWidth      =   14955
   ShowInTaskbar   =   0   'False
   WindowState     =   2  'Maximized
   Begin VB.PictureBox PBBoutons 
      Align           =   2  'Align Bottom
      DrawStyle       =   2  'Dot
      DrawWidth       =   16891
      Height          =   1095
      Left            =   0
      ScaleHeight     =   1035
      ScaleWidth      =   14895
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   7440
      Width           =   14955
      Begin VB.CommandButton CBQuitter 
         BackColor       =   &H00FFFFFF&
         Cancel          =   -1  'True
         Caption         =   "&Quitter"
         DownPicture     =   "FCyclesPonts.frx":014A
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
         Picture         =   "FCyclesPonts.frx":084C
         Style           =   1  'Graphical
         TabIndex        =   11
         ToolTipText     =   " Quitter cette fenêtre "
         Top             =   120
         UseMaskColor    =   -1  'True
         Width           =   1515
      End
      Begin VB.Timer TimerCyclesPonts 
         Enabled         =   0   'False
         Interval        =   500
         Left            =   12120
         Top             =   120
      End
      Begin MSComctlLib.ImageList ILOutilsCyclesPonts 
         Left            =   11400
         Top             =   60
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   185
         ImageHeight     =   19
         MaskColor       =   16711935
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   6
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FCyclesPonts.frx":0F4E
               Key             =   "cycles du pont 1"
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FCyclesPonts.frx":38E4
               Key             =   "cycles du pont 1 en selection"
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FCyclesPonts.frx":627A
               Key             =   "cycles du pont 2"
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FCyclesPonts.frx":8C10
               Key             =   "cycles du pont 2 en selection"
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FCyclesPonts.frx":B5A6
               Key             =   "cycles des ponts 1 et 2"
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FCyclesPonts.frx":DF3C
               Key             =   "cycles des ponts 1 et 2 en selection"
            EndProperty
         EndProperty
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
      Begin VB.Label LLibellesLegende 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Un clic dans une grille permet d'alterner CYCLE ACTUEL / PROCHAIN CYCLE et inversement"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   360
         TabIndex        =   10
         Top             =   420
         Width           =   10815
      End
      Begin VB.Shape SLegende 
         BackColor       =   &H00C0FFC0&
         BackStyle       =   1  'Opaque
         BorderWidth     =   2
         Height          =   555
         Left            =   240
         Shape           =   4  'Rounded Rectangle
         Top             =   270
         Width           =   11085
      End
   End
   Begin VB.PictureBox PBCyclesPonts 
      Align           =   1  'Align Top
      Height          =   6495
      Left            =   0
      ScaleHeight     =   6435
      ScaleWidth      =   14895
      TabIndex        =   3
      Top             =   375
      Width           =   14955
      Begin ComCtl3.CoolBar COBConteneurOutilsCycles 
         Height          =   435
         Left            =   0
         TabIndex        =   4
         Top             =   0
         Width           =   10635
         _ExtentX        =   18759
         _ExtentY        =   767
         BandCount       =   1
         FixedOrder      =   -1  'True
         VariantHeight   =   0   'False
         EmbossHighlight =   8388608
         EmbossShadow    =   16776960
         _CBWidth        =   10635
         _CBHeight       =   435
         _Version        =   "6.7.9782"
         Child1          =   "TOBOutilsCyclesPonts"
         MinHeight1      =   375
         Width1          =   4095
         NewRow1         =   0   'False
         AllowVertical1  =   0   'False
         Begin MSComctlLib.Toolbar TOBOutilsCyclesPonts 
            Height          =   375
            Left            =   30
            TabIndex        =   5
            Top             =   30
            Width           =   10515
            _ExtentX        =   18547
            _ExtentY        =   661
            ButtonWidth     =   5080
            ButtonHeight    =   661
            AllowCustomize  =   0   'False
            Wrappable       =   0   'False
            Style           =   1
            ImageList       =   "ILOutilsCyclesPonts"
            HotImageList    =   "ILOutilsCyclesPonts"
            _Version        =   393216
            BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
               NumButtons      =   6
               BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "cycles des ponts 1 et 2"
                  Object.ToolTipText     =   " Cycles des ponts 1 et 2 en même temps "
                  ImageIndex      =   6
               EndProperty
               BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Style           =   3
               EndProperty
               BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "cycles du pont 1"
                  Object.ToolTipText     =   " Cycles du pont 1 uniquement "
                  ImageIndex      =   1
               EndProperty
               BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Style           =   3
               EndProperty
               BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "cycles du pont 2"
                  Object.ToolTipText     =   " Cycles du pont 2 uniquement "
                  ImageIndex      =   3
               EndProperty
               BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Style           =   3
               EndProperty
            EndProperty
         End
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFGCyclesPonts 
         Height          =   915
         Index           =   1
         Left            =   480
         TabIndex        =   6
         Top             =   1560
         Width           =   4155
         _ExtentX        =   7329
         _ExtentY        =   1614
         _Version        =   393216
         BackColor       =   16777215
         ForeColor       =   0
         Rows            =   31
         Cols            =   6
         BackColorFixed  =   16512
         ForeColorFixed  =   16777215
         BackColorSel    =   16777215
         BackColorBkg    =   12648447
         GridColor       =   0
         GridColorFixed  =   0
         GridColorUnpopulated=   -2147483644
         AllowBigSelection=   0   'False
         FocusRect       =   0
         HighLight       =   0
         ScrollBars      =   2
         Appearance      =   0
         RowSizingMode   =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty FontFixed {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _NumberOfBands  =   1
         _Band(0).Cols   =   6
         _Band(0).GridLinesBand=   0
         _Band(0).TextStyleBand=   0
         _Band(0).TextStyleHeader=   0
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFGCyclesPonts 
         Height          =   975
         Index           =   2
         Left            =   480
         TabIndex        =   7
         Top             =   2760
         Width           =   4155
         _ExtentX        =   7329
         _ExtentY        =   1720
         _Version        =   393216
         BackColor       =   16777215
         ForeColor       =   0
         Rows            =   31
         Cols            =   6
         BackColorFixed  =   16512
         ForeColorFixed  =   16777215
         BackColorBkg    =   12648447
         GridColor       =   12632256
         GridColorUnpopulated=   -2147483644
         AllowBigSelection=   0   'False
         FocusRect       =   0
         HighLight       =   0
         ScrollBars      =   2
         Appearance      =   0
         RowSizingMode   =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty FontFixed {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _NumberOfBands  =   1
         _Band(0).Cols   =   6
         _Band(0).GridLinesBand=   0
         _Band(0).TextStyleBand=   0
         _Band(0).TextStyleHeader=   0
      End
      Begin VB.Label LTitresGrillesCyclesPonts 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H000040C0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   315
         Index           =   1
         Left            =   480
         TabIndex        =   9
         Top             =   1260
         Width           =   4155
      End
      Begin VB.Label LTitresGrillesCyclesPonts 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H000040C0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   315
         Index           =   2
         Left            =   480
         TabIndex        =   8
         Top             =   2460
         Width           =   4155
      End
   End
   Begin VB.PictureBox PBRenseignementsFenetre 
      Align           =   1  'Align Top
      BackColor       =   &H00FF0000&
      Height          =   375
      Left            =   0
      Picture         =   "FCyclesPonts.frx":108D2
      ScaleHeight     =   315
      ScaleWidth      =   14895
      TabIndex        =   1
      Top             =   0
      Width           =   14955
      Begin VB.Label LRenseignementsFenetre 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00008000&
         BackStyle       =   0  'Transparent
         Caption         =   "CYCLES DES PONTS"
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
         TabIndex        =   2
         Top             =   0
         Width           =   11415
         WordWrap        =   -1  'True
      End
   End
End
Attribute VB_Name = "FCyclesPonts"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Rôle                    : Fenêtre gérant les cycles des ponts
' Nom                    : FCyclesPonts.frm
' Date de création : 30/07/2001
' Détails                :
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

'--- déclarations obligatoires ---
Option Explicit

'--- options générales ---
Option Base 1
DefVar A-Z
    
'--- constantes privées ---
Private Const TITRE_FENETRE As String = "CYCLES DES PONTS"
Private Const TITRE_MESSAGES As String = TITRE_FENETRE

'--- énumérations privées ---

'--- formes pour les cycles des ponts ---
Public Enum FORMES_CYCLES_PONTS
    F_CYCLES_PONTS_1_ET_2 = 0
    F_CYCLES_PONT_1 = 1
    F_CYCLES_PONT_2 = 2
End Enum

'--- variables privées ---
Private PremiereActivation As Boolean
Private MemDernierBouton As Long                'mémoire du dernier bouton
Private FormeCyclesPonts As Integer              'forme des cycles des ponts

'--- tableaux privés ---

'--- variables publiques ---
Public NumFenetre As Long                               'numéro de la fenêtre lorsqu'elle devient active

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

Private Sub Form_Activate()
    
    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next
    
    '--- renseigne la fenêtre principale ---
    RenseigneFPrincipale
    
    '--- placement du focus ---
    If PremiereActivation = False Then
        Me.Refresh
        PremiereActivation = True
    End If

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    
    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next
    
    Select Case KeyCode
        
        Case vbKeyF1 To vbKeyF11
            '--- touches de fonctions ---
            OccFSynoptique.SetFocus
            Call OccFSynoptique.GestionTouches(KeyCode, Shift)
        
        Case vbKeyF12
            '--- acquittement des alarmes ---
            AcquittementAlarmes
        
        Case Else
    End Select

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    On Error Resume Next
    CBQuitter_Click
    If UnloadMode = vbFormControlMenu Then
        Cancel = 1
    End If
End Sub

Private Sub Form_Resize()
    
    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next
    
    '--- déclaration ---
    Dim a As Integer
    
    '--- activation du scrolling horizontal des grilles en fonction de la forme de la fenetre ---
    For a = MSHFGCyclesPonts.LBound To MSHFGCyclesPonts.UBound
        With MSHFGCyclesPonts(a)
            If Me.WindowState = vbMaximized Then
                .LeftCol = 1
                .ScrollBars = flexScrollBarVertical
            Else
                .ScrollBars = flexScrollBarBoth
            End If
        End With
    Next a
    
    '--- zone des cycles des ponts ---
    PBCyclesPonts.Height = Abs(Me.ScaleHeight - PBRenseignementsFenetre.Height - PBBoutons.Height)
    
End Sub

Private Sub LRenseignementsFenetre_DblClick()
    On Error Resume Next
    If Me.WindowState = vbMaximized Then
        Me.WindowState = vbNormal
    Else
        Me.WindowState = vbMaximized
    End If
End Sub

Private Sub MSHFGCyclesPonts_Click(Index As Integer)
    
    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next
    
    '--- changement du type affichage ---
    TEtatsPonts(Index).TypesAffichagesCyclesPonts = Not (TEtatsPonts(Index).TypesAffichagesCyclesPonts)

    '--- initialisation et affichage de la grille concernée ---
    GestionGrillesCyclesPonts Index, GG_VIDAGE
    GestionGrillesCyclesPonts Index, GG_INITIALISATION
    GestionGrillesCyclesPonts Index, GG_AFFICHAGE

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

Private Sub PBCyclesPonts_Resize()
    
    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next
    
    '--- déclaration ---
    
    '--- calculs de l'emplacement de la barre d'outils ---
    COBConteneurOutilsCycles.Width = PBCyclesPonts.ScaleWidth
    SelectionneCyclesPonts FormeCyclesPonts

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

'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Rôle      : Effectue le paramètrage de la fenêtre
' Entrées : FormeCyclesPonts -> Formes des cycles des ponts souhaitées
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Sub ParametrageFenetre(ByVal FormeCyclesPonts_ As FORMES_CYCLES_PONTS)
    
    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next

    '--- déclaration ---
    Dim a As Integer
    
    '--- initialisation des grilles ---
    For a = LBound(TEtatsPonts()) To UBound(TEtatsPonts())
        GestionGrillesCyclesPonts a, GG_INITIALISATION
    Next a
    
    '--- modification de la forme des cycles des ponts ---
    If FormeCyclesPonts <> FormeCyclesPonts_ Then
        FormeCyclesPonts = FormeCyclesPonts_
        SelectionneCyclesPonts FormeCyclesPonts
    End If
    
    '--- affichage ---
    For a = LBound(TEtatsPonts()) To UBound(TEtatsPonts())
        GestionGrillesCyclesPonts a, GG_AFFICHAGE
    Next a

    '--- lancement du timer ---
    TimerCyclesPonts.Enabled = True

End Sub

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
    PBBoutons.Picture = ImgFondDesBoutons
    
    '--- gestion de l'états des boutons ---
    GestionBoutons E_CHARGEMENT_FENETRE
    
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
    With TimerCyclesPonts
        .Enabled = False
        .Interval = 0
    End With

    '--- déchargement de la fenêtre ---
    Me.Visible = False
    Unload Me
    Set OccFCyclesPonts = Nothing

End Sub

Private Sub TimerCyclesPonts_Timer()
    
    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next
    
    '--- déclaration ---
    Dim a As Integer

    '--- neutralisation du timer ---
    TimerCyclesPonts.Enabled = False

    '--- rafraichissement des grilles ---
    For a = LBound(TEtatsPonts()) To UBound(TEtatsPonts())
        GestionGrillesCyclesPonts a, GG_AFFICHAGE
    Next a

    '--- réactivation du timer ---
    TimerCyclesPonts.Enabled = True

    '--- bip de passage dans la routine UNIQUEMENT POUR LES TESTS ---
    If PROGRAMME_AVEC_AUTOMATE = False Then Beep

End Sub

'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Rôle      : Gestion des grilles des cycles des ponts
' Entrées :       NumPont -> Fonction de l'énumération PONTS
'                 EtatSouhaite -> Fonction de l'énumération GESTION_GRILLES
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Private Sub GestionGrillesCyclesPonts(ByVal NumPont As PONTS, _
                                                                ByVal EtatSouhaite As GESTION_GRILLES)
    
    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next

    '--- constantes privées ---
    Const LARGEUR_COL_LIBELLE_ACTION_CA_1_PONT As Single = 154.7       'pour le cycle actuel pour un pont visualisé
    Const LARGEUR_COL_LIBELLE_ACTION_PC_1_PONT As Single = 169.7       'pour le prochain cycle pour un pont visualisé
    
    Const LARGEUR_COL_LIBELLE_ACTION_CA_2_PONTS As Single = 52.2       'pour le cycle actuel pour deux ponts visualisés
    Const LARGEUR_COL_LIBELLE_ACTION_PC_2_PONTS As Single = 67.2       'pour le prochain cycle pour deux ponts visualisés
    
    Const EPAISSEUR_CARACTERE As Integer = 140
    
    '--- déclaration ---
    Dim a As Integer, _
            b As Integer, _
            PtrLigne As Integer, _
            Cycle As Integer, _
            NbrColonnesGrillesCyclesPonts As Integer, _
            ColonneNumLignes As Integer, _
            ColonneCodeAction As Integer, _
            ColonneParametre As Integer, _
            ColonneEtatParametre As Integer, _
            ColonneLibelleAction As Integer, _
            NumAction As Integer, _
            PtrActionEnCoursAPI As Integer, _
            NbrParametres As Integer, _
            PtrReelAction As Integer
    Dim LargeurColonneLibelleAction As Single
    Dim Couleur1FondCellule As Long, _
            Couleur2FondCellule As Long, _
            CouleurPlanSelection As Long, _
            CouleurFondSelection As Long
    Dim Parametre As String, _
            EtatParametre As String
        
    '--- affectation du cycle ---
    Cycle = IIf(TEtatsPonts(NumPont).TypesAffichagesCyclesPonts = False, CYCLES.C_ACTUEL, CYCLES.C_PROCHAIN)

    '--- variables des colonnes en fonction du cycle ---
    Select Case Cycle
        
        Case CYCLES.C_ACTUEL
            '--- couleurs des cellules ---
            Couleur1FondCellule = COULEURS.ORANGE_0
            Couleur2FondCellule = COULEURS.ORANGE_1
                            
            '--- affectation des couleurs en fonction des défauts du pont ---
            If TEtatsPonts(NumPont).Alarmes <> "" Then
                CouleurFondSelection = COULEURS.ROUGE_3
                CouleurPlanSelection = COULEURS.JAUNE_3
            Else
                CouleurFondSelection = COULEURS.VERT_3
                CouleurPlanSelection = COULEURS.NOIR
            End If

            '--- colonnes des grilles ---
            NbrColonnesGrillesCyclesPonts = 4
            ColonneNumLignes = 0
            ColonneCodeAction = 1
            ColonneParametre = 2
            ColonneEtatParametre = 3
            ColonneLibelleAction = 4
        
        Case CYCLES.C_PROCHAIN
            '--- couleurs des cellules ---
            Couleur1FondCellule = COULEURS.BLEU_0
            Couleur2FondCellule = COULEURS.BLEU_1
            
            '--- colonnes des grilles ---
            NbrColonnesGrillesCyclesPonts = 3
            ColonneNumLignes = 0
            ColonneCodeAction = 1
            ColonneParametre = 2
            ColonneLibelleAction = 3
        
        Case Else
    End Select
    
    Select Case EtatSouhaite

        Case GESTION_GRILLES.GG_INITIALISATION
            '--- initialisation de la grille ---
            With MSHFGCyclesPonts(NumPont)
                        
                .Redraw = False
                    
                .Clear
                        
                .FixedCols = 1
                .FixedRows = 1
                .Rows = NBR_LIGNES_CYCLES_PONTS + .FixedRows
                .Cols = NbrColonnesGrillesCyclesPonts + .FixedCols
                .RowHeightMin = 310                                                                 'épaisseur mini des lignes
                .Row = 0

                '--- paramétrages de chaque colonne ---
                .Col = ColonneNumLignes
                .ColWidth(.Col) = 3 * EPAISSEUR_CARACTERE: .Text = ""
                .ColAlignment(.Col) = flexAlignRightCenter
                
                .Col = ColonneCodeAction
                .ColWidth(.Col) = 15 * EPAISSEUR_CARACTERE: .Text = "Code"
                .ColAlignment(.Col) = flexAlignCenterCenter
                
                .Col = ColonneParametre
                .ColWidth(.Col) = 15 * EPAISSEUR_CARACTERE: .Text = "Paramètre"
                .ColAlignment(.Col) = flexAlignCenterCenter
                
                If Cycle = CYCLES.C_ACTUEL Then
                    .Col = ColonneEtatParametre
                    .ColWidth(.Col) = 15 * EPAISSEUR_CARACTERE: .Text = "Etat actuel"
                    .ColAlignment(.Col) = flexAlignCenterCenter
                End If
                
                .Col = ColonneLibelleAction
                If Cycle = CYCLES.C_ACTUEL Then
                    .ColWidth(.Col) = LARGEUR_COL_LIBELLE_ACTION_CA_2_PONTS * EPAISSEUR_CARACTERE
                Else
                    .ColWidth(.Col) = LARGEUR_COL_LIBELLE_ACTION_PC_2_PONTS * EPAISSEUR_CARACTERE
                End If
                .Text = "Action"
                .ColAlignment(.Col) = flexAlignLeftCenter
                        
                '--- centrage des titres ---
                .Row = 0
                For a = 1 To Pred(.Cols)
                    .Col = a
                    .CellAlignment = flexAlignCenterCenter
                Next a

                '--- N° de lignes, vidage des champs ---
                .Col = 0
                For a = 1 To NBR_LIGNES_CYCLES_PONTS
                    .Row = a: .Text = CStr(a)
                Next a
        
                '--- couleur du titre et du fond des colonnes fixes ---
                Select Case Cycle
                    
                    Case CYCLES.C_ACTUEL
                        '--- cycle actuel ---
                        With LTitresGrillesCyclesPonts(NumPont)
                            .BackColor = COULEURS.ORANGE_4
                            .ForeColor = COULEURS.BLANC
                            .Caption = "CYCLE ACTUEL - PONT " & IIf(NumPont = PONTS.P_1, "1", "2")
                        End With
                        .BackColorFixed = COULEURS.ORANGE_5
                    
                    Case CYCLES.C_PROCHAIN
                        '--- prochain cycle ---
                        With LTitresGrillesCyclesPonts(NumPont)
                            .BackColor = COULEURS.BLEU_4
                            .ForeColor = COULEURS.BLANC
                            .Caption = "PROCHAIN CYCLE - PONT " & IIf(NumPont = PONTS.P_1, "1", "2")
                        End With
                        .BackColorFixed = COULEURS.BLEU_5
                    
                    Case Else
                End Select
                        
                '--- couleur indépendante pour chaque colonne ---
                .BackColor = Couleur2FondCellule
                .FillStyle = flexFillRepeat
                For a = .FixedCols To .Cols() - 1 Step 2
                    .Col = a
                    .Row = .FixedRows
                    .RowSel = .Rows - 1
                    .CellBackColor = Couleur1FondCellule
                Next a
                .FillStyle = flexFillSingle

                '--- fixer le curseur ---
                .Row = 1
                .Col = ColonneCodeAction

                .Redraw = True

            End With

        Case GESTION_GRILLES.GG_VIDAGE
            '--- vidage de la grille ---
            With MSHFGCyclesPonts(NumPont)
                .Redraw = False
                For a = 1 To NBR_LIGNES_CYCLES_PONTS
                    .Row = a
                    .FillStyle = flexFillRepeat
                    .Col = .FixedCols
                    .ColSel = .Cols - 1
                    .Text = ""
                    .FillStyle = flexFillSingle
                Next a
                .TopRow = 1
                .LeftCol = ColonneCodeAction
                .Redraw = True
            End With

        Case GESTION_GRILLES.GG_AFFICHAGE
            '--- construction de la grille ---
            With MSHFGCyclesPonts(NumPont)

                .Redraw = False
    
                '--- modification de la largeur de la colonne du libellé de l'action ---
                If FormeCyclesPonts = FORMES_CYCLES_PONTS.F_CYCLES_PONTS_1_ET_2 Then
                    LargeurColonneLibelleAction = IIf(Cycle = CYCLES.C_ACTUEL, LARGEUR_COL_LIBELLE_ACTION_CA_2_PONTS, LARGEUR_COL_LIBELLE_ACTION_PC_2_PONTS) * EPAISSEUR_CARACTERE
                Else
                    LargeurColonneLibelleAction = IIf(Cycle = CYCLES.C_ACTUEL, LARGEUR_COL_LIBELLE_ACTION_CA_1_PONT, LARGEUR_COL_LIBELLE_ACTION_PC_1_PONT) * EPAISSEUR_CARACTERE
                End If
                .ColWidth(ColonneLibelleAction) = LargeurColonneLibelleAction
                
                '--- index de l'action en cours pour le changement de couleur ---
                PtrActionEnCoursAPI = TEtatsPonts(NumPont).PtrEtActionEnCoursAPI.PtrAction
                
                '--- affectation ---
                NbrParametres = 0
                
                For a = 1 To NBR_LIGNES_CYCLES_PONTS
                    
                    '--- affectation ---
                    With TEtatsPonts(NumPont).TCyclesPonts(Cycle, a)
                        NumAction = .NumAction
                        Parametre = .Parametre
                        EtatParametre = .EtatParametre
                    End With
                    
                    .Row = a
                    
                    '--- recherche du pointeur réel (à cause du décalage du aux paramètres) ---
                    If PtrActionEnCoursAPI > 0 Then
                        If a + NbrParametres = PtrActionEnCoursAPI Then
                            PtrReelAction = a
                        End If
                        If Parametre <> "" Then Inc NbrParametres
                    End If
                    
                    If NumAction = 0 Then
                        
                        '--- effacement de la ligne ---
                        If .Text <> "" Then
                            .Col = ColonneCodeAction:  .CellBackColor = Couleur1FondCellule: .CellForeColor = COULEURS.NOIR: .Text = ""
                            .Col = ColonneParametre: .CellBackColor = Couleur2FondCellule: .CellForeColor = COULEURS.NOIR: .Text = ""
                            If Cycle = CYCLES.C_ACTUEL Then
                                .Col = ColonneEtatParametre: .CellBackColor = Couleur1FondCellule: .CellForeColor = COULEURS.NOIR: .Text = ""
                                .Col = ColonneLibelleAction: .CellBackColor = Couleur2FondCellule: .CellForeColor = COULEURS.NOIR: .Text = ""
                            Else
                                .Col = ColonneLibelleAction: .CellBackColor = Couleur1FondCellule: .CellForeColor = COULEURS.NOIR: .Text = ""
                            End If
                        End If
                    
                    Else
                    
                        '--- code de l'action ---
                        .Col = ColonneCodeAction
                        If a = PtrReelAction And Cycle = CYCLES.C_ACTUEL Then
                            If .CellBackColor <> CouleurFondSelection Then .CellBackColor = CouleurFondSelection: .CellForeColor = CouleurPlanSelection
                        Else
                            If .CellBackColor <> Couleur1FondCellule Then .CellBackColor = Couleur1FondCellule: .CellForeColor = COULEURS.NOIR
                        End If
                        If .Text <> TActions(NumAction).CodeAction Then .Text = TActions(NumAction).CodeAction

                        '--- paramètre ---
                        .Col = ColonneParametre
                        If a = PtrReelAction And Cycle = CYCLES.C_ACTUEL Then
                            If .CellBackColor <> CouleurFondSelection Then .CellBackColor = CouleurFondSelection: .CellForeColor = CouleurPlanSelection
                        Else
                            If .CellBackColor <> Couleur2FondCellule Then .CellBackColor = Couleur2FondCellule: .CellForeColor = COULEURS.NOIR
                        End If
                        If .Text <> Parametre Then .Text = Parametre
                        
                        If Cycle = CYCLES.C_ACTUEL Then
                            
                            '--- état du paramètre ---
                            .Col = ColonneEtatParametre
                            If a = PtrReelAction Then
                                If .CellBackColor <> CouleurFondSelection Then .CellBackColor = CouleurFondSelection: .CellForeColor = CouleurPlanSelection
                            Else
                                If .CellBackColor <> Couleur1FondCellule Then .CellBackColor = Couleur1FondCellule: .CellForeColor = COULEURS.NOIR
                            End If
                            If .Text <> EtatParametre Then .Text = EtatParametre
                        
                            '--- lillellé de l'action ---
                            .Col = ColonneLibelleAction
                            If a = PtrReelAction Then
                                If .CellBackColor <> CouleurFondSelection Then .CellBackColor = CouleurFondSelection: .CellForeColor = CouleurPlanSelection
                            Else
                                If .CellBackColor <> Couleur2FondCellule Then .CellBackColor = Couleur2FondCellule: .CellForeColor = COULEURS.NOIR
                            End If
                            If .Text <> TActions(NumAction).LibelleAction Then .Text = TActions(NumAction).LibelleAction
                                                                            
                        Else
                        
                            '--- lillellé de l'action ---
                            .Col = ColonneLibelleAction
                            If .CellBackColor <> Couleur1FondCellule Then .CellBackColor = Couleur1FondCellule: .CellForeColor = COULEURS.NOIR
                            If .Text <> TActions(NumAction).LibelleAction Then .Text = TActions(NumAction).LibelleAction
                        
                        End If
                    
                    End If
                Next a
                
                '--- rendre toujours visible l'indication de l'index en cours ---
                If PtrReelAction > 0 And PtrReelAction + 1 <= NBR_LIGNES_CYCLES_PONTS Then
                    If .RowIsVisible(PtrReelAction + 1) = False Then
                        .TopRow = PtrReelAction
                        .LeftCol = ColonneCodeAction
                    End If
                End If
                            
                '--- restitution des valeurs ligne, colonne ---
                .Redraw = True

            End With

        Case Else

    End Select

End Sub

'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Rôle      : Sélectionne l'affichage des cycles des ponts
' Entrées : FormeSouhaitee -> Forme souhaitée fonction de l'énumération FORMES_CYCLES_PONTS
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Private Sub SelectionneCyclesPonts(ByVal FormeSouhaitee As FORMES_CYCLES_PONTS)

    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next

    '--- déclaration ---
    Dim a As Integer
    Dim TwipsParPixelX As Single, _
            TwipsParPixelY As Single
    
    '--- désélection de toutes les formes des outils ---
    TOBOutilsCyclesPonts.buttons("cycles des ponts 1 et 2").Image = ILOutilsCyclesPonts.ListImages("cycles des ponts 1 et 2").Index
    TOBOutilsCyclesPonts.buttons("cycles du pont 1").Image = ILOutilsCyclesPonts.ListImages("cycles du pont 1").Index
    TOBOutilsCyclesPonts.buttons("cycles du pont 2").Image = ILOutilsCyclesPonts.ListImages("cycles du pont 2").Index
    
    '--- chargement de la forme des outils ---
    Select Case FormeSouhaitee
        Case FORMES_CYCLES_PONTS.F_CYCLES_PONTS_1_ET_2
            TOBOutilsCyclesPonts.buttons("cycles des ponts 1 et 2").Image = ILOutilsCyclesPonts.ListImages("cycles des ponts 1 et 2 en selection").Index
        Case FORMES_CYCLES_PONTS.F_CYCLES_PONT_1
            TOBOutilsCyclesPonts.buttons("cycles du pont 1").Image = ILOutilsCyclesPonts.ListImages("cycles du pont 1 en selection").Index
        Case FORMES_CYCLES_PONTS.F_CYCLES_PONT_2
            TOBOutilsCyclesPonts.buttons("cycles du pont 2").Image = ILOutilsCyclesPonts.ListImages("cycles du pont 2 en selection").Index
        Case Else
    End Select
    
    '--- rafraichissement des outils ---
    TOBOutilsCyclesPonts.Refresh
    
    '--- effacement des titres et grilles par défaut ---
    For a = MSHFGCyclesPonts.LBound To MSHFGCyclesPonts.UBound
        LTitresGrillesCyclesPonts(a).Visible = False
        MSHFGCyclesPonts(a).Visible = False
    Next a
    
    '--- affectation ---
    TwipsParPixelX = Screen.TwipsPerPixelX
    TwipsParPixelY = Screen.TwipsPerPixelY
            
    '--- forme des cycles des ponts ---
    Select Case FormeSouhaitee
        
        Case FORMES_CYCLES_PONTS.F_CYCLES_PONTS_1_ET_2
            '--- cycles du pont 1 et 2 ---
            With LTitresGrillesCyclesPonts(PONTS.P_1)
                .Left = 0
                .Top = COBConteneurOutilsCycles.Height + 2 * TwipsParPixelY
                .Width = COBConteneurOutilsCycles.Width / 2 - TwipsParPixelX
                .Visible = True
            End With
            With MSHFGCyclesPonts(PONTS.P_1)
                .Left = 0
                .Top = COBConteneurOutilsCycles.Height + LTitresGrillesCyclesPonts(PONTS.P_1).Height
                .Width = COBConteneurOutilsCycles.Width / 2 - TwipsParPixelX
                .Height = Abs(Int(PBCyclesPonts.ScaleHeight - COBConteneurOutilsCycles.Height - LTitresGrillesCyclesPonts(PONTS.P_1).Height + TwipsParPixelY))
                .Visible = True
                .Redraw = True
            End With
            With LTitresGrillesCyclesPonts(PONTS.P_2)
                .Left = COBConteneurOutilsCycles.Width / 2
                .Top = COBConteneurOutilsCycles.Height + 2 * TwipsParPixelY
                .Width = COBConteneurOutilsCycles.Width / 2 - TwipsParPixelX
                .Visible = True
            End With
            With MSHFGCyclesPonts(PONTS.P_2)
                .Left = LTitresGrillesCyclesPonts(PONTS.P_2).Left
                .Top = MSHFGCyclesPonts(PONTS.P_1).Top
                .Width = MSHFGCyclesPonts(PONTS.P_1).Width
                .Height = MSHFGCyclesPonts(PONTS.P_1).Height
                .Visible = True
                .Redraw = True
            End With
        
        Case FORMES_CYCLES_PONTS.F_CYCLES_PONT_1
            '--- cycles du pont 1 ---
            With LTitresGrillesCyclesPonts(PONTS.P_1)
                .Left = 0
                .Top = COBConteneurOutilsCycles.Height + 2 * TwipsParPixelY
                .Width = COBConteneurOutilsCycles.Width
                .Visible = True
            End With
            With MSHFGCyclesPonts(PONTS.P_1)
                .Left = 0
                .Top = COBConteneurOutilsCycles.Height + LTitresGrillesCyclesPonts(PONTS.P_1).Height
                .Width = COBConteneurOutilsCycles.Width
                .Height = PBCyclesPonts.ScaleHeight - COBConteneurOutilsCycles.Height - LTitresGrillesCyclesPonts(PONTS.P_1).Height
                .Visible = True
                .Redraw = True
            End With
    
        Case FORMES_CYCLES_PONTS.F_CYCLES_PONT_2
            '--- cycles du pont 2 ---
            With LTitresGrillesCyclesPonts(PONTS.P_2)
                .Left = 0
                .Top = COBConteneurOutilsCycles.Height + 2 * TwipsParPixelY
                .Width = COBConteneurOutilsCycles.Width
                .Visible = True
            End With
            With MSHFGCyclesPonts(PONTS.P_2)
                .Left = 0
                .Top = COBConteneurOutilsCycles.Height + LTitresGrillesCyclesPonts(PONTS.P_2).Height
                .Width = COBConteneurOutilsCycles.Width
                .Height = PBCyclesPonts.ScaleHeight - COBConteneurOutilsCycles.Height - LTitresGrillesCyclesPonts(PONTS.P_2).Height
                .Visible = True
                .Redraw = True
            End With
        
        Case Else
    
    End Select
    
    '--- affectation ---
    FormeCyclesPonts = FormeSouhaitee
    
    '--- rafraichissement des grilles ---
    For a = LBound(TEtatsPonts()) To UBound(TEtatsPonts())
        GestionGrillesCyclesPonts a, GG_AFFICHAGE
    Next a

End Sub

Private Sub LTitresGrillesCyclesPonts_Click(Index As Integer)
    On Error Resume Next
    MSHFGCyclesPonts_Click (Index)
End Sub

Private Sub TOBOutilsCyclesPonts_ButtonClick(ByVal Button As MSComctlLib.Button)
    
    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next

    '--- sélection en fonction de l'outil cliqué ---
    Select Case Button.Key
        
        Case "cycles des ponts 1 et 2"
            '--- cycles des ponts 1 et 2 ---
            SelectionneCyclesPonts F_CYCLES_PONTS_1_ET_2
        
        Case "cycles du pont 1"
            '--- cycle du pont 1 ---
            SelectionneCyclesPonts F_CYCLES_PONT_1
        
        Case "cycles du pont 2"
            '--- cycle du pont 2 ---
            SelectionneCyclesPonts F_CYCLES_PONT_2
        
        Case Else

    End Select

End Sub

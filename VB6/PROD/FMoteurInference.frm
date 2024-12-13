VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.ocx"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "mshflxgd.ocx"
Begin VB.Form FMoteurInference 
   ClientHeight    =   11010
   ClientLeft      =   1455
   ClientTop       =   135
   ClientWidth     =   13395
   Icon            =   "FMoteurInference.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   NegotiateMenus  =   0   'False
   ScaleHeight     =   11010
   ScaleWidth      =   13395
   ShowInTaskbar   =   0   'False
   WindowState     =   2  'Maximized
   Begin VB.PictureBox PBRenseignementsFenetre 
      Align           =   1  'Align Top
      BackColor       =   &H00FF0000&
      ForeColor       =   &H0000FFFF&
      Height          =   375
      Left            =   0
      Picture         =   "FMoteurInference.frx":014A
      ScaleHeight     =   315
      ScaleWidth      =   13335
      TabIndex        =   3
      Top             =   0
      Width           =   13395
      Begin VB.Label LRenseignementsFenetre 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "MOTEUR D'INFERENCE"
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
         Height          =   225
         Left            =   240
         TabIndex        =   4
         Top             =   0
         Width           =   11415
         WordWrap        =   -1  'True
      End
   End
   Begin VB.PictureBox PBBoutons 
      Align           =   2  'Align Bottom
      DrawStyle       =   2  'Dot
      DrawWidth       =   16891
      Height          =   1095
      Left            =   0
      ScaleHeight     =   1035
      ScaleWidth      =   13335
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   9915
      Width           =   13395
      Begin VB.CommandButton CBReduire 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Réduire la fenêtre"
         DownPicture     =   "FMoteurInference.frx":24A8C
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   840
         Left            =   7920
         MaskColor       =   &H00FF00FF&
         Picture         =   "FMoteurInference.frx":2518E
         Style           =   1  'Graphical
         TabIndex        =   11
         ToolTipText     =   " Réduire cette fenêtre à la taille minimum "
         Top             =   120
         UseMaskColor    =   -1  'True
         Width           =   2115
      End
      Begin VB.CommandButton CBQuitter 
         BackColor       =   &H00FFFFFF&
         Cancel          =   -1  'True
         Caption         =   "&Quitter"
         DownPicture     =   "FMoteurInference.frx":25890
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
         Left            =   10200
         MaskColor       =   &H00FF00FF&
         Picture         =   "FMoteurInference.frx":25F92
         Style           =   1  'Graphical
         TabIndex        =   10
         ToolTipText     =   " Quitter cette fenêtre "
         Top             =   120
         UseMaskColor    =   -1  'True
         Width           =   1515
      End
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
            DownPicture     =   "FMoteurInference.frx":26694
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
            Picture         =   "FMoteurInference.frx":2683E
            Style           =   1  'Graphical
            TabIndex        =   7
            ToolTipText     =   " Agrandissement de la fenêtre "
            Top             =   0
            UseMaskColor    =   -1  'True
            Width           =   900
         End
      End
      Begin VB.Timer TimerMoteurInference 
         Enabled         =   0   'False
         Interval        =   500
         Left            =   4980
         Top             =   180
      End
      Begin MSComctlLib.ImageList ILOutilsMoteurInference 
         Left            =   4320
         Top             =   120
         _ExtentX        =   794
         _ExtentY        =   794
         BackColor       =   -2147483643
         ImageWidth      =   16
         ImageHeight     =   16
         MaskColor       =   16711935
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   4
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FMoteurInference.frx":269E8
               Key             =   "chargement"
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FMoteurInference.frx":26B4C
               Key             =   "charges en ligne"
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FMoteurInference.frx":26CA8
               Key             =   "ponts"
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FMoteurInference.frx":26E04
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin VB.Shape SFocus 
         BorderColor     =   &H000000FF&
         BorderWidth     =   5
         Height          =   405
         Left            =   5580
         Top             =   180
         Visible         =   0   'False
         Width           =   600
      End
   End
   Begin VB.PictureBox PBDeplacementFenetre 
      Align           =   1  'Align Top
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   14055
      Index           =   0
      Left            =   0
      ScaleHeight     =   14055
      ScaleWidth      =   13395
      TabIndex        =   1
      Top             =   375
      Width           =   13395
      Begin VB.PictureBox PBDeplacementFenetre 
         BorderStyle     =   0  'None
         Height          =   13935
         Index           =   1
         Left            =   0
         ScaleHeight     =   13935
         ScaleWidth      =   28695
         TabIndex        =   2
         Top             =   0
         Width           =   28695
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFGAnalysesMoteurInference 
            Height          =   2880
            Index           =   0
            Left            =   240
            TabIndex        =   5
            Top             =   480
            Width           =   13995
            _ExtentX        =   24686
            _ExtentY        =   5080
            _Version        =   393216
            BackColor       =   16777215
            ForeColor       =   0
            Rows            =   31
            Cols            =   3
            FixedRows       =   0
            BackColorFixed  =   16576
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
            _Band(0).Cols   =   3
            _Band(0).GridLinesBand=   0
            _Band(0).TextStyleBand=   0
            _Band(0).TextStyleHeader=   0
         End
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFGAnalysesMoteurInference 
            Height          =   12075
            Index           =   3
            Left            =   17760
            TabIndex        =   12
            Top             =   120
            Width           =   13995
            _ExtentX        =   24686
            _ExtentY        =   21299
            _Version        =   393216
            BackColor       =   16777215
            ForeColor       =   0
            Rows            =   31
            Cols            =   3
            FixedRows       =   0
            BackColorFixed  =   12582912
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
            _Band(0).Cols   =   3
            _Band(0).GridLinesBand=   0
            _Band(0).TextStyleBand=   0
            _Band(0).TextStyleHeader=   0
         End
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFGAnalysesMoteurInference 
            Height          =   5775
            Index           =   2
            Left            =   240
            TabIndex        =   13
            Top             =   6780
            Width           =   13995
            _ExtentX        =   24686
            _ExtentY        =   10186
            _Version        =   393216
            BackColor       =   16777215
            ForeColor       =   0
            Rows            =   31
            Cols            =   3
            FixedRows       =   0
            BackColorFixed  =   32768
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
            _Band(0).Cols   =   3
            _Band(0).GridLinesBand=   0
            _Band(0).TextStyleBand=   0
            _Band(0).TextStyleHeader=   0
         End
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFGAnalysesMoteurInference 
            Height          =   2895
            Index           =   1
            Left            =   240
            TabIndex        =   14
            Top             =   3360
            Width           =   13995
            _ExtentX        =   24686
            _ExtentY        =   5106
            _Version        =   393216
            BackColor       =   16777215
            ForeColor       =   0
            Rows            =   31
            Cols            =   3
            FixedRows       =   0
            BackColorFixed  =   16576
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
            _Band(0).Cols   =   3
            _Band(0).GridLinesBand=   0
            _Band(0).TextStyleBand=   0
            _Band(0).TextStyleHeader=   0
         End
         Begin VB.Label LLibelles 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00008000&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "PONTS"
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
            Height          =   255
            Index           =   2
            Left            =   240
            TabIndex        =   17
            Top             =   6540
            Width           =   13995
         End
         Begin VB.Label LLibelles 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00C00000&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "CHARGES EN LIGNE"
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
            Height          =   255
            Index           =   1
            Left            =   14460
            TabIndex        =   16
            Top             =   240
            Width           =   13995
         End
         Begin VB.Label LLibelles 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H000040C0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "CHARGEMENT / DECHARGEMENT"
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
            Height          =   255
            Index           =   0
            Left            =   240
            TabIndex        =   15
            Top             =   240
            Width           =   13995
         End
      End
   End
End
Attribute VB_Name = "FMoteurInference"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Rôle                    : Fenêtre gérant l'affichage des analyses du moteur d'inférence
' Nom                    : FMoteurInference.frm
' Date de création : 11/01/2001
' Détails                :
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

'--- déclarations obligatoires ---
Option Explicit

'--- options générales ---
Option Base 1
DefVar A-Z
    
'--- constantes privées ---
Private Const NBR_COLONNES_ANALYSES_MOTEUR_INFERENCE As Integer = 3

Private Const TITRE_FENETRE As String = "MOTEUR D'INFERENCE"
Private Const TITRE_MESSAGES As String = TITRE_FENETRE

'--- énumérations privées ---
Private Enum NBR_LIGNES_MOTEUR_INFERENCE
    NL_CHARGEMENT = 30
    NL_DECHARGEMENT = 30
    NL_PONTS = 30
    NL_CHARGES_EN_LIGNE = 50
End Enum

Private Enum ANALYSES_MOTEUR_INFERENCE
    A_CHARGEMENT = 0
    A_DECHARGEMENT = 1
    A_PONTS = 2
    A_CHARGES_EN_LIGNE = 3
End Enum

Private Enum COLONNES_ANALYSES_MOTEUR_INFERENCE
    C_NUM_LIGNES = 0         'n° de lignes
    C_DONNEES_1 = 1           'données 1
    C_DONNEES_2 = 2           'données 2
    C_DONNEES_3 = 3           'données 3
End Enum

'--- types privés ---

'--- variables privées ---
Private PremiereActivation As Boolean
Private InterdireEvenements As Boolean             'pour interdire certains évènements
Private LigneDepartDeplacement As Integer        'ligne de départ en cas de déplacement d'un détail
Private LigneArriveeDeplacement As Integer       'ligne de d'arrivée en cas de déplacement d'un détail
Private MemDernierBouton As Long                     'mémoire du dernier bouton

'--- tableaux privés ---

'--- variables publiques ---
Public NumFenetre As Long                                  'numéro de la fenêtre lorsqu'elle devient active

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

Private Sub CBReduire_Click()
    On Error Resume Next
    Me.WindowState = vbMinimized
End Sub

Private Sub CBReduire_GotFocus()
    
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

Private Sub CBReduire_LostFocus()
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
    PBDeplacementFenetre(ZONES_DEPLACEMENT_FENETRE.Z_FILLE).Left = -HSDeplacementFenetre.value
End Sub

Private Sub LRenseignementsFenetre_DblClick()
    On Error Resume Next
    If Me.WindowState = vbMaximized Then
        Me.WindowState = vbNormal
    Else
        Me.WindowState = vbMaximized
    End If
End Sub

Private Sub PBBoutons_Resize()

    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next
    
    '--- calculs de l'emplacement des boutons ---
    CBQuitter.Left = PBBoutons.ScaleWidth - MARGES.M_BORD_DROIT - CBQuitter.Width
    CBReduire.Left = CBQuitter.Left - MARGES.M_ENTRE_BOUTONS - CBReduire.Width
    
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
            If MemDernierBouton = ETATS_BOUTONS.E_APRES_NOUVEAU Then Exit Sub
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

    End If
            
    '--- valeur des curseurs ---
    If Me.WindowState <> vbMaximized And Me.WindowState <> vbMinimized Then
        HSDeplacementFenetre.Max = PBDeplacementFenetre(ZONES_DEPLACEMENT_FENETRE.Z_FILLE).Width - _
                                                         PBDeplacementFenetre(ZONES_DEPLACEMENT_FENETRE.Z_MERE).Width
        VSDeplacementFenetre.Max = PBDeplacementFenetre(ZONES_DEPLACEMENT_FENETRE.Z_FILLE).Height - _
                                                         PBDeplacementFenetre(ZONES_DEPLACEMENT_FENETRE.Z_MERE).Height
    End If

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
' Entrées :
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Sub ParametrageFenetre()
    
    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next

    '--- affectation ---
    
    '--- affichage ---
    GestionMoteurInference A_CHARGEMENT, GG_INITIALISATION
    GestionMoteurInference A_CHARGEMENT, GG_AFFICHAGE
    
    GestionMoteurInference A_DECHARGEMENT, GG_INITIALISATION
    GestionMoteurInference A_DECHARGEMENT, GG_AFFICHAGE
    
    GestionMoteurInference A_PONTS, GG_INITIALISATION
    GestionMoteurInference A_PONTS, GG_AFFICHAGE
    
    GestionMoteurInference A_CHARGES_EN_LIGNE, GG_INITIALISATION
    GestionMoteurInference A_CHARGES_EN_LIGNE, GG_AFFICHAGE
    
    '--- lancement du timer ---
    TimerMoteurInference.Enabled = True
    
End Sub

'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Rôle      : Décharge la fenêtre
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Private Sub DechargeFenetre()
    
    '--- aiguillage en cas d'erreur ---
    On Error Resume Next

    '--- curseur souris par défaut ---
    SourisEnAttente False

    '--- neutralisation du timer ---
    With TimerMoteurInference
        .Enabled = False
        .Interval = 0
    End With

    '--- déchargement de la fenêtre ---
    Me.Visible = False
    Unload Me
    Set OccFMoteurInference = Nothing
    
End Sub

'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Rôle      : Initialise la fenêtre (chargement ou en vue de la rendre visible)
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Sub InitialisationFenetre()
    
    '--- aiguillage en cas d'erreurs ---
    On Error GoTo GestionErreurs

    '--- déclaration ---

    '--- affectation ---

    '--- divers sur la fenêtre ---
    With Me
        .Caption = TITRE_FENETRE
        .WindowState = vbMaximized
    End With
    
    '--- images des fonds ---
    PBDeplacementFenetre(ZONES_DEPLACEMENT_FENETRE.Z_FILLE).Picture = ImgFondOrange2
    PBBoutons.Picture = ImgFondDesBoutons
    
    '--- gestion de l'états des boutons ---
    GestionBoutons E_CHARGEMENT_FENETRE

    Exit Sub

'--- gestion des erreurs ---
GestionErreurs:

    '--- affichage du message d'erreur ---
    MessageErreur TITRE_MESSAGES, Err.Description, Err.Number
    
End Sub

Private Sub PBRenseignementsFenetre_Resize()
    
    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next
    
    '--- calculs des emplacements ---
    With PBRenseignementsFenetre
        LRenseignementsFenetre.Left = .ScaleLeft
        LRenseignementsFenetre.Top = .ScaleTop + 30
        LRenseignementsFenetre.Width = .ScaleWidth
        LRenseignementsFenetre.Height = .ScaleHeight
    End With

End Sub

Private Sub TimerMoteurInference_Timer()
    
    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next
    
    '--- déclaration ---

    '--- neutralisation du timer ---
    TimerMoteurInference.Enabled = False

    '--- rafraichissement des grilles ---
    GestionMoteurInference A_CHARGEMENT, GG_AFFICHAGE
    GestionMoteurInference A_DECHARGEMENT, GG_AFFICHAGE
    GestionMoteurInference A_PONTS, GG_AFFICHAGE
    GestionMoteurInference A_CHARGES_EN_LIGNE, GG_AFFICHAGE

    '--- réactivation du timer ---
    TimerMoteurInference.Enabled = True

    '--- bip de passage dans la routine UNIQUEMENT POUR LES TESTS ---
    If PROGRAMME_AVEC_AUTOMATE = False Then Beep

End Sub

Private Sub VSDeplacementFENETRE_Change()
    On Error Resume Next
    PBDeplacementFenetre(ZONES_DEPLACEMENT_FENETRE.Z_FILLE).Top = -VSDeplacementFenetre.value
End Sub

'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Rôle      : Affiche une cellule dans la grille du moteur d'inférence
' Entrées : AnalyseChoisie -> Analyse choisie fonction de l'énumération ANALYSES_MOTEUR_INFERENCE
'                     CouleurFond -> Couleur de fond du texte
'                      CouleurPlan -> Couleur du texte
'                TypeAlignement -> Type d'alignement en fonction des constantes de la grille
'                                 Texte -> Texte à afficher
' Retours :
' Détails  :
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Private Sub AfficheTexte(ByVal AnalyseChoisie As ANALYSES_MOTEUR_INFERENCE, _
                                        ByVal CouleurFond As Long, _
                                        ByVal CouleurPlan As Long, _
                                        ByVal TypeAlignement As Integer, _
                                        ByVal Texte As String)
    
    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next
    
    With MSHFGAnalysesMoteurInference(AnalyseChoisie)
    
        '--- ligne / colonne ---
        .Col = COLONNES_ANALYSES_MOTEUR_INFERENCE.C_DONNEES_2
    
        '--- couleur de fond, de plan, alignement ---
        If .CellBackColor <> CouleurFond Then .CellBackColor = CouleurFond
        If .CellForeColor <> CouleurPlan Then .CellForeColor = CouleurPlan
        If .CellAlignment <> TypeAlignement Then .CellAlignment = TypeAlignement
        
        '--- texte ---
        If .Text <> Texte Then .Text = Texte
        
        '--- passage à la ligne suivante ---
        If .Row < Pred(.Rows) Then .Row = .Row + 1
    
    End With

End Sub

'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Rôle      : Affiche l'analyse des charges en ligne
' Entrées : NumLigneDepart -> Numéro de ligne de départ de l'affichage
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Private Sub AfficheAnalyseChargesEnLigne(ByVal NumLigneDepart As Long)
    
    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next

    '--- constantes privées ---
    
    '--- déclaration ---
    Dim ChargeAuPosteDetectee As Boolean
    
    Dim a As Integer                                                                                    'pour les boucles FOR...NEXT
    Dim b As Integer                                                                                    'pour les boucles FOR...NEXT
    
    Dim AnalyseChoisie As ANALYSES_MOTEUR_INFERENCE
    
    Dim Texte As String
    
    Dim FicheOrdreSortiePonts As VarOrdreSortieCharges

    '--- affectation de l'analyse choisie ---
    AnalyseChoisie = ANALYSES_MOTEUR_INFERENCE.A_CHARGES_EN_LIGNE

    '--- affectation de la ligne de départ ---
    MSHFGAnalysesMoteurInference(AnalyseChoisie).Row = NumLigneDepart
                
    '*************************************************************************************************************
    '                                              Affichage de l'ordre de sortie des charges
    '*************************************************************************************************************
    For a = PONTS.P_1 To PONTS.P_2
            
        '******************************************************************************************************
        '                                                       Affichage du pont
        '******************************************************************************************************
        Texte = "CHARGES DANS LES BAINS DESTINEES AU PONT " & a
        AfficheTexte AnalyseChoisie, COULEURS.CYAN_5, COULEURS.BLANC, flexAlignCenterCenter, Texte
                
        '--- RAZ d'une charge au moins à un poste a été détectée ---
        ChargeAuPosteDetectee = False
    
        For b = 1 To CHARGES.C_NUM_MAXI
        
            '--- affectation ---
            FicheOrdreSortiePonts = TMoteurInference.TOrdreSortiePonts(a, b)
            
            If FicheOrdreSortiePonts.NumPoste >= PREMIER_BAIN And _
               FicheOrdreSortiePonts.NumPoste <= DERNIER_POSTE And _
               IsNumeric(FicheOrdreSortiePonts.DecompteDuTempsAuPosteReelSecondes) = True Then
                
                '******************************************************************************************************
                '                                                       Affichage du poste
                '******************************************************************************************************
                Texte = " - Au poste " & TEtatsPostes(FicheOrdreSortiePonts.NumPoste).DefinitionPoste.NomPoste & _
                            " (" & TEtatsPostes(FicheOrdreSortiePonts.NumPoste).DefinitionPoste.LibellePoste & ")"
                AfficheTexte AnalyseChoisie, COULEURS.JAUNE_0, COULEURS.NOIR, flexAlignLeftCenter, Texte
                
                '******************************************************************************************************
                '                                           Affichage de la condamnation du poste
                '******************************************************************************************************
                If FicheOrdreSortiePonts.Condamnation = True Then
                    Texte = Space(10) & "- ATTENTION CE POSTE EST CONDAMNE"
                    AfficheTexte AnalyseChoisie, COULEURS.JAUNE_0, COULEURS.ROUGE_4, flexAlignLeftCenter, Texte
                End If
                
                '******************************************************************************************************
                '                                              Affichage du décompte au poste
                '******************************************************************************************************
                Texte = Space(10) & "- Le décompte du temps au poste est de " & CTemps(CLng(FicheOrdreSortiePonts.DecompteDuTempsAuPosteReelSecondes))
                AfficheTexte AnalyseChoisie, COULEURS.JAUNE_0, COULEURS.BLEU_4, flexAlignLeftCenter, Texte
                            
                '******************************************************************************************************
                '                                               Affichage du numéro de charge
                '******************************************************************************************************
                Texte = Space(10) & "- " & TextePourUneCharge(FicheOrdreSortiePonts.NumCharge)
                AfficheTexte AnalyseChoisie, COULEURS.JAUNE_0, COULEURS.BLEU_4, flexAlignLeftCenter, Texte
            
                '--- affectation, une charge au moins à un poste a été détectée ---
                ChargeAuPosteDetectee = True
            
            End If
    
        Next b
    
        '--- affichage du message si l'ordre de sortie des charges n'est pas défini ---
        If ChargeAuPosteDetectee = False Then
            AfficheTexte AnalyseChoisie, COULEURS.JAUNE_0, COULEURS.NOIR, flexAlignCenterCenter, "ORDRE DE SORTIE DES CHARGES INEXISTANT - PAS DE CHARGE DANS UN DES POSTES POUR CE PONT"
            AfficheTexte AnalyseChoisie, COULEURS.JAUNE_0, COULEURS.NOIR, flexAlignLeftCenter, ""
        Else
            AfficheTexte AnalyseChoisie, COULEURS.JAUNE_0, COULEURS.NOIR, flexAlignLeftCenter, ""
            AfficheTexte AnalyseChoisie, COULEURS.JAUNE_0, COULEURS.NOIR, flexAlignLeftCenter, ""
        End If
    
    Next a
    
    '*************************************************************************************************************
    '                                 Affichage de l'indication d'une charge sur au moins un pont
    '*************************************************************************************************************
    Texte = ""
    For a = LBound(TEtatsPonts()) To UBound(TEtatsPonts())
        If TEtatsPonts(a).NumCharge >= CHARGES.C_NUM_MINI And _
           TEtatsPonts(a).NumCharge <= CHARGES.C_NUM_MAXI Then
            Texte = Texte & "CHARGE n° " & TEtatsPonts(a).NumCharge & " EN TRANSFERT sur le PONT " & a
        Else
            Texte = Texte & "PAS de CHARGE sur le PONT " & a
        End If
        If a = PONTS.P_1 Then
            Texte = Texte & ", "
        End If
    Next a
    AfficheTexte AnalyseChoisie, COULEURS.JAUNE_0, COULEURS.ROUGE_4, flexAlignCenterCenter, Texte
    AfficheTexte AnalyseChoisie, COULEURS.JAUNE_0, COULEURS.ROUGE_4, flexAlignCenterCenter, "(Voir l'analyse sur les ponts pour plus d'informations)"
    
End Sub

'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Rôle      : Affiche une charge
' Entrées : NumCharge -> Numéro de la charge
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Private Function TextePourUneCharge(ByVal NumCharge As Integer) As String
    
    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next

    '--- constantes privées ---

    '--- déclaration ---
    Dim a As Integer
    Dim EnsembleNumCommandesInternes As String
    
    '--- affectation par défaut ---
    TextePourUneCharge = ""
    
    If NumCharge >= CHARGES.C_NUM_MINI And NumCharge <= CHARGES.C_NUM_MAXI Then
    
        '--- construction de la chaine des numéros des commandes internes ---
        With TEtatsCharges(NumCharge)
            For a = LBound(.TDetailsCharges()) To UBound(.TDetailsCharges())
                With TEtatsCharges(NumCharge).TDetailsCharges(a)
                    If .NumCommandeInterne > 0 Then
                        EnsembleNumCommandesInternes = EnsembleNumCommandesInternes & .NumCommandeInterne & "/"
                    Else
                        Exit For
                    End If
                End With
            Next a
        End With

        '--- construction de la valeur de retour ---
        If EnsembleNumCommandesInternes = "" Then
            
            '--- affectation du texte ---
            TextePourUneCharge = "Charge n° " & NumCharge & ", SANS COMMANDE INTERNE"
            
        Else
            
            '--- suppression du dernier séparateur ---
            EnsembleNumCommandesInternes = Left(EnsembleNumCommandesInternes, Pred(Len(EnsembleNumCommandesInternes)))
            
            '--- affectation du texte ---
            TextePourUneCharge = "Charge n° " & NumCharge & ", commande(s) interne(s) " & EnsembleNumCommandesInternes
        
        End If
                    
    Else
                    
        '--- affectation du texte ---
        TextePourUneCharge = "PAS DE CHARGE"
                    
    End If

End Function

'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Rôle      : Affiche l'analyse des ponts
' Entrées : NumLigneDepart -> Numéro de ligne de départ de l'affichage
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Private Sub AfficheAnalysePonts(ByVal NumLigneDepart As Long)
    
    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next

    '--- constantes privées ---
    
    '--- déclaration ---
    Dim a As Integer

    Dim AnalyseChoisie As ANALYSES_MOTEUR_INFERENCE
    
    Dim Texte As String
    
    '--- affectation de l'analyse choisie ---
    AnalyseChoisie = ANALYSES_MOTEUR_INFERENCE.A_PONTS

    '--- affectation de la ligne de départ ---
    MSHFGAnalysesMoteurInference(AnalyseChoisie).Row = NumLigneDepart
                
    '*************************************************************************************************************
    '                                              Affichage d'une partie des états des ponts
    '*************************************************************************************************************
    For a = LBound(TEtatsPonts()) To UBound(TEtatsPonts())
                    
        '******************************************************************************************************
        '                                                      Affichage du nom du pont
        '******************************************************************************************************
        Texte = " - PONT " & a
        AfficheTexte AnalyseChoisie, COULEURS.ORANGE_0, COULEURS.NOIR, flexAlignLeftCenter, Texte
        
        '******************************************************************************************************
        '                                                        Condamnation du pont
        '******************************************************************************************************
        If TEtatsPonts(a).Condamnation = True Then
            Texte = Space(10) & "- Le pont est CONDAMNE"
            AfficheTexte AnalyseChoisie, COULEURS.ORANGE_0, COULEURS.ROUGE_4, flexAlignLeftCenter, Texte
        End If
        
        '******************************************************************************************************
        '                                                              Mode du pont
        '******************************************************************************************************
        Select Case TEtatsPonts(a).ModePont
            Case MODES_PONTS.M_MAINTENANCE
                Texte = Space(10) & "- Le pont est en MAINTENANCE"
                AfficheTexte AnalyseChoisie, COULEURS.ORANGE_0, COULEURS.ROUGE_4, flexAlignLeftCenter, Texte
            Case MODES_PONTS.M_MANUEL
                Texte = Space(10) & "- Le pont est en MANUEL"
                AfficheTexte AnalyseChoisie, COULEURS.ORANGE_0, COULEURS.ROUGE_4, flexAlignLeftCenter, Texte
            Case MODES_PONTS.M_SEMI_AUTOMATIQUE
                Texte = Space(10) & "- Le pont est en SEMI-AUTOMATIQUE"
                AfficheTexte AnalyseChoisie, COULEURS.ORANGE_0, COULEURS.ROUGE_4, flexAlignLeftCenter, Texte
            Case MODES_PONTS.M_AUTOMATIQUE
                Texte = Space(10) & "- Le pont est en AUTOMATIQUE"
                AfficheTexte AnalyseChoisie, COULEURS.ORANGE_0, COULEURS.BLEU_4, flexAlignLeftCenter, Texte
            Case Else
                Texte = Space(10) & "- Mode du pont NON DEFINI"
                AfficheTexte AnalyseChoisie, COULEURS.ORANGE_0, COULEURS.ROUGE_4, flexAlignLeftCenter, Texte
        End Select
        
        '******************************************************************************************************
        '                                                         Type de séquence
        '******************************************************************************************************
        Select Case TEtatsPonts(a).TypeSequence
            Case TYPES_SEQUENCES.TS_CYCLIQUE_PAR_IMPULSIONS
                Texte = Space(10) & "- Le type de séquence appliqué au pont est le CYCLIQUE PAR IMPLUSIONS"
                AfficheTexte AnalyseChoisie, COULEURS.ORANGE_0, COULEURS.BLEU_4, flexAlignLeftCenter, Texte
            Case TYPES_SEQUENCES.TS_CYCLIQUE_OPTIMISE
                Texte = Space(10) & "- Le type de séquence appliqué au pont est le CYCLIQUE OPTIMISE"
                AfficheTexte AnalyseChoisie, COULEURS.ORANGE_0, COULEURS.BLEU_4, flexAlignLeftCenter, Texte
            Case TYPES_SEQUENCES.TS_ALEATOIRE
                Texte = Space(10) & "- Le type de séquence appliqué au pont est ALEATOIRE"
                AfficheTexte AnalyseChoisie, COULEURS.ORANGE_0, COULEURS.BLEU_4, flexAlignLeftCenter, Texte
            Case Else
                Texte = Space(10) & "- Pas de définition du TYPE de SEQUENCE"
                AfficheTexte AnalyseChoisie, COULEURS.ORANGE_0, COULEURS.ROUGE_4, flexAlignLeftCenter, Texte
        End Select

        '******************************************************************************************************
        '                                                   Contrôle par l'opérateur
        '******************************************************************************************************
        If TEtatsPonts(a).ModePont = MODES_PONTS.M_AUTOMATIQUE Then
            If TEtatsPonts(a).ControleParOperateur = False Then
                Texte = Space(10) & "- Ce pont est actuellement géré par l'ORDINATEUR"
                AfficheTexte AnalyseChoisie, COULEURS.ORANGE_0, COULEURS.BLEU_4, flexAlignLeftCenter, Texte
            Else
                Texte = Space(10) & "- Ce pont est actuellement sous le CONTROLE DE L'OPERATEUR avec l'ORDINATEUR"
                AfficheTexte AnalyseChoisie, COULEURS.ORANGE_0, COULEURS.ROUGE_4, flexAlignLeftCenter, Texte
            End If
        Else
            Texte = Space(10) & "- Ce pont est actuellement sous le CONTROLE DE L'OPERATEUR avec la BOITE A BOUTONS"
            AfficheTexte AnalyseChoisie, COULEURS.ORANGE_0, COULEURS.ROUGE_4, flexAlignLeftCenter, Texte
        End If
        
        '******************************************************************************************************
        '                            Paramètres des cycles des ponts pour le CYCLE ACTUEL
        '******************************************************************************************************
        If TEtatsPonts(a).ModePont = MODES_PONTS.M_AUTOMATIQUE Then
            If TEtatsPonts(a).TParametresCyclesPonts(CYCLES.C_ACTUEL).NumPosteDepart >= POSTES.P_CHGT_1 And _
               TEtatsPonts(a).TParametresCyclesPonts(CYCLES.C_ACTUEL).NumPosteDepart <= DERNIER_POSTE And _
               TEtatsPonts(a).TParametresCyclesPonts(CYCLES.C_ACTUEL).NumPosteArrivee >= POSTES.P_CHGT_1 And _
               TEtatsPonts(a).TParametresCyclesPonts(CYCLES.C_ACTUEL).NumPosteArrivee <= DERNIER_POSTE Then
            
                '--- affectation du texte de base ---
                Select Case TEtatsPonts(a).TParametresCyclesPonts(CYCLES.C_ACTUEL).TypeCycle
                    
                    Case TYPES_CYCLES.TC_DEPLACEMENT_PONT
                        '--- déplacement du pont ---
                        Texte = Space(10) & _
                                     "- DEPLACEMENT du PONT du poste " & _
                                     TEtatsPostes(TEtatsPonts(a).TParametresCyclesPonts(CYCLES.C_ACTUEL).NumPosteDepart).DefinitionPoste.NomPoste & _
                                     " au poste " & _
                                     TEtatsPostes(TEtatsPonts(a).TParametresCyclesPonts(CYCLES.C_ACTUEL).NumPosteArrivee).DefinitionPoste.NomPoste
                    
                    Case TYPES_CYCLES.TC_TRANSFERT_CHARGE
                        '--- transfert d'une charge ---
                        Texte = Space(10) & _
                                     "- TRANSFERT de la CHARGE du poste " & _
                                     TEtatsPostes(TEtatsPonts(a).TParametresCyclesPonts(CYCLES.C_ACTUEL).NumPosteDepart).DefinitionPoste.NomPoste & _
                                     " au poste " & _
                                     TEtatsPostes(TEtatsPonts(a).TParametresCyclesPonts(CYCLES.C_ACTUEL).NumPosteArrivee).DefinitionPoste.NomPoste
                                                 
                    Case Else
                        '--- tous les autres cas ---
                        Texte = ""
                
                End Select
        
                '--- affichage du transfert ---
                If Texte <> "" Then
                    If TEtatsPonts(a).PtrEtActionEnCoursAPI.PtrAction > 0 Then
                        Texte = Texte & " EN COURS"
                    Else
                        Texte = Texte & " EFFECTUE"
                    End If
                    AfficheTexte AnalyseChoisie, COULEURS.ORANGE_0, COULEURS.BLEU_4, flexAlignLeftCenter, Texte
                End If
            
            Else
                Texte = Space(10) & "- Déplacement du pont ou transfert de charge NON DEFINI"
                AfficheTexte AnalyseChoisie, COULEURS.ORANGE_0, COULEURS.ROUGE_4, flexAlignLeftCenter, Texte
            End If
        Else
            Texte = Space(10) & "- Déplacement du pont ou transfert de charge NON DEFINI"
            AfficheTexte AnalyseChoisie, COULEURS.ORANGE_0, COULEURS.ROUGE_4, flexAlignLeftCenter, Texte
        End If
        
        '******************************************************************************************************
        '                            Paramètres des cycles des ponts pour le PROCHAIN CYCLE
        '******************************************************************************************************
        If TEtatsPonts(a).ModePont = MODES_PONTS.M_AUTOMATIQUE Then
            If TEtatsPonts(a).TParametresCyclesPonts(CYCLES.C_PROCHAIN).NumPosteDepart >= POSTES.P_CHGT_1 And _
               TEtatsPonts(a).TParametresCyclesPonts(CYCLES.C_PROCHAIN).NumPosteDepart <= DERNIER_POSTE And _
               TEtatsPonts(a).TParametresCyclesPonts(CYCLES.C_PROCHAIN).NumPosteArrivee >= POSTES.P_CHGT_1 And _
               TEtatsPonts(a).TParametresCyclesPonts(CYCLES.C_PROCHAIN).NumPosteArrivee <= DERNIER_POSTE Then
            
                '--- affectation du texte de base ---
                Select Case TEtatsPonts(a).TParametresCyclesPonts(CYCLES.C_PROCHAIN).TypeCycle
                    
                    Case TYPES_CYCLES.TC_DEPLACEMENT_PONT
                        '--- déplacement du pont ---
                        Texte = Space(10) & _
                                     "- LE PROCHAIN DEPLACEMENT du PONT se fera du poste " & _
                                     TEtatsPostes(TEtatsPonts(a).TParametresCyclesPonts(CYCLES.C_PROCHAIN).NumPosteDepart).DefinitionPoste.NomPoste & _
                                     " au poste " & _
                                     TEtatsPostes(TEtatsPonts(a).TParametresCyclesPonts(CYCLES.C_PROCHAIN).NumPosteArrivee).DefinitionPoste.NomPoste
                    
                    Case TYPES_CYCLES.TC_TRANSFERT_CHARGE
                        '--- transfert d'une charge ---
                        Texte = Space(10) & _
                                     "- LE PROCHAIN TRANSFERT de la CHARGE se fera du poste " & _
                                     TEtatsPostes(TEtatsPonts(a).TParametresCyclesPonts(CYCLES.C_PROCHAIN).NumPosteDepart).DefinitionPoste.NomPoste & _
                                     " au poste " & _
                                     TEtatsPostes(TEtatsPonts(a).TParametresCyclesPonts(CYCLES.C_PROCHAIN).NumPosteArrivee).DefinitionPoste.NomPoste
                                                 
                    Case Else
                        '--- tous les autres cas ---
                        Texte = ""
                
                End Select
                                 
                '--- affichage du transfert ---
                If Texte <> "" Then
                    AfficheTexte AnalyseChoisie, COULEURS.ORANGE_0, COULEURS.BLEU_4, flexAlignLeftCenter, Texte
                End If
            
            Else
                Texte = Space(10) & "- Prochain cycle NON DEFINI"
                AfficheTexte AnalyseChoisie, COULEURS.ORANGE_0, COULEURS.ROUGE_4, flexAlignLeftCenter, Texte
            End If
        Else
            Texte = Space(10) & "- Prochain cycle NON DEFINI"
            AfficheTexte AnalyseChoisie, COULEURS.ORANGE_0, COULEURS.ROUGE_4, flexAlignLeftCenter, Texte
        End If
        
        '******************************************************************************************************
        '                                                          Charge sur le pont
        '******************************************************************************************************
        If TEtatsPonts(a).NumCharge >= CHARGES.C_NUM_MINI And TEtatsPonts(a).NumCharge <= CHARGES.C_NUM_MAXI Then
            Texte = Space(10) & "- " & TextePourUneCharge(TEtatsPonts(a).NumCharge)
        Else
            Texte = Space(10) & "- PAS DE CHARGE SUR LE PONT"
        End If
        AfficheTexte AnalyseChoisie, COULEURS.ORANGE_0, COULEURS.BLEU_4, flexAlignLeftCenter, Texte
        
        '--- ajout d'une ligne ---
        If a = PONTS.P_1 Then
            AfficheTexte AnalyseChoisie, COULEURS.BLANC, COULEURS.BLEU_4, flexAlignLeftCenter, ""
        End If

    Next a
        
End Sub

'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Rôle      : Affiche l'analyse du déchargement
' Entrées : NumLigneDepart -> Numéro de ligne de départ de l'affichage
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Private Sub AfficheAnalyseDechargement(ByVal NumLigneDepart As Long)
    
    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next

    '--- constantes privées ---
    
    '--- déclaration ---
    Dim CptChariotsAbsents As Integer           'compteur des chariots absents
    
    Dim a As Integer
    
    Dim AnalyseChoisie As ANALYSES_MOTEUR_INFERENCE
    
    Dim Texte As String
    
    '--- affectation de l'analyse choisie ---
    AnalyseChoisie = ANALYSES_MOTEUR_INFERENCE.A_DECHARGEMENT

    '--- affectation de la ligne de départ ---
   MSHFGAnalysesMoteurInference(AnalyseChoisie).Row = NumLigneDepart
                    
    '*************************************************************************************************************
    '                                               Affichage de l'état du déchargement
    '*************************************************************************************************************
    For a = POSTES.P_D1 To POSTES.P_D2
        
        Select Case TEtatsPostes(a).EtatsChariots
        
            Case ETATS_CHARIOTS.E_ABSENT
                '--- chariot absent ---
                Inc CptChariotsAbsents
            
            Case ETATS_CHARIOTS.E_PRESENT
                '--- chariot présent (non verrouillé) ---
                Texte = Space(10) & "- Chariot PRESENT, NON VERROUILLE SANS CHARGE"
                AfficheTexte AnalyseChoisie, COULEURS.BLEU_0, COULEURS.ROUGE_4, flexAlignLeftCenter, Texte
            
            Case ETATS_CHARIOTS.E_PRESENT_VERROUILLE
                '--- chariot présent verrouillé ---
            
                '************************************************************************************************
                '                                                      Affichage du poste
                '************************************************************************************************
                Texte = " - Au poste " & TEtatsPostes(a).DefinitionPoste.NomPoste & _
                            " (" & TEtatsPostes(a).DefinitionPoste.LibellePoste & ")"
                AfficheTexte AnalyseChoisie, COULEURS.BLEU_0, COULEURS.NOIR, flexAlignLeftCenter, Texte
         
                '************************************************************************************************
                '                                       Affichage de la condamnation du poste
                '************************************************************************************************
                If TEtatsPostes(a).Condamnation = True Then
                    Texte = Space(10) & "- ATTENTION CE POSTE EST CONDAMNE"
                    AfficheTexte AnalyseChoisie, COULEURS.BLEU_0, COULEURS.ROUGE_4, flexAlignLeftCenter, Texte
                End If
         
                '************************************************************************************************
                '                                                  Affichage de la charge
                '************************************************************************************************
                If TEtatsPostes(a).NumCharge >= CHARGES.C_NUM_MINI And _
                   TEtatsPostes(a).NumCharge <= CHARGES.C_NUM_MAXI Then
                    Texte = Space(10) & "- Chariot PRESENT, VERROUILLE AVEC LA CHARGE n° " & TEtatsPostes(a).NumCharge
                    AfficheTexte AnalyseChoisie, COULEURS.BLEU_0, COULEURS.BLEU_4, flexAlignLeftCenter, Texte
                    Texte = Space(10) & "- " & TextePourUneCharge(TEtatsPostes(a).NumCharge)
                    AfficheTexte AnalyseChoisie, COULEURS.BLEU_0, COULEURS.BLEU_4, flexAlignLeftCenter, Texte
                Else
                    Texte = Space(10) & "- Chariot PRESENT, VERROUILLE SANS CHARGE"
                    AfficheTexte AnalyseChoisie, COULEURS.BLEU_0, COULEURS.BLEU_4, flexAlignLeftCenter, Texte
                End If
         
            Case Else
         
        End Select
         
    Next a

    '************************************************************************************************
    '                                        Affichage en cas de déchargement vide
    '************************************************************************************************
    If CptChariotsAbsents = 6 Then
        Texte = " LE DECHARGEMENT EST VIDE"
        AfficheTexte AnalyseChoisie, COULEURS.BLEU_0, COULEURS.NOIR, flexAlignLeftCenter, Texte
    End If
    
End Sub

'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Rôle      : Affiche l'analyse du chargement
' Entrées : NumLigneDepart -> Numéro de ligne de départ de l'affichage
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Private Sub AfficheAnalyseChargement(ByVal NumLigneDepart As Long)

    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next

    '--- constantes privées ---
    
    '--- déclaration ---
    Dim a As Integer
    
    Dim NumPosteAnalyse As Integer
    
    Dim AnalyseChoisie As ANALYSES_MOTEUR_INFERENCE
    
    Dim Texte As String
    
    '--- affectation de l'analyse choisie ---
    AnalyseChoisie = ANALYSES_MOTEUR_INFERENCE.A_CHARGEMENT

    '--- affectation de la ligne de départ ---
    MSHFGAnalysesMoteurInference(AnalyseChoisie).Row = NumLigneDepart
                
    '*************************************************************************************************************
    '            Affichage du prochain n° de poste de chargement si le poste d'anodisation C13 est imposé
    '*************************************************************************************************************
    NumPosteAnalyse = TMoteurInference.ProchainNumPosteChargementSiAnodisationC13Impose
    If NumPosteAnalyse = 0 Then
        Texte = " - Pas de charge avec choix du poste d'anodisation sur C13 IMPOSE"
        AfficheTexte AnalyseChoisie, COULEURS.CYAN_0, COULEURS.NOIR, flexAlignLeftCenter, Texte
    Else
        Texte = " - Sélection de la charge avec choix du poste d'anodisation sur C13 IMPOSE en " & TEtatsPostes(NumPosteAnalyse).DefinitionPoste.NomPoste
        AfficheTexte AnalyseChoisie, COULEURS.CYAN_0, COULEURS.BLEU_4, flexAlignLeftCenter, Texte
    End If
        
    '*************************************************************************************************************
    '            Affichage du prochain n° de poste de chargement si le poste d'anodisation C14 est imposé
    '*************************************************************************************************************
    NumPosteAnalyse = TMoteurInference.ProchainNumPosteChargementSiAnodisationC14Impose
    If NumPosteAnalyse = 0 Then
        Texte = " - Pas de charge avec choix du poste d'anodisation sur C14 IMPOSE"
        AfficheTexte AnalyseChoisie, COULEURS.CYAN_0, COULEURS.NOIR, flexAlignLeftCenter, Texte
    Else
        Texte = " - Sélection de la charge avec choix du poste d'anodisation sur C14 IMPOSE en " & TEtatsPostes(NumPosteAnalyse).DefinitionPoste.NomPoste
        AfficheTexte AnalyseChoisie, COULEURS.CYAN_0, COULEURS.BLEU_4, flexAlignLeftCenter, Texte
    End If
    
    '*************************************************************************************************************
    '            Affichage du prochain n° de poste de chargement si le poste d'anodisation C15 est imposé
    '*************************************************************************************************************
    NumPosteAnalyse = TMoteurInference.ProchainNumPosteChargementSiAnodisationC15Impose
    If NumPosteAnalyse = 0 Then
        Texte = " - Pas de charge avec choix du poste d'anodisation sur C15 IMPOSE"
        AfficheTexte AnalyseChoisie, COULEURS.CYAN_0, COULEURS.NOIR, flexAlignLeftCenter, Texte
    Else
        Texte = " - Sélection de la charge avec choix du poste d'anodisation sur C15 IMPOSE en " & TEtatsPostes(NumPosteAnalyse).DefinitionPoste.NomPoste
        AfficheTexte AnalyseChoisie, COULEURS.CYAN_0, COULEURS.BLEU_4, flexAlignLeftCenter, Texte
    End If
    
    '*************************************************************************************************************
    '            Affichage du prochain n° de poste de chargement si le poste d'anodisation C16 est imposé
    '*************************************************************************************************************
    NumPosteAnalyse = TMoteurInference.ProchainNumPosteChargementSiAnodisationC16Impose
    If NumPosteAnalyse = 0 Then
        Texte = " - Pas de charge avec choix du poste d'anodisation sur C16 IMPOSE"
        AfficheTexte AnalyseChoisie, COULEURS.CYAN_0, COULEURS.NOIR, flexAlignLeftCenter, Texte
    Else
        Texte = " - Sélection de la charge avec choix du poste d'anodisation sur C16 IMPOSE en " & TEtatsPostes(NumPosteAnalyse).DefinitionPoste.NomPoste
        AfficheTexte AnalyseChoisie, COULEURS.CYAN_0, COULEURS.BLEU_4, flexAlignLeftCenter, Texte
    End If
    
    '*************************************************************************************************************
    '        Affichage du prochain n° de poste de chargement si le poste d'anodisation est automatique
    '*************************************************************************************************************
    NumPosteAnalyse = TMoteurInference.ProchainNumPosteChargementSiAnodisationAutomatique
    If NumPosteAnalyse = 0 Then
        Texte = " - Pas de charge avec choix du poste d'anodisation sur AUTOMATIQUE"
        AfficheTexte AnalyseChoisie, COULEURS.CYAN_0, COULEURS.NOIR, flexAlignLeftCenter, Texte
    Else
        Texte = " - Sélection de la charge avec choix du poste d'anodisation sur AUTOMATIQUE en " & TEtatsPostes(NumPosteAnalyse).DefinitionPoste.NomPoste
        AfficheTexte AnalyseChoisie, COULEURS.CYAN_0, COULEURS.BLEU_4, flexAlignLeftCenter, Texte
    End If
    
    '*************************************************************************************************************
    '                               Affichage de la prochaine charge à rentrer dans la ligne
    '*************************************************************************************************************
    AfficheTexte AnalyseChoisie, COULEURS.CYAN_0, COULEURS.NOIR, flexAlignLeftCenter, ""
    NumPosteAnalyse = TMoteurInference.ProchainNumPosteChargement
    If NumPosteAnalyse = 0 Then
        Texte = " - Pas de décision de prendre une charge"
        AfficheTexte AnalyseChoisie, COULEURS.CYAN_0, COULEURS.NOIR, flexAlignLeftCenter, Texte
    Else
        Texte = " - Décision de prendre la charge du poste " & TEtatsPostes(NumPosteAnalyse).DefinitionPoste.NomPoste
        AfficheTexte AnalyseChoisie, COULEURS.CYAN_0, COULEURS.ROUGE_3, flexAlignLeftCenter, Texte
    End If
    
    '*************************************************************************************************************
    '                        Affichage de la condamnation des postes C13, C14, C15, C16
    '*************************************************************************************************************
    AfficheTexte AnalyseChoisie, COULEURS.BLANC, COULEURS.NOIR, flexAlignLeftCenter, ""
    For a = POSTES.P_C13 To POSTES.P_C16
        If TEtatsPostes(a).Condamnation = True Then
            Texte = " - ATTENTION, LE POSTE " & TEtatsPostes(a).DefinitionPoste.NomPoste & " EST CONDAMNE"
            AfficheTexte AnalyseChoisie, COULEURS.CYAN_0, COULEURS.ROUGE_3, flexAlignLeftCenter, Texte
        End If
    Next a
    
End Sub

'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Rôle      : Gestion du moteur d'inférence (affichage des analyses effectuées chaque seconde)
' Entrées : AnalyseChoisie -> Analyse choisie fonction de l'énumération ANALYSES_MOTEUR_INFERENCE
'                     EtatSouhaite -> Fonction de l'énumération GESTION_GRILLES
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Private Sub GestionMoteurInference(ByVal AnalyseChoisie As ANALYSES_MOTEUR_INFERENCE, _
                                                           ByVal EtatSouhaite As GESTION_GRILLES)
    
    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next

    '--- constantes privées ---
    Const EPAISSEUR_CARACTERE As Integer = 140
    
    '--- déclaration ---
    Dim a As Integer, _
            b As Integer, _
            NumLigneDebutAnalyseEnCours As Integer
    Dim Texte As String
        
    Select Case EtatSouhaite

        Case GESTION_GRILLES.GG_INITIALISATION
            '--- initialisation de la grille ---
            With MSHFGAnalysesMoteurInference(AnalyseChoisie)
                
                .Redraw = False

                .Clear

                .FixedCols = 1
                .FixedRows = 0
                
                '--- affectation du nombre de lignes de chaque tableau ---
                Select Case AnalyseChoisie
                    Case ANALYSES_MOTEUR_INFERENCE.A_CHARGEMENT: .Rows = NBR_LIGNES_MOTEUR_INFERENCE.NL_CHARGEMENT + 1
                    Case ANALYSES_MOTEUR_INFERENCE.A_DECHARGEMENT: .Rows = NBR_LIGNES_MOTEUR_INFERENCE.NL_DECHARGEMENT + 1
                    Case ANALYSES_MOTEUR_INFERENCE.A_PONTS: .Rows = NBR_LIGNES_MOTEUR_INFERENCE.NL_PONTS + 1
                    Case ANALYSES_MOTEUR_INFERENCE.A_CHARGES_EN_LIGNE: .Rows = NBR_LIGNES_MOTEUR_INFERENCE.NL_CHARGES_EN_LIGNE + 1
                    Case Else
                End Select
                    
                .Cols = NBR_COLONNES_ANALYSES_MOTEUR_INFERENCE + .FixedCols
                .RowHeightMin = 315
                .Row = 0

                '--- paramétrages de chaque colonne ---
                .Col = COLONNES_ANALYSES_MOTEUR_INFERENCE.C_NUM_LIGNES
                .ColWidth(.Col) = 3 * EPAISSEUR_CARACTERE
                .ColAlignment(.Col) = flexAlignRightCenter
                
                .Col = COLONNES_ANALYSES_MOTEUR_INFERENCE.C_DONNEES_1
                .ColWidth(.Col) = 2 * EPAISSEUR_CARACTERE
                .ColAlignment(.Col) = flexAlignCenterCenter
                
                .Col = COLONNES_ANALYSES_MOTEUR_INFERENCE.C_DONNEES_2
                .ColWidth(.Col) = 90.7 * EPAISSEUR_CARACTERE
                .ColAlignment(.Col) = flexAlignCenterCenter
                
                .Col = COLONNES_ANALYSES_MOTEUR_INFERENCE.C_DONNEES_3
                .ColWidth(.Col) = 2 * EPAISSEUR_CARACTERE
                .ColAlignment(.Col) = flexAlignCenterCenter
                
                '--- centrage des titres ---
                .Row = 0
                For a = 1 To Pred(.Cols)
                    .Col = a
                    .CellAlignment = flexAlignCenterCenter
                Next a

                '--- N° de lignes, vidage des champs ---
                For a = 0 To .Rows - 1
                
                    '--- N° de lignes ---
                    .Col = COLONNES_ANALYSES_MOTEUR_INFERENCE.C_NUM_LIGNES
                    .Row = a
                    .Text = IIf(a = 0, "-  ", CStr(a))
                
                    '--- couleurs des lignes ---
                    .Col = COLONNES_ANALYSES_MOTEUR_INFERENCE.C_DONNEES_1
                    .FillStyle = flexFillRepeat
                    .ColSel = COLONNES_ANALYSES_MOTEUR_INFERENCE.C_DONNEES_2
                    .CellBackColor = COULEURS.BLANC
                
                Next a

                '--- fixer le curseur ---
                .Row = 1
                .Col = COLONNES_ANALYSES_MOTEUR_INFERENCE.C_DONNEES_1

                .Redraw = True
                        
            End With

        Case GESTION_GRILLES.GG_VIDAGE
            '--- vidage de la grille ---

        Case GESTION_GRILLES.GG_AFFICHAGE
            '--- construction de la grille ---
            Select Case AnalyseChoisie
            
                Case ANALYSES_MOTEUR_INFERENCE.A_CHARGEMENT
                    '***********************************************************************************************
                    '                                               ANALYSE DU CHARGEMENT
                    '***********************************************************************************************
            
                    '--- affichage de l'analyse du chargement ---
                    AfficheAnalyseChargement NumLigneDepart:=1
    
                    '--- effacement du reste de la grille ---
                    With MSHFGAnalysesMoteurInference(AnalyseChoisie)
                        For a = .Row To Pred(.Rows)
                            AfficheTexte AnalyseChoisie, COULEURS.BLANC, COULEURS.NOIR, flexAlignLeftCenter, ""
                        Next a
                    End With
    
                Case ANALYSES_MOTEUR_INFERENCE.A_DECHARGEMENT
                    '***********************************************************************************************
                    '                                            ANALYSE DU DECHARGEMENT
                    '***********************************************************************************************
                    
                    '--- affichage de l'analyse du déchargement ---
                    AfficheAnalyseDechargement NumLigneDepart:=1
                    
                    '--- effacement du reste de la grille ---
                    With MSHFGAnalysesMoteurInference(AnalyseChoisie)
                        For a = .Row To Pred(.Rows)
                           AfficheTexte AnalyseChoisie, COULEURS.BLANC, COULEURS.NOIR, flexAlignLeftCenter, ""
                        Next a
                    End With
                
                Case ANALYSES_MOTEUR_INFERENCE.A_PONTS
                    '***********************************************************************************************
                    '                                                   ANALYSE DES PONTS
                    '***********************************************************************************************
                    
                    '--- affichage de l'analyse des ponts ---
                    AfficheAnalysePonts NumLigneDepart:=1
                    
                    '--- effacement du reste de la grille ---
                    With MSHFGAnalysesMoteurInference(AnalyseChoisie)
                        For a = .Row To Pred(.Rows)
                           AfficheTexte AnalyseChoisie, COULEURS.BLANC, COULEURS.NOIR, flexAlignLeftCenter, ""
                        Next a
                    End With
                
                Case ANALYSES_MOTEUR_INFERENCE.A_CHARGES_EN_LIGNE
                    '***********************************************************************************************
                    '                                         ANALYSE DES CHARGES EN LIGNE
                    '***********************************************************************************************
                    
                    '--- affichage de l'analyse des charges en ligne ---
                    AfficheAnalyseChargesEnLigne NumLigneDepart:=1
                    
                    '--- effacement du reste de la grille ---
                    With MSHFGAnalysesMoteurInference(AnalyseChoisie)
                        For a = .Row To Pred(.Rows)
                            AfficheTexte AnalyseChoisie, COULEURS.BLANC, COULEURS.NOIR, flexAlignLeftCenter, ""
                        Next a
                    End With
                
                Case Else
            End Select
                
        Case Else

    End Select

End Sub



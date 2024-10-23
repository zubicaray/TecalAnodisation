VERSION 5.00
Object = "{27395F88-0C0C-101B-A3C9-08002B2F49FB}#1.1#0"; "PICCLP32.OCX"
Begin VB.Form FAcceuil 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   5370
   ClientLeft      =   9465
   ClientTop       =   3720
   ClientWidth     =   7560
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   5370
   ScaleWidth      =   7560
   ShowInTaskbar   =   0   'False
   Visible         =   0   'False
   Begin VB.Frame fraMainFrame 
      Height          =   3990
      Left            =   0
      TabIndex        =   0
      Top             =   -60
      Width           =   7515
      Begin VB.PictureBox Picture1 
         BackColor       =   &H0000FF00&
         Height          =   1395
         Left            =   4200
         ScaleHeight     =   1335
         ScaleWidth      =   3015
         TabIndex        =   3
         Top             =   1740
         Width           =   3075
         Begin VB.Label LVersion 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00000000&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000FFFF&
            Height          =   210
            Left            =   120
            TabIndex        =   4
            Top             =   1020
            Width           =   2775
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            BackColor       =   &H0000FFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "Ligne d'anodisation"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   -1  'True
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   855
            Left            =   120
            TabIndex        =   5
            Top             =   120
            Width           =   2775
         End
      End
      Begin VB.PictureBox PBLogoAnime 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Height          =   3585
         Index           =   0
         Left            =   120
         Picture         =   "FAcceuil.frx":0000
         ScaleHeight     =   235
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   250
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   225
         Width           =   3810
      End
      Begin VB.Label LEtsVerbrugge 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "TECAL VERBRUGGE"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   465
         Left            =   4080
         TabIndex        =   2
         Top             =   1020
         Width           =   3315
      End
      Begin VB.Image IFondEcran 
         Height          =   7950
         Left            =   -1320
         Picture         =   "FAcceuil.frx":2B292
         Top             =   60
         Width           =   14850
      End
   End
   Begin VB.PictureBox PBBoutons 
      Align           =   2  'Align Bottom
      DrawStyle       =   2  'Dot
      DrawWidth       =   16891
      Height          =   1095
      Left            =   0
      Picture         =   "FAcceuil.frx":2DEAF
      ScaleHeight     =   1035
      ScaleWidth      =   7500
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   4275
      Width           =   7560
      Begin VB.CommandButton CBQuitter 
         BackColor       =   &H00FFFFFF&
         Cancel          =   -1  'True
         Caption         =   "&Quitter"
         DownPicture     =   "FAcceuil.frx":30549
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
         Left            =   6000
         MaskColor       =   &H00FF00FF&
         Picture         =   "FAcceuil.frx":30C4B
         Style           =   1  'Graphical
         TabIndex        =   9
         ToolTipText     =   " Quitter le logiciel "
         Top             =   120
         UseMaskColor    =   -1  'True
         Width           =   1395
      End
      Begin VB.CommandButton CBValider 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Valider"
         DownPicture     =   "FAcceuil.frx":3134D
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   4500
         MaskColor       =   &H00FF00FF&
         Picture         =   "FAcceuil.frx":31A4F
         Style           =   1  'Graphical
         TabIndex        =   8
         ToolTipText     =   " Valider la date et l'heure et accéder au programme "
         Top             =   120
         UseMaskColor    =   -1  'True
         Width           =   1395
      End
      Begin VB.Shape SFocus 
         BorderColor     =   &H000000FF&
         BorderWidth     =   5
         Height          =   345
         Left            =   300
         Top             =   180
         Visible         =   0   'False
         Width           =   540
      End
   End
   Begin VB.Timer TimerDateHeure 
      Interval        =   200
      Left            =   1080
      Top             =   4020
   End
   Begin VB.Timer Timer_Animation 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   540
      Top             =   4020
   End
   Begin PicClip.PictureClip PCAnimation 
      Left            =   1740
      Top             =   4020
      _ExtentX        =   1217
      _ExtentY        =   291
      _Version        =   393216
      Rows            =   2
      Cols            =   5
   End
   Begin VB.Label LDateHeureActuel 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   0
      TabIndex        =   6
      Top             =   3960
      Width           =   7515
   End
End
Attribute VB_Name = "FAcceuil"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Rôle                    : Fenêtre d'acceuil
' Nom                    : FAcceuil.frm
' Date de création : 31/03/1999
' Détails                :
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

'--- déclarations obligatoires ---
Option Explicit

'--- options générales ---
Option Base 1
DefVar A-Z

Private Sub CBQuitter_Click()
    On Error Resume Next
    End
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

Private Sub CBValider_Click()
    
    '--- aiguillage en cas d'erreur ---
    On Error Resume Next
    
    '--- annulation des timers ---
    With Timer_Animation
        .Enabled = False
        .Interval = 0
    End With
    With TimerDateHeure
        .Enabled = False
        .Interval = 0
    End With
    Set PCAnimation.Picture = LoadPicture("")
    
    Unload Me

End Sub

Private Sub CBValider_GotFocus()
    
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

Private Sub CBValider_LostFocus()
    On Error Resume Next
    SFocus.Visible = False
End Sub

Private Sub Form_Activate()
    
    '--- aiguillage en cas d'erreur ---
    On Error Resume Next
    
    '--- lancement du timer ---
    Timer_Animation.Enabled = True

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error Resume Next
    GestionTouches KeyCode, Shift
End Sub

Private Sub Form_Load()
    
    '--- aiguillage en cas d'erreur ---
    On Error Resume Next

    '--- divers sur la fenêtre ---
    LVersion.Caption = "Version " & App.Major & "." & App.Minor
    Centrefenetre Me

    '--- chargement de l'animation ---
    Set PCAnimation.Picture = LoadPicture(App.Path & "\Images\" & "Animation Tecal Verbrugge.BMP")
    DoEvents

End Sub

'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Rôle : Gére les touches du clavier
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Private Sub GestionTouches(KeyCode As Integer, Shift As Integer)
    
    '--- aiguillage en cas d'erreur ---
    On Error Resume Next
    
    '--- aiguillage en fonction de la touche ---
    Select Case KeyCode
        Case vbKeyQ, vbKeyEscape: CBQuitter_Click
        Case vbKeyV: CBValider_Click
        Case Else
    End Select

End Sub

Private Sub PBBoutons_Resize()
    
    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next
    
    '--- calculs de l'emplacement des boutons ---
    CBQuitter.Left = PBBoutons.ScaleWidth - MARGES.M_BORD_DROIT - CBQuitter.Width
    CBValider.Left = CBQuitter.Left - MARGES.M_ENTRE_BOUTONS - CBValider.Width
    
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

Private Sub Timer_Animation_Timer()
    
    '--- aiguillage en cas d'erreur ---
    On Error Resume Next
    
    '--- déclaration ---
    Static NumImages As Byte
    
    '--- affichage de l'image ---
    PBLogoAnime(0).Picture = PCAnimation.GraphicCell(NumImages)
    
    '--- cotrôle du nombre d'images ---
    If NumImages = Pred(PCAnimation.Rows * PCAnimation.Cols) Then
         NumImages = 0
        CBQuitter.Enabled = True
    Else
        Inc NumImages
    End If
    
    '--- bip de passage dans la routine UNIQUEMENT POUR LES TESTS ---
    If PROGRAMME_AVEC_AUTOMATE = False Then Beep

End Sub

Private Sub TimerDateHeure_Timer()

    '--- aiguillage en cas d'erreur ---
    On Error Resume Next

    '--- déclaration ---
    Dim DateHeureActuel As String
    Static MemDateHeureActuel As String

    '--- affectation et affichage de la date et heure ---
    DateHeureActuel = StrConv(FormatDateTime(Date, vbLongDate), vbProperCase) & " - " & _
                                  FormatDateTime(Time, vbLongTime)
                                 
    '--- affichage ---
    If MemDateHeureActuel <> DateHeureActuel Then
        LDateHeureActuel = DateHeureActuel
        MemDateHeureActuel = DateHeureActuel
    End If
    
    '--- bip de passage dans la routine UNIQUEMENT POUR LES TESTS ---
    If PROGRAMME_AVEC_AUTOMATE = False Then Beep

End Sub

VERSION 5.00
Begin VB.Form FFinDeJournee 
   BackColor       =   &H00C0C0C0&
   ClientHeight    =   7560
   ClientLeft      =   825
   ClientTop       =   2745
   ClientWidth     =   11895
   Icon            =   "FFinDeJournee.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7560
   ScaleWidth      =   11895
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox PBRenseignementsFenetre 
      Align           =   1  'Align Top
      BackColor       =   &H00FF0000&
      Height          =   375
      Left            =   0
      Picture         =   "FFinDeJournee.frx":0442
      ScaleHeight     =   315
      ScaleWidth      =   11835
      TabIndex        =   9
      Top             =   0
      Width           =   11895
      Begin VB.Label LRenseignementsFenetre 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00008000&
         BackStyle       =   0  'Transparent
         Caption         =   "FIN DE JOURNEE"
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
         Left            =   120
         TabIndex        =   10
         Top             =   0
         Width           =   11595
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
      ScaleWidth      =   11835
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   6465
      Width           =   11895
      Begin VB.CommandButton CBFinDeJournee 
         Caption         =   "-"
         DownPicture     =   "FFinDeJournee.frx":24D84
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   675
         Left            =   7080
         MaskColor       =   &H00FF00FF&
         Picture         =   "FFinDeJournee.frx":269E6
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   60
         UseMaskColor    =   -1  'True
         Width           =   2895
      End
      Begin VB.PictureBox PBOutilsDeplacementFenetre 
         BackColor       =   &H00E0E0E0&
         Height          =   795
         Left            =   0
         ScaleHeight     =   735
         ScaleWidth      =   990
         TabIndex        =   5
         Top             =   0
         Visible         =   0   'False
         Width           =   1050
         Begin VB.CommandButton CBAgrandirFenetre 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Agrandir"
            DownPicture     =   "FFinDeJournee.frx":28648
            Height          =   480
            Left            =   0
            MaskColor       =   &H00FF00FF&
            Picture         =   "FFinDeJournee.frx":287F2
            Style           =   1  'Graphical
            TabIndex        =   8
            ToolTipText     =   " Agrandissement de la fenêtre "
            Top             =   0
            UseMaskColor    =   -1  'True
            Width           =   735
         End
         Begin VB.VScrollBar VSDeplacementFenetre 
            Height          =   735
            LargeChange     =   300
            Left            =   735
            SmallChange     =   100
            TabIndex        =   7
            Top             =   0
            Width           =   255
         End
         Begin VB.HScrollBar HSDeplacementFenetre 
            Height          =   255
            LargeChange     =   300
            Left            =   0
            SmallChange     =   100
            TabIndex        =   6
            Top             =   480
            Width           =   735
         End
      End
      Begin VB.Timer TimerEtatsFinDeJournee 
         Enabled         =   0   'False
         Interval        =   250
         Left            =   6180
         Top             =   180
      End
      Begin VB.CommandButton CBQuitter 
         Cancel          =   -1  'True
         Caption         =   "&Quitter"
         DownPicture     =   "FFinDeJournee.frx":2899C
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   675
         Left            =   10980
         MaskColor       =   &H00FF00FF&
         Picture         =   "FFinDeJournee.frx":2909E
         Style           =   1  'Graphical
         TabIndex        =   1
         ToolTipText     =   " Quitter cette fenêtre "
         Top             =   60
         UseMaskColor    =   -1  'True
         Width           =   795
      End
      Begin VB.CommandButton CBValider 
         Caption         =   "&Valider"
         DownPicture     =   "FFinDeJournee.frx":297A0
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   675
         Left            =   10080
         MaskColor       =   &H00FF00FF&
         Picture         =   "FFinDeJournee.frx":29EA2
         Style           =   1  'Graphical
         TabIndex        =   0
         ToolTipText     =   " Valider l'enregistrement "
         Top             =   60
         UseMaskColor    =   -1  'True
         Width           =   795
      End
   End
   Begin VB.PictureBox PBDeplacementFenetre 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   6675
      Index           =   0
      Left            =   0
      ScaleHeight     =   6675
      ScaleWidth      =   11895
      TabIndex        =   3
      Top             =   375
      Width           =   11895
      Begin VB.PictureBox PBDeplacementFenetre 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   6675
         Index           =   1
         Left            =   0
         ScaleHeight     =   6675
         ScaleWidth      =   11895
         TabIndex        =   4
         Top             =   0
         Width           =   11895
      End
   End
End
Attribute VB_Name = "FFinDeJournee"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Rôle                    : Fenêtre gérant la fin de journée
' Nom                    : FFinDeJournee.frm
' Date de création : 26/09/2002
' Détails                :
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

'--- déclarations obligatoires ---
Option Explicit

'--- options générales ---
Option Base 1
DefVar A-Z

'--- constantes privées ---
Private Const TITRE_FENETRE As String = "FIN DE JOURNEE"
Private Const TITRE_MESSAGES As String = TITRE_FENETRE

'--- énumérations privées ---

'--- variables privées ---
Private PremiereActivation As Boolean
Private InterdireEvenements As Boolean       'pour interdire certains évènements

'--- tableaux privées ---

'--- variables publiques ---
Public NumFenetre As Long                             'numéro de la fenêtre lorsqu'elle devient active

Private Sub CBAgrandirFENETRE_Click()
    On Error Resume Next
    Me.WindowState = vbMaximized
End Sub

Private Sub CBFinDeJournee_Click()
    
    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next

    '--- confirmation ---
    If GestionFinDeJourneeEnCours = False Then
    
        If AppelFenetre(F_MESSAGE, _
                                TITRE_MESSAGES, _
                                vbCrLf & vbCrLf & vbCrLf & _
                                "cs|Voulez-vous réellement LANCER la fin de journée ?", _
                                1, 0, 1) = vbYes Then
            
            '--- affectation ---
            GestionFinDeJourneeEnCours = True
    
        End If
    
    Else
        
        If AppelFenetre(F_MESSAGE, _
                                TITRE_MESSAGES, _
                                vbCrLf & vbCrLf & vbCrLf & _
                                "cs|Voulez-vous réellement ANNULER la fin de journée ?", _
                                1, 0, 1) = vbYes Then
    
            '--- affectation ---
            GestionFinDeJourneeEnCours = False
    
        End If
    
    End If

End Sub

Private Sub CBQuitter_Click()
    
    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next
    
    '--- demande de confirmation ---
    If CBValider.Enabled = True Then
        If AppelFenetre(F_MESSAGE, _
                               TITRE_MESSAGES, _
                               MESSAGE_1, _
                               TYPES_MESSAGES.T_AVERTISSEMENT, _
                               TYPES_BOUTONS.T_OUI_NON, _
                               EMPLACEMENT_FOCUS.E_SUR_OUI) = vbYes Then
            CBValider_Click
        End If
    End If
        
    '--- déchargement de la fenêtre ---
    DechargeFenetre
    
End Sub

Private Sub CBValider_Click()
    
    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next
    
    '--- affectation ---
    CBQuitter.Enabled = False
    
    '--- curseur de la souris ---
    SourisEnAttente True

    '--- curseur de la souris ---
    SourisEnAttente False
    
    '--- affectation ---
    CBQuitter.Enabled = True
    
End Sub

Private Sub Form_Activate()
    
    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next
    
    '--- renseigne la fenêtre principale ---
    RenseigneFPrincipale
    
    '--- placement du focus ---
    If PremiereActivation = False Then
        PremiereActivation = True
        Me.Refresh
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

Private Sub Form_Resize()
    
    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next
    
    '--- déclaration ---
    
    '--- zone mére et fille du déplacement de la fenetre ---
    PBDeplacementFenetre(ZONES_DEPLACEMENT_FENETRE.Z_MERE).Height = Abs(Me.ScaleHeight - PBBoutons.Height)
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
    CBValider.Left = CBQuitter.Left - MARGES.M_ENTRE_BOUTONS - CBValider.Width

End Sub

'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Rôle      : Restitue les dernières manipulations et valeurs sur la fenêtre
' Entrées :
' Retours :
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Private Sub LectureValeursFenetre()
    
    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next

    '--- déclaration ---
    Dim a As Integer

    '--- interdire certains évènements ---
    InterdireEvenements = True

    '--- affactation ---
    CBValider.Enabled = False
    
    '--- autoriser les évènements ---
    InterdireEvenements = False

End Sub

'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Rôle      : Enregistre les dernières manipulations et valeurs de la fenêtre
' Entrées :
' Retours :
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Private Sub EnregistreValeursfenetre()
    
    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next

    '--- déclaration ---
    Dim a As Integer

End Sub

'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Rôle      : Visualisation des différents états de la fin de journée
' Entrées :
' Retours :
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Private Sub EtatsFinDeJournee()
    
    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next

    '--- constantee privées ---
    
    '--- déclaration ---
    Dim a As Integer
    Dim ClignotantDefaut As Integer
    Static CptAppels As Integer
    Dim CouleurFond As Long, CouleurPlan As Long
    Dim Texte As String

    '--- contrôle du clignotement ---
    CptAppels = IIf(CptAppels > 11, 1, CptAppels + 1)
    ClignotantDefaut = Choose(CptAppels, 1, 1, 0, 0, 1, 1, 0, 0, 1, 1, 0, 0)
    
    
    '*************************************************************************************************
    '                                                       FIN DE JOURNEE
    '*************************************************************************************************
    'With OccFSynoptique.CBFinDeJournee
    '    If GestionFinDeJourneeEnCours = False Then
    '        If .BackColor <> COULEURS.CYAN_0 Then
    '            .Caption = "LANCER la fin de journée"
    '            .BackColor = COULEURS.CYAN_0
     '           .Refresh
     '       End If
     '   Else
     '       If .BackColor <> COULEURS.ROUGE_0 Then
     '           .Caption = "ANNULER la fin de journée"
     '           .BackColor = COULEURS.ROUGE_0
     '           .Refresh
     '       End If
     '   End If
    'End With
    
 
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
' Rôle      : Initialise la fenêtre (chargement ou en vue de la rendre visible)
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Sub InitialisationFenetre()
    
    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next

    '--- constantes privées ---
    
    '--- déclaration ---
    Dim a As Integer

    '--- affectation ---

    '--- divers sur la fenêtre ---
    PBDeplacementFenetre(ZONES_DEPLACEMENT_FENETRE.Z_FILLE).Picture = ImgFondDeFenetre
    With Me
        .Caption = TITRE_FENETRE
        .WindowState = vbMaximized
    End With
    
    '--- valeurs de la fenêtre ---
    LectureValeursFenetre
                
    '--- visualisation des différents états de la fin de journée ---
    EtatsFinDeJournee

    '--- lancement du timer ---
    TimerEtatsFinDeJournee.Enabled = True

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
    With TimerEtatsFinDeJournee
        .Enabled = False
        .Interval = 0
    End With
    
    '--- déchargement de la fenêtre ---
    Me.Visible = False
    Unload Me
    Set OccFFinDeJournee = Nothing
    
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

Private Sub TimerEtatsFinDeJournee_Timer()
    
    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next
    
    '--- appel de la routine ---
    TimerEtatsFinDeJournee.Enabled = False
    EtatsFinDeJournee
    TimerEtatsFinDeJournee.Enabled = True
    
    '--- bip de passage dans la routine UNIQUEMENT POUR LES TESTS ---
    If PROGRAMME_AVEC_AUTOMATE = False Then Beep

End Sub

Private Sub VSDeplacementFENETRE_Change()
    On Error Resume Next
    PBDeplacementFenetre(ZONES_DEPLACEMENT_FENETRE.Z_FILLE).Top = -VSDeplacementFenetre.Value
End Sub


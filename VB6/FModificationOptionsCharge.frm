VERSION 5.00
Begin VB.Form FModificationOptionsCharge 
   BackColor       =   &H00E0E0E0&
   Caption         =   "MODIFICATION DES OPTIONS D'UNE CHARGE"
   ClientHeight    =   6405
   ClientLeft      =   1455
   ClientTop       =   2655
   ClientWidth     =   7005
   Icon            =   "FModificationOptionsCharge.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6405
   ScaleWidth      =   7005
   Begin VB.PictureBox PBOptions 
      Height          =   4335
      Left            =   240
      ScaleHeight     =   4275
      ScaleWidth      =   6435
      TabIndex        =   6
      Top             =   660
      Width           =   6495
      Begin VB.CheckBox CBOptionsPostes 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Activer l'air dans le bain de BRILLANTAGE"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   420
         TabIndex        =   14
         Top             =   3660
         Width           =   5475
      End
      Begin VB.TextBox TBDelaiSupStabilisationCharge 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   3600
         MaxLength       =   2
         TabIndex        =   11
         Top             =   2640
         Width           =   435
      End
      Begin VB.CheckBox CBOptionsPonts 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Forcer la DESCENTE en PETITE VITESSE"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   3
         Left            =   420
         TabIndex        =   10
         Top             =   1800
         Width           =   5475
      End
      Begin VB.CheckBox CBOptionsPonts 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Forcer la MONTEE en TRES PETITE VITESSE"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   0
         Left            =   420
         TabIndex        =   9
         Top             =   360
         Width           =   5475
      End
      Begin VB.CheckBox CBOptionsPonts 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Forcer la MONTEE en PETITE VITESSE"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   1
         Left            =   420
         TabIndex        =   8
         Top             =   720
         Width           =   5475
      End
      Begin VB.CheckBox CBOptionsPonts 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Forcer la DESCENTE en TRES PETITE VITESSE"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   2
         Left            =   420
         TabIndex        =   7
         Top             =   1440
         Width           =   5475
      End
      Begin VB.Shape SDecorationActiverAirBainBrillantage 
         BackColor       =   &H00C0FFC0&
         BackStyle       =   1  'Opaque
         Height          =   615
         Left            =   180
         Shape           =   4  'Rounded Rectangle
         Top             =   3480
         Width           =   6075
      End
      Begin VB.Label LLibelles 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Délai supplémentaire de stabilisation de la charge en ARRET au POSTE"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Index           =   13
         Left            =   360
         TabIndex        =   13
         Top             =   2460
         Width           =   3135
      End
      Begin VB.Label LLibelles 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Secondes"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   11
         Left            =   4140
         TabIndex        =   12
         Top             =   2700
         Width           =   1215
      End
      Begin VB.Shape SDecoration 
         BackColor       =   &H00C0FFC0&
         BackStyle       =   1  'Opaque
         Height          =   915
         Index           =   6
         Left            =   180
         Shape           =   4  'Rounded Rectangle
         Top             =   180
         Width           =   6075
      End
      Begin VB.Shape SDecoration 
         BackColor       =   &H00C0FFC0&
         BackStyle       =   1  'Opaque
         Height          =   915
         Index           =   0
         Left            =   180
         Shape           =   4  'Rounded Rectangle
         Top             =   1260
         Width           =   6075
      End
      Begin VB.Shape SDecoration 
         BackColor       =   &H00C0FFC0&
         BackStyle       =   1  'Opaque
         Height          =   975
         Index           =   8
         Left            =   180
         Shape           =   4  'Rounded Rectangle
         Top             =   2340
         Width           =   6075
      End
   End
   Begin VB.PictureBox PBRenseignementsFenetre 
      Align           =   1  'Align Top
      BackColor       =   &H00FF0000&
      Height          =   375
      Left            =   0
      Picture         =   "FModificationOptionsCharge.frx":014A
      ScaleHeight     =   315
      ScaleWidth      =   6945
      TabIndex        =   1
      Top             =   0
      Width           =   7005
      Begin VB.Label LRenseignementsFenetre 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00008000&
         BackStyle       =   0  'Transparent
         Caption         =   "CHARGE GEREE"
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
         Left            =   60
         TabIndex        =   2
         Top             =   0
         Width           =   5295
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
      ScaleWidth      =   6945
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   5310
      Width           =   7005
      Begin VB.CommandButton CBAnnuler 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Annuler"
         DownPicture     =   "FModificationOptionsCharge.frx":24A8C
         Enabled         =   0   'False
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
         Left            =   720
         MaskColor       =   &H00FF00FF&
         Picture         =   "FModificationOptionsCharge.frx":2518E
         Style           =   1  'Graphical
         TabIndex        =   5
         ToolTipText     =   " Annuler les dernières modifications "
         Top             =   120
         UseMaskColor    =   -1  'True
         Width           =   1515
      End
      Begin VB.CommandButton CBValider 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Valider"
         DownPicture     =   "FModificationOptionsCharge.frx":25890
         Enabled         =   0   'False
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
         Left            =   2340
         MaskColor       =   &H00FF00FF&
         Picture         =   "FModificationOptionsCharge.frx":25F92
         Style           =   1  'Graphical
         TabIndex        =   4
         ToolTipText     =   " Valider l'enregistrement "
         Top             =   120
         UseMaskColor    =   -1  'True
         Width           =   1515
      End
      Begin VB.CommandButton CBQuitter 
         BackColor       =   &H00FFFFFF&
         Cancel          =   -1  'True
         Caption         =   "&Quitter"
         DownPicture     =   "FModificationOptionsCharge.frx":26694
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
         Left            =   3960
         MaskColor       =   &H00FF00FF&
         Picture         =   "FModificationOptionsCharge.frx":26D96
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   " Quitter cette fenêtre "
         Top             =   120
         UseMaskColor    =   -1  'True
         Width           =   1515
      End
      Begin VB.Shape SFocus 
         BorderColor     =   &H000000FF&
         BorderWidth     =   5
         Height          =   315
         Left            =   120
         Top             =   120
         Visible         =   0   'False
         Width           =   360
      End
   End
End
Attribute VB_Name = "FModificationOptionsCharge"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Rôle                    : Fenêtre gérant la modification des options d'une charge
' Nom                    : FModificationOptionsCharge.frm
' Date de création : 03/02/2010
' Détails                :
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

'--- déclarations obligatoires ---
Option Explicit

'--- options générales ---
Option Base 1
DefVar A-Z
    
'--- constantes privées ---
Private Const TITRE_FENETRE As String = "MODIFICATION DES OPTIONS D'UNE CHARGE"
Private Const TITRE_MESSAGES As String = TITRE_FENETRE

'--- énumérations privées ---

'--- types privées ---
Private PremiereActivation As Boolean
Private InterdireEvenements As Boolean      'pour interdire certains évènements

'--- variables privées ---
Private NumChargeEnCours As Integer          'numéro de la charge en cours
Private MemDernierBouton As Long               'mémoire du dernier bouton

'--- variables publiques ---
Public NumFenetre As Long                            'numéro de la fenêtre lorsqu'elle devient active

Private Sub CBAnnuler_Click()
    
    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next
    
    '--- lecture des options de la charge ---
    LectureOptionsCharge

End Sub

Private Sub CBAnnuler_GotFocus()
    
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

Private Sub CBAnnuler_LostFocus()
    On Error Resume Next
    SFocus.Visible = False
End Sub

Private Sub CBOptionsPonts_Click(Index As Integer)
    
    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next

    '--- analyse du changement d'état ---
    If InterdireEvenements = False Then
        CBValider.Enabled = True
        CBAnnuler.Enabled = True
    End If

    '--- céchochage des croix inutiles ---
    Select Case Index
        
        Case OPTIONS_GAMME.O_FORCER_MONTEE_EN_TPV
            '--- forcer la montée d'une charge en très petite vitesse ---
            If CBOptionsPonts(Index).Value = vbChecked Then
                CBOptionsPonts(OPTIONS_GAMME.O_FORCER_MONTEE_EN_PV).Value = vbUnchecked
            End If
        
        Case OPTIONS_GAMME.O_FORCER_MONTEE_EN_PV
            '--- forcer la montée d'une charge en petite vitesse ---
            If CBOptionsPonts(Index).Value = vbChecked Then
                CBOptionsPonts(OPTIONS_GAMME.O_FORCER_MONTEE_EN_TPV).Value = vbUnchecked
            End If
        
        Case OPTIONS_GAMME.O_FORCER_DESCENTE_EN_TPV
            '--- forcer la descente d'une charge en très petite vitesse ---
            If CBOptionsPonts(Index).Value = vbChecked Then
                CBOptionsPonts(OPTIONS_GAMME.O_FORCER_DESCENTE_EN_PV).Value = vbUnchecked
            End If
        
        Case OPTIONS_GAMME.O_FORCER_DESCENTE_EN_PV
            '--- forcer la descente d'une charge en petite vitesse ---
            If CBOptionsPonts(Index).Value = vbChecked Then
                CBOptionsPonts(OPTIONS_GAMME.O_FORCER_DESCENTE_EN_TPV).Value = vbUnchecked
            End If
        
        Case Else
    End Select

End Sub

Private Sub CBOptionsPostes_Click(Index As Integer)
    
    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next

    '--- analyse du changement d'état ---
    If InterdireEvenements = False Then
        CBValider.Enabled = True
        CBAnnuler.Enabled = True
    End If

End Sub

Private Sub CBQuitter_Click()
    On Error Resume Next
    If CBValider.Enabled = True Then
        Select Case AppelFenetre(F_MESSAGE, _
                                                 TITRE_MESSAGES, _
                                                 MESSAGE_1, _
                                                 TYPES_MESSAGES.T_AVERTISSEMENT, _
                                                 TYPES_BOUTONS.T_OUI_NON_ANNULER, _
                                                 EMPLACEMENT_FOCUS.E_SUR_OUI)
            Case vbYes
                CBValider_Click
                DechargeFenetre
            Case vbNo
                CBAnnuler_Click
                DechargeFenetre
            Case Else
        End Select
    Else
        DechargeFenetre
    End If
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
    
    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next
    
    '--- affectation ---
    CBQuitter.Enabled = False
    
    '--- curseur de la souris ---
    SourisEnAttente True
    
    '--- enregistrement des options de la charge ---
    EnregistreOptionsCharge

    '--- ne plus permettre la validation ---
    CBValider.Enabled = False
    CBAnnuler.Enabled = False
    
    '--- curseur de la souris ---
    SourisEnAttente False

    '--- affectation ---
    CBQuitter.Enabled = True

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

Private Sub PBBoutons_Resize()

    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next
    
    '--- calculs de l'emplacement des boutons ---
    CBQuitter.Left = PBBoutons.ScaleWidth - MARGES.M_BORD_DROIT - CBQuitter.Width
    CBValider.Left = CBQuitter.Left - MARGES.M_ENTRE_BOUTONS - CBValider.Width
    CBAnnuler.Left = CBValider.Left - MARGES.M_ENTRE_BOUTONS - CBAnnuler.Width

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
' Rôle      : Décharge la fenêtre
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Private Sub DechargeFenetre()
    
    '--- aiguillage en cas d'erreur ---
    On Error Resume Next
    
    '--- affectation ---
    PremiereActivation = False
    
    '--- curseur souris par défaut ---
    SourisEnAttente False

    '--- déchargement de la fenêtre ---
    Me.Visible = False
    Unload Me
    Set FModificationOptionsCharge = Nothing
    
End Sub

'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Rôle      : Effectue le paramètrage de la fenêtre
' Entrées : NumCharge -> Numéro de charge souhaité
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Sub ParametrageFenetre(ByVal NumCharge As Integer)
    
    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next

    '--- numéro de charge en cours ---
    NumChargeEnCours = NumCharge
    
    '--- lecture des options de la charge ---
    LectureOptionsCharge
    
End Sub

'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Rôle      : Change le curseur de la souris en fonction de l'attente
' Entrées : AttenteOuiNon -> TRUE   = Curseur en forme de sablier
'                                             FALSE = Curseur par défaut
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Private Sub SourisEnAttente(ByVal AttenteOuiNon As Boolean)
    
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
' Rôle      : Initialise la fenêtre (chargement ou en vue de la rendre visible)
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Sub InitialisationFenetre()
    
    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next

    '--- déclaration ---

    '--- affectation ---
  
    '--- divers sur la fenêtre ---
    Me.Caption = TITRE_FENETRE
    Centrefenetre Me
    
    '--- images des fonds ---
    Me.Picture = ImgFondBleu1
    PBOptions.Picture = ImgFondVert1
    PBBoutons.Picture = ImgFondDesBoutons

End Sub

'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Rôle      : Enregistre les options de la charge (transfert également le mot des options dans l'automate)
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Private Sub EnregistreOptionsCharge()
    
    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next
    
    '--- déclaration ---
    
    '--- enregistrement des valeurs de la charge ---
    If NumChargeEnCours >= CHARGES.C_NUM_MINI And NumChargeEnCours <= CHARGES.C_NUM_MAXI Then

        With TEtatsCharges(NumChargeEnCours)

            '--- transfert du temps en secondes du délai de stabilisation de la charge ---
            If IsNumeric(TBDelaiSupStabilisationCharge.Text) = True Then
                .DelaiSupStabilisationChargeSecondes = CInt(TBDelaiSupStabilisationCharge.Text)
            Else
                .DelaiSupStabilisationChargeSecondes = 0
            End If

            '--- construction du mot des options 1 (mot transmis à l'automate) ---
            'Poids FORT du mot transmis
            '---------------------------------------------------------------------------------------
            '|  Bit 7 |  Bit 6 | Bit 5 | Bit 4 | Bit 3 | Bit 2 | Bit 1 | Bit 0 |
            '---------------------------------------------------------------------------------------
            '|  128   |   64   |   32  |   16   |    8   |    4    |    2   |     1   |
            '---------------------------------------------------------------------------------------
            '      |           |          |         |         |          |          |         |_____  forcer la montée en très petite vitesse
            '      |           |          |         |         |          |          |__________  forcer la montée en petite vitesse
            '      |           |          |         |         |          |________________ forcer la descente en très petite vitesse
            '      |           |          |         |         |_____________________  forcer la descente en petite vitesse
            '      |           |          |         |__________________________
            '      |           |          |_______________________________
            '      |           |_____________________________________
            '      |___________________________________________
                    
            '--- construction du mot des options 2 (mot transmis à l'automate) ---
            'Poids FORT du mot transmis
            '---------------------------------------------------------------------------------------
            '|  Bit 7 |  Bit 6 | Bit 5 | Bit 4 | Bit 3 | Bit 2 | Bit 1 | Bit 0 |
            '---------------------------------------------------------------------------------------
            '|  128   |   64   |   32  |   16   |    8   |    4    |    2   |     1   |
            '---------------------------------------------------------------------------------------
            '      |           |          |         |         |          |          |         |_____  gestion de l'électro-vanne du brillantage avec les gammes
            '      |           |          |         |         |          |          |__________
            '      |           |          |         |         |          |________________
            '      |           |          |         |         |_____________________
            '      |           |          |         |__________________________
            '      |           |          |_______________________________
            '      |           |_____________________________________
            '      |___________________________________________
            
            '--- initialisation du mot contenant les options 1 et 2 ---
            .Options1 = 0
            .Options2 = 0

            '--- options 1 ---
            If CBOptionsPonts(OPTIONS_GAMME.O_FORCER_DESCENTE_EN_PV).Value = 1 Then
                .Options1 = .Options1 + 8                         'bit 3 du mot des options 1
            End If
            If CBOptionsPonts(OPTIONS_GAMME.O_FORCER_DESCENTE_EN_TPV).Value = 1 Then
                .Options1 = .Options1 + 4                         'bit 2 du mot des options 1
            End If
            If CBOptionsPonts(OPTIONS_GAMME.O_FORCER_MONTEE_EN_PV).Value = 1 Then
                .Options1 = .Options1 + 2                         'bit 1 du mot des options 1
            End If
            If CBOptionsPonts(OPTIONS_GAMME.O_FORCER_MONTEE_EN_TPV).Value = 1 Then
                .Options1 = .Options1 + 1                         'bit 0 du mot des options 1
            End If
            
            '--- options 2 ---
            If CBOptionsPostes(OPTIONS_GAMME.O_ACTIVER_AIR_BRILLANTAGE).Value = 1 Then
                .Options2 = .Options2 + 1                         'bit 0 du mot des options 2
            End If
            
            '--- envoi dans l'automate du numéro de charge ---
            If PROGRAMME_AVEC_AUTOMATE = True Then
                EnvoiOptionsPourUneCharge NumCharge:=NumChargeEnCours, _
                                                                Options1:=.Options1, _
                                                                Options2:=.Options2
            End If

        End With

    End If

End Sub

'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Rôle      : Lecture des options de la charge
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Private Sub LectureOptionsCharge()
    
    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next

    '--- déclaration ---
    Dim Options1Binaire As String * 16
    Dim Options2Binaire As String * 16
    
    '--- interdire certains évènements ---
    InterdireEvenements = True

    '--- affichage des valeurs de la charge ---
    If NumChargeEnCours >= CHARGES.C_NUM_MINI And NumChargeEnCours <= CHARGES.C_NUM_MAXI Then
            
        '--- affichage du numéro de charge ---
        LRenseignementsFenetre.Caption = "Charge n° " & NumChargeEnCours
        
        With TEtatsCharges(NumChargeEnCours)
        
            '--- affichage du délai supplémentaire de stabilisation de la charge ---
            If .DelaiSupStabilisationChargeSecondes = 0 Then
                TBDelaiSupStabilisationCharge.Text = ""
            Else
                TBDelaiSupStabilisationCharge.Text = Format(.DelaiSupStabilisationChargeSecondes, FORMAT_DELAI_SUP_STABILISATION_CHARGE)
            End If
        
            '--- rendre visible ou non l'activation de l'air dans le bain de brillantage ---
            If PassageBrillantage(.TGammesAnodisation) = True Then
                SDecorationActiverAirBainBrillantage.Visible = True                                                                             'rendre visible l'activation de l'air dans le bain de brillantage
                CBOptionsPostes(OPTIONS_GAMME.O_ACTIVER_AIR_BRILLANTAGE).Visible = True
            Else
                SDecorationActiverAirBainBrillantage.Visible = False                                                                           'rendre invisible l'activation de l'air dans le bain de brillantage car il n'y a pas de brillantage dans la gamme
                CBOptionsPostes(OPTIONS_GAMME.O_ACTIVER_AIR_BRILLANTAGE).Visible = False
            End If

            '--- conversion en binaire ---
            Options1Binaire = CBin(.Options1)
            Options2Binaire = CBin(.Options2)

        End With
    
        '--- décodage du mot des options 1 ---
        '---------------------------------------------------------------------------------------
        '|  Bit 7 |  Bit 6 | Bit 5 | Bit 4 | Bit 3 | Bit 2 | Bit 1 | Bit 0 |
        '---------------------------------------------------------------------------------------
        '|  128   |   64   |   32  |   16   |    8   |    4    |    2   |     1   |
        '---------------------------------------------------------------------------------------
        '      |           |          |         |         |          |          |         |_____  forcer la montée en très petite vitesse
        '      |           |          |         |         |          |          |__________  forcer la montée en petite vitesse
        '      |           |          |         |         |          |________________ forcer la descente en très petite vitesse
        '      |           |          |         |         |_____________________  forcer la descente en petite vitesse
        '      |           |          |         |__________________________
        '      |           |          |_______________________________
        '      |           |_____________________________________
        '      |___________________________________________
                    
        '--- décodage du mot des options 2 ---
        '---------------------------------------------------------------------------------------
        '|  Bit 7 |  Bit 6 | Bit 5 | Bit 4 | Bit 3 | Bit 2 | Bit 1 | Bit 0 |
        '---------------------------------------------------------------------------------------
        '|  128   |   64   |   32  |   16   |    8   |    4    |    2   |     1   |
        '---------------------------------------------------------------------------------------
        '      |           |          |         |         |          |          |         |_____  activer l'air dans le brillantage
        '      |           |          |         |         |          |          |__________
        '      |           |          |         |         |          |________________
        '      |           |          |         |         |_____________________
        '      |           |          |         |__________________________
        '      |           |          |_______________________________
        '      |           |_____________________________________
        '      |___________________________________________
        
        '--- forcer la montée en très petite vitesse ---
        CBOptionsPonts(OPTIONS_GAMME.O_FORCER_MONTEE_EN_TPV).Value = Bit(Options1Binaire, OPTIONS_GAMME.O_FORCER_MONTEE_EN_TPV)
        
        '--- forcer la montée en petite vitesse ---
        CBOptionsPonts(OPTIONS_GAMME.O_FORCER_MONTEE_EN_PV).Value = Bit(Options1Binaire, OPTIONS_GAMME.O_FORCER_MONTEE_EN_PV)
        
        '--- forcer la descente en très petite vitesse ---
        CBOptionsPonts(OPTIONS_GAMME.O_FORCER_DESCENTE_EN_TPV).Value = Bit(Options1Binaire, OPTIONS_GAMME.O_FORCER_DESCENTE_EN_TPV)
        
        '--- forcer la descente en petite vitesse ---
        CBOptionsPonts(OPTIONS_GAMME.O_FORCER_DESCENTE_EN_PV).Value = Bit(Options1Binaire, OPTIONS_GAMME.O_FORCER_DESCENTE_EN_PV)
        
        '--- activer l'air dans le brillantage ---
        CBOptionsPostes(OPTIONS_GAMME.O_ACTIVER_AIR_BRILLANTAGE).Value = Bit(Options2Binaire, OPTIONS_GAMME.O_ACTIVER_AIR_BRILLANTAGE)
    
    Else
    
        '--- affichage indiquant qu'il n'y a pas de charge en cours (cas normallement impossible) ---
        LRenseignementsFenetre.Caption = "PAS DE CHARGE EN COURS"
    
    End If
    
    '--- affectation ---
    CBValider.Enabled = False
    CBAnnuler.Enabled = False
    
    '--- autoriser les évènements ---
    InterdireEvenements = False
 
End Sub

Private Sub TBDelaiSupStabilisationCharge_Change()
    
    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next

    '--- analyse du changement d'état ---
    If InterdireEvenements = False Then
        CBValider.Enabled = True
        CBAnnuler.Enabled = True
    End If

End Sub

Private Sub TBDelaiSupStabilisationCharge_GotFocus()
    
    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next
    
    With TBDelaiSupStabilisationCharge
        If IsNumeric(.Text) = True Then
            .Text = CStr(CLng(.Text))
        Else
            .Text = ""
        End If
        .SelStart = 0          'met en surbrillance la sélection saisie
        .SelLength = Len(.Text)
    End With

End Sub

Private Sub TBDelaiSupStabilisationCharge_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error Resume Next
    FiltreToucheFonction KeyCode, Shift
End Sub

Private Sub TBDelaiSupStabilisationCharge_KeyPress(KeyAscii As Integer)
    On Error Resume Next
    FiltreToucheASCII KeyAscii, DONNEES.D_NBR_NATURELS
End Sub

Private Sub TBDelaiSupStabilisationCharge_LostFocus()
    
     '--- aiguillage en cas d'erreurs ---
    On Error Resume Next
    
    With TBDelaiSupStabilisationCharge
        If IsNumeric(.Text) = True Then
            .Text = Format(CLng(.Text), FORMAT_DELAI_SUP_STABILISATION_CHARGE)
        Else
            .Text = ""
        End If
    End With

End Sub

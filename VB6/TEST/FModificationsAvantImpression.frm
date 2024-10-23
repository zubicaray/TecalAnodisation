VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form FModificationsAvantImpression 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   5700
   ClientLeft      =   1440
   ClientTop       =   2355
   ClientWidth     =   9555
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5700
   ScaleWidth      =   9555
   Begin VB.TextBox TBEditionModificationsAvantImpression 
      BackColor       =   &H00C0C0FF&
      BorderStyle     =   0  'None
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   270
      TabIndex        =   7
      Top             =   255
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.PictureBox PBBoutons 
      Align           =   2  'Align Bottom
      DrawStyle       =   2  'Dot
      DrawWidth       =   16891
      Height          =   855
      Left            =   0
      ScaleHeight     =   795
      ScaleWidth      =   9495
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   4845
      Width           =   9555
      Begin VB.CommandButton CBInsererSurGrille 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Insérer"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   250
         Left            =   60
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   510
         Width           =   975
      End
      Begin VB.CommandButton CBCompacterSurGrille 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Compacter"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   250
         Left            =   60
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   270
         Width           =   975
      End
      Begin VB.CommandButton CBSupprimerSurGrille 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Supprimer"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   250
         Left            =   60
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   30
         Width           =   975
      End
      Begin VB.CommandButton CBRetablir 
         Caption         =   "&Rétablir"
         DownPicture     =   "FModificationsAvantImpression.frx":0000
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
         Left            =   6840
         Picture         =   "FModificationsAvantImpression.frx":0702
         Style           =   1  'Graphical
         TabIndex        =   0
         ToolTipText     =   " Rétablir la grille d'origine "
         Top             =   60
         Width           =   795
      End
      Begin VB.CommandButton CBValider 
         Caption         =   "&Valider"
         DownPicture     =   "FModificationsAvantImpression.frx":0E04
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
         Left            =   7740
         Picture         =   "FModificationsAvantImpression.frx":1506
         Style           =   1  'Graphical
         TabIndex        =   1
         ToolTipText     =   " Valider l'enregistrement "
         Top             =   60
         Width           =   795
      End
      Begin VB.CommandButton CBQuitter 
         Cancel          =   -1  'True
         Caption         =   "&Quitter"
         DownPicture     =   "FModificationsAvantImpression.frx":1C08
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
         Left            =   8640
         Picture         =   "FModificationsAvantImpression.frx":230A
         Style           =   1  'Graphical
         TabIndex        =   2
         ToolTipText     =   " Quitter cette fenêtre "
         Top             =   60
         Width           =   795
      End
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFGModificationsAvantImpression 
      Height          =   4575
      Left            =   120
      TabIndex        =   8
      Top             =   120
      Width           =   9315
      _ExtentX        =   16431
      _ExtentY        =   8070
      _Version        =   393216
      BackColor       =   16777215
      ForeColor       =   -2147483646
      Rows            =   31
      Cols            =   6
      GridColor       =   12632256
      GridColorUnpopulated=   -2147483644
      AllowBigSelection=   0   'False
      FocusRect       =   2
      HighLight       =   0
      ScrollBars      =   2
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontFixed {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _NumberOfBands  =   1
      _Band(0).Cols   =   6
      _Band(0).GridLinesBand=   1
      _Band(0).TextStyleBand=   0
      _Band(0).TextStyleHeader=   0
   End
   Begin VB.Shape SFocusTableModificationsAvantImpression 
      BorderColor     =   &H000000FF&
      BorderWidth     =   4
      Height          =   4590
      Left            =   120
      Top             =   120
      Visible         =   0   'False
      Width           =   9330
   End
End
Attribute VB_Name = "FModificationsAvantImpression"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Rôle                    : Feuille permettant la modification des données avant impression
' Nom                    : FModificationsAvantImpression.frm
' Date de création : 20/01/2000
' Détails                :
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

'--- déclarations obligatoires ---
Option Explicit

'--- options générales ---
Option Base 1
DefVar A-Z
    
'--- constantes privées ---

'--- énumérations privées ---
Private Enum COLONNES_MODIFICATIONS_AVANT_IMPRESSION
    C_NUM_LIGNES = 0
    C_COLONNE_1 = 1
    C_COLONNE_2 = 2
End Enum

'--- variables privées ---
Private PremiereActivation As Boolean
Private LigneDepartDeplacement As Integer    'ligne de départ en cas de déplacement d'un détail
Private LigneArriveeDeplacement As Integer   'ligne de d'arrivée en cas de déplacement d'un détail
Private NbrLignesModificationsAvantImpression As Integer      'nombre de lignes des modifications avant impression
Private NbrColonnesModificationsAvantImpression As Integer  'nombre de colonnes des modifications avant impression
Private MemDernierBouton As Long                 'mémoire du dernier bouton
Private NumFeuilleAppel As Long                     'numéro de feuille ayant lancé l'appel
Private CritereRecherche As String                   'critère de recherche
Private TitreFeuille As String                             'titre de la feuille
Private TitreMessages As String                       'titre des messages
Private RepereFeuilleCritereRecherche As String 'repère de la feuille d'appel et du critère de recherche

'--- tableaux privés ---

'--- variables publiques ---
Public NumFeuille As Long                                'numéro de la feuille lorsqu'elle devient active
Public VariableRetourneeFeuille As Variant      'variable retournée par la feuille (représente la date sélectionnée)

Private Sub CBCompacterSurGrille_Click()
    On Error Resume Next
    GestionModificationsAvantImpression GG_COMPRESSION_GRILLE
    GestionModificationsAvantImpression GG_CONSTRUCTION_GRILLE
End Sub

Private Sub CBInsererSurGrille_Click()
    
    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next

    '--- déclaration ---
    Dim a As Integer, _
           NumLigne As Integer
    Dim FicheVide As EnrModificationsAvantImpression
    
    '--- affectation ---
    NumLigne = MSHFGModificationsAvantImpression.Row
    
    '--- suppression de la ligne ---
    If NumLigne > 0 And NumLigne <= NbrLignesModificationsAvantImpression Then
        For a = Pred(NbrLignesModificationsAvantImpression) To NumLigne Step -1
            TModificationsAvantImpression(Succ(a)) = TModificationsAvantImpression(a)
        Next a
        TModificationsAvantImpression(NumLigne) = FicheVide
        GestionModificationsAvantImpression GG_CONSTRUCTION_GRILLE
        With MSHFGModificationsAvantImpression
            .Col = COLONNES_MODIFICATIONS_AVANT_IMPRESSION.C_COLONNE_1
            .SetFocus
        End With
    End If

End Sub

Private Sub CBQuitter_Click()
    On Error Resume Next
    VariableRetourneeFeuille = ETATS_BOUTONS.E_APRES_QUITTER
    Me.Hide
    DoEvents
End Sub

Private Sub CBRetablir_Click()
    
    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next
        
    '--- curseur de la souris ---
    SourisEnAttente True
    
    '--- gestion des boutons ---
    GestionBoutons E_AVANT_RETABLIR
    
    '--- annuler puis réaffichage ---
    GestionModificationsAvantImpression GG_VIDAGE_GRILLE
    GestionModificationsAvantImpression GG_TRANSFERT_DONNEES_VERS_GRILLE
    GestionModificationsAvantImpression GG_CONSTRUCTION_GRILLE
    
    '--- gestion des boutons ---
    GestionBoutons E_APRES_RETABLIR
    
    '--- curseur de la souris ---
    SourisEnAttente False

End Sub

Private Sub CBSupprimerSurGrille_Click()
    
    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next

    '--- déclaration ---
    Dim NumLigne As Integer
    Dim FicheVide As EnrModificationsAvantImpression
    
    '--- affectation ---
    NumLigne = MSHFGModificationsAvantImpression.Row
    
    '--- suppression de la ligne ---
    If NumLigne > 0 And NumLigne <= NbrLignesModificationsAvantImpression Then
        If MsgBox(MESSAGE_3 & CStr(NumLigne) & " ?", vbYesNo + vbExclamation + vbDefaultButton2, TitreMessages) = vbYes Then
            TModificationsAvantImpression(NumLigne) = FicheVide
            GestionModificationsAvantImpression GG_COMPRESSION_GRILLE
            GestionModificationsAvantImpression GG_CONSTRUCTION_GRILLE
            GestionBoutons E_MODIFICATION_EN_COURS
        End If
        MSHFGModificationsAvantImpression.SetFocus
    End If

End Sub

Private Sub CBValider_Click()
    On Error Resume Next
    VariableRetourneeFeuille = ETATS_BOUTONS.E_APRES_VALIDER
    Me.Hide
    DoEvents
End Sub

Private Sub Form_Activate()
    
    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next
    
    '--- renseigne la feuille principale ---
    RenseigneFPrincipale
    
    '--- placement du focus ---
    If PremiereActivation = False Then
        Me.Refresh
        PremiereActivation = True
    End If

End Sub

Private Sub Form_Load()
    
    '--- aiguillage en cas d'erreurs ---
    On Error GoTo GestionErreurs
    
    '--- déclaration ---
    
    '--- affectation ---
    PremiereActivation = False

    '--- divers sur la feuille ---
    Me.Picture = ImgFondDeFenetre
    CentreFeuille Me
    
    '--- gestion de l'états des boutons ---
    GestionBoutons E_CHARGEMENT_FEUILLE
    
    Exit Sub

'--- gestion des erreurs ---
GestionErreurs:
      
    '--- affichage du message d'erreur ---
    MessageErreur TitreMessages, Err.Description, Err.Number

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    On Error Resume Next
    Select Case UnloadMode
        Case vbFormCode
        Case Else: CBQuitter_Click
    End Select
End Sub

Private Sub MSHFGModificationsAvantImpression_DblClick()
    On Error Resume Next
    EditionModificationsAvantImpression vbKeySpace  'simule un espace
End Sub

Private Sub MSHFGModificationsAvantImpression_GotFocus()
    On Error Resume Next
    SFocusTableModificationsAvantImpression.Visible = True
End Sub

Private Sub MSHFGModificationsAvantImpression_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error Resume Next
    Select Case KeyCode
        Case vbKeyDelete: EditionModificationsAvantImpression vbKeyBack  'simule un retour arrière (effacement)
        Case Else
    End Select
End Sub

Private Sub MSHFGModificationsAvantImpression_KeyPress(KeyAscii As Integer)
    On Error Resume Next
    EditionModificationsAvantImpression KeyAscii 'envoi de la touche frappée
End Sub

Private Sub MSHFGModificationsAvantImpression_LeaveCell()
    On Error Resume Next
    TBEditionModificationsAvantImpression.Visible = False
End Sub

Private Sub MSHFGModificationsAvantImpression_LostFocus()
    On Error Resume Next
    SFocusTableModificationsAvantImpression.Visible = False
End Sub

Private Sub MSHFGModificationsAvantImpression_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next

    '--- mémorisation de la ligne de départ ---
    With MSHFGModificationsAvantImpression
        If Button = vbKeyLButton And .MouseCol = COLONNES_MODIFICATIONS_AVANT_IMPRESSION.C_NUM_LIGNES Then
            LigneDepartDeplacement = .MouseRow
        End If
    End With

End Sub

Private Sub MSHFGModificationsAvantImpression_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next
        
    '--- déclaration ---
    
    '--- RAZ des variables de déplacement ---
    If Button <> vbKeyLButton Then
        LigneDepartDeplacement = 0
        LigneArriveeDeplacement = 0
    End If

End Sub

Private Sub MSHFGModificationsAvantImpression_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next
        
    '--- mémorisation de la ligne d'arrivée ---
    With MSHFGModificationsAvantImpression
        If Button = vbKeyLButton And .MouseCol = COLONNES_MODIFICATIONS_AVANT_IMPRESSION.C_NUM_LIGNES Then
            LigneArriveeDeplacement = .MouseRow
            If LigneDepartDeplacement > 0 And _
               LigneArriveeDeplacement > 0 And _
               LigneDepartDeplacement <> LigneArriveeDeplacement Then
                    DeplacementLigne
            End If
        End If
    End With

End Sub

Private Sub MSHFGModificationsAvantImpression_Scroll()
    On Error Resume Next
    MSHFGModificationsAvantImpression.SetFocus
End Sub

Private Sub PBBoutons_Resize()

    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next
    
    '--- calculs de l'emplacement des boutons ---
    CBQuitter.Left = PBBoutons.ScaleWidth - MARGES.M_BORD_DROIT - CBQuitter.Width
    CBValider.Left = CBQuitter.Left - MARGES.M_ENTRE_BOUTONS - CBValider.Width
    CBRetablir.Left = CBValider.Left - MARGES.M_ENTRE_BOUTONS - CBRetablir.Width

End Sub

'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Rôle      : Gére l'états des boutons après une action de l'opèrateur
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Private Sub GestionBoutons(ByVal Situation As ETATS_BOUTONS)
    
    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next

    Select Case Situation
        
        Case ETATS_BOUTONS.E_CHARGEMENT_FEUILLE
            '--- au chargement de la feuille ---
            CBQuitter.Enabled = True
            CBValider.Enabled = True
            CBRetablir.Enabled = True
        
        Case ETATS_BOUTONS.E_DECHARGEMENT_FEUILLE
            '--- au déchargement de la feuille ---
        
        Case ETATS_BOUTONS.E_AVANT_VALIDER
            '--- avant valider ---
        
        Case ETATS_BOUTONS.E_APRES_VALIDER
            '--- après valider ---
            CBQuitter.Enabled = True
            CBValider.Enabled = True
            CBRetablir.Enabled = True

        Case ETATS_BOUTONS.E_AVANT_RETABLIR
            '--- avant annuler ---
        
        Case ETATS_BOUTONS.E_APRES_RETABLIR
            '--- après annuler ---
            CBQuitter.Enabled = True
            CBValider.Enabled = True
            CBRetablir.Enabled = True

        Case ETATS_BOUTONS.E_MODIFICATION_EN_COURS
            '--- après modifier (à ne pas traiter si nouvel enregistrement) ---
            If MemDernierBouton = ETATS_BOUTONS.E_APRES_NOUVEAU Then Exit Sub
            CBQuitter.Enabled = True
            CBValider.Enabled = True
            CBRetablir.Enabled = True

        Case Else
    
    End Select

    '--- affectation ---
    MemDernierBouton = Situation

End Sub

Private Sub TBeditionmodificationsavantimpression_Change()
    
    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next
    
    '--- déclaration ---
    Dim TexteATraiter As String
    
    If PremiereActivation = True Then
            
        If Me.ActiveControl.Name = TBEditionModificationsAvantImpression.Name Then
                
            '--- affectation ---
            TexteATraiter = TBEditionModificationsAvantImpression.Text
                
            '--- indiquer la modification en cours ---
            GestionBoutons E_MODIFICATION_EN_COURS
            
            With MSHFGModificationsAvantImpression
            
                Select Case .Col
                    
                    Case COLONNES_MODIFICATIONS_AVANT_IMPRESSION.C_COLONNE_1
                        '--- colonne 1 ---
                        TModificationsAvantImpression(.Row).Colonne1 = TexteATraiter
                    
                    Case COLONNES_MODIFICATIONS_AVANT_IMPRESSION.C_COLONNE_2
                        '--- colonne 2 ---
                        TModificationsAvantImpression(.Row).Colonne2 = TexteATraiter

                    Case Else
            
                End Select
    
            End With
    
        End If

    End If

End Sub

Private Sub TBeditionmodificationsavantimpression_GotFocus()
    On Error Resume Next
    SFocusTableModificationsAvantImpression.Visible = True
End Sub

Private Sub TBeditionmodificationsavantimpression_KeyDown(KeyCode As Integer, Shift As Integer)
    
    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next
    
    With MSHFGModificationsAvantImpression
    
        '--- analyse en fonction de la touche ---
        Select Case KeyCode
        
            Case vbKeyInsert, vbKeyF12
                FiltreToucheFonction KeyCode, Shift
        
            Case vbKeyDown
                '--- flèche basse ---
                .SetFocus
                DoEvents
                If .Row < .Rows - 1 Then
                    .Row = .Row + 1
                End If
                KeyCode = 0
            
            Case vbKeyUp
                '--- flèche haute ---
                .SetFocus
                DoEvents
                If .Row > .FixedRows Then
                    .Row = .Row - 1
                End If
                KeyCode = 0
        
            Case Else

        End Select
    
    End With

End Sub

Private Sub TBeditionmodificationsavantimpression_KeyPress(KeyAscii As Integer)
    
    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next

    '--- déclaration ---
    Dim TexteATraiter As String
    
    With MSHFGModificationsAvantImpression
    
        '--- analyse de la touche ---
        Select Case KeyAscii
        
            Case vbKeyReturn
                '--- touche entrée ---
                TexteATraiter = TBEditionModificationsAvantImpression.Text

'                Select Case .Col
'
'                    Case COLONNES_MODIFICATIONS_AVANT_IMPRESSION.C_COLONNE_1
'                        '--- colonne 1 ---
'                        TModificationsAvantImpression(.Row).Colonne1 = TexteATraiter
'                        Select Case NumFeuilleAppel
'                            Case FEUILLES.F_DEVIS
'                                .Col = COLONNES_MODIFICATIONS_AVANT_IMPRESSION.C_COLONNE_2
'                            Case FEUILLES.F_FICHES_ATELIER
'                                If .Row < .Rows - 1 Then .Row = .Row + 1
'                            Case Else
'                        End Select
'
'                    Case COLONNES_MODIFICATIONS_AVANT_IMPRESSION.C_COLONNE_2
'                        '--- colonne 2 ---
'                        TModificationsAvantImpression(.Row).Colonne2 = TexteATraiter
'                        Select Case NumFeuilleAppel
'                            Case FEUILLES.F_DEVIS
'                                .Col = COLONNES_MODIFICATIONS_AVANT_IMPRESSION.C_COLONNE_1
'                            Case FEUILLES.F_FICHES_ATELIER
'                            Case Else
'                        End Select
'
'                    Case Else
'
'                End Select
            
                '--- mettre le focus sur le tableau ---
                .SetFocus
                DoEvents
                KeyAscii = 0
        
            Case Else
                Select Case .Col
                    Case COLONNES_MODIFICATIONS_AVANT_IMPRESSION.C_COLONNE_1: FiltreToucheASCII KeyAscii, DONNEES.D_GENERALE, 0
                    Case COLONNES_MODIFICATIONS_AVANT_IMPRESSION.C_COLONNE_2: FiltreToucheASCII KeyAscii, DONNEES.D_GENERALE, 0
                    Case Else
                End Select
    
        End Select

    End With

End Sub

Private Sub TBEditionModificationsAvantImpression_LostFocus()
    
    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next
    
    '--- focus ---
    SFocusTableModificationsAvantImpression.Visible = False
    
    '--- rendre le contrôle texte invisible ---
    TBEditionModificationsAvantImpression.Visible = False

    '--- construction de la grille / calculs ---
    GestionModificationsAvantImpression GG_CONSTRUCTION_GRILLE
    
End Sub

'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Rôle      : Gestion des modifications avant impression
' Entrées : EtatSouhaite -> Fonction de l'énumération GESTION_GRILLES
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Private Sub GestionModificationsAvantImpression(ByVal EtatSouhaite As GESTION_GRILLES)
    
    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next
    
    '--- constantes privées ---
    Const EPAISSEUR_CARACTERE As Integer = 140

    '--- déclaration ---
    Dim a As Integer, _
            b As Integer, _
            MemLigne As Integer, _
            MemColonne As Integer, _
            PtrLigne As Integer, _
            NumArticle As Integer
    Dim NumDevis As String, _
            NumFicheAtelier As String
    Dim FicheVide As EnrModificationsAvantImpression

    '--- déclaration de certaine variable ---
    ReDim TCopieModificationsAvantImpression(1 To NbrLignesModificationsAvantImpression) As EnrModificationsAvantImpression
    
    Select Case EtatSouhaite
    
        Case GESTION_GRILLES.GG_INITIALISATION_GRILLE
            '--- initialisation du tableau ---
            If MemRepereFeuilleCritereRecherche <> RepereFeuilleCritereRecherche Then
                ReDim TModificationsAvantImpression(1 To NbrLignesModificationsAvantImpression) As EnrModificationsAvantImpression
            End If
            
            '--- initialisation de la grille des détails ---
            With MSHFGModificationsAvantImpression
                
                .Redraw = False
                
                .Clear
                
                .FixedCols = 1
                .FixedRows = 1
                .Rows = NbrLignesModificationsAvantImpression + .FixedRows
                .Cols = NbrColonnesModificationsAvantImpression + .FixedCols
                .Row = 0
        
                '--- paramétrages de chaque colonne ---
                .Col = COLONNES_MODIFICATIONS_AVANT_IMPRESSION.C_NUM_LIGNES
                .ColWidth(.Col) = 2.5 * EPAISSEUR_CARACTERE: .Text = ""
'                Select Case NumFeuilleAppel
'
'                    Case FEUILLES.F_DEVIS
'                        '--- feuille des devis ---
'                        .Col = COLONNES_MODIFICATIONS_AVANT_IMPRESSION.C_COLONNE_1
'                        .ColWidth(.Col) = 20 * EPAISSEUR_CARACTERE: .Text = "Libellés destinés au client"
'                        .Col = COLONNES_MODIFICATIONS_AVANT_IMPRESSION.C_COLONNE_2
'                        .ColWidth(.Col) = 42.1 * EPAISSEUR_CARACTERE: .Text = "Renseignements"
'
'                    Case FEUILLES.F_FICHES_ATELIER
'                        '--- feuille des fiches d'atelier ---
'                        .Col = COLONNES_MODIFICATIONS_AVANT_IMPRESSION.C_COLONNE_1
'                        .ColWidth(.Col) = 62.1 * EPAISSEUR_CARACTERE: .Text = "Libellés destinés au client"
'
'                    Case Else
'
'                End Select
                
                '--- centrage des titres ---
                .Row = 0
                For a = 1 To Pred(.Cols)
                    .Col = a
                    .CellAlignment = flexAlignCenterCenter
                Next a
                
                '--- alignement des données ---
                .ColAlignment(COLONNES_MODIFICATIONS_AVANT_IMPRESSION.C_NUM_LIGNES) = flexAlignRightCenter
'                Select Case NumFeuilleAppel
'
'                    Case FEUILLES.F_DEVIS
'                        '--- feuille des devis ---
'                        .ColAlignment(COLONNES_MODIFICATIONS_AVANT_IMPRESSION.C_COLONNE_1) = flexAlignLeftCenter
'                        .ColAlignment(COLONNES_MODIFICATIONS_AVANT_IMPRESSION.C_COLONNE_2) = flexAlignLeftCenter
'
'                    Case FEUILLES.F_DEVIS, FEUILLES.F_FICHES_ATELIER
'                        '--- feuille des fiches d'atelier ---
'                        .ColAlignment(COLONNES_MODIFICATIONS_AVANT_IMPRESSION.C_COLONNE_1) = flexAlignLeftCenter
'
'                    Case Else
'
'                End Select
        
                '--- N° de lignes, vidage des champs ---
                .Col = 0
                For a = 1 To NbrLignesModificationsAvantImpression
                    .Row = a: .Text = CStr(a)
                Next a
            
                '--- fixer le curseur ---
                .Row = 1
                .Col = COLONNES_MODIFICATIONS_AVANT_IMPRESSION.C_COLONNE_1
            
                .Redraw = True
            
            End With
    
        Case GESTION_GRILLES.GG_VIDAGE_GRILLE
            '--- vidage du tableau ---
            ReDim TModificationsAvantImpression(1 To NbrLignesModificationsAvantImpression) As EnrModificationsAvantImpression
        
        Case GESTION_GRILLES.GG_TRANSFERT_DONNEES_VERS_GRILLE
            '--- transfert des données dans le tableau ---
            Select Case NumFeuilleAppel
                
'                Case FEUILLES.F_DEVIS
'                    '--- feuille des devis ---
'                    NumDevis = CritereRecherche
'                    PtrLigne = 1
'                    If RechercheDetailsDevis(NumDevis) = TROUVE Then
'                        For a = 1 To UBound(TTempEnrDetailsDevis())
'                            If TTempEnrDetailsDevis(a).AImprimer = True Then
'                                With TModificationsAvantImpression(PtrLigne)
'                                    NumArticle = TTempEnrDetailsDevis(a).NumArticle
'                                    If RechercheArticlesDevis(NumArticle) = TROUVE Then
'                                        .Colonne1 = DecodeLibelle(TTempEnrArticlesDevis.LibelleArticle, P_PARTIE_DROITE)
'                                    Else
'                                        .Colonne1 = ""
'                                    End If
'                                    .Colonne2 = TTempEnrDetailsDevis(a).Renseignements
'                                    If TTempEnrDetailsDevis(a).PrixTotalEuros = 0 Then
'                                        .Utilitaire = True      'signifie réaliser par eux
'                                    End If
'                                End With
'                                Inc PtrLigne
'                            End If
'                        Next a
'                    End If
'
'                Case FEUILLES.F_FICHES_ATELIER
'                    '--- feuille des fiches d'atelier ---
'                    NumFicheAtelier = CritereRecherche
'                    PtrLigne = 1
                    
                    '--- lecture des travaux ---
                    'If RechercheTravaux(NumFicheAtelier) = TROUVE Then
                    '    For a = 1 To UBound(TTempEnrTravaux())
                    '        If TTempEnrTravaux(a).AImprimer = True Then
                    '            With TModificationsAvantImpression(PtrLigne)
                    '                NumArticle = TTempEnrTravaux(a).NumArticle
                    '                'If RechercheArticlesDevis(NumArticle) = TROUVE Then
                    '                '    .Colonne1 = DecodeLibelle(TTempEnrArticlesDevis.LibelleArticle, P_PARTIE_DROITE)
                    '                'Else
                    '                '    .Colonne1 = ""
                    '                'End If
                    '                .Colonne2 = "" 'pas de colonne à l'écran
                    '            End With
                    '            Inc PtrLigne
                    '        End If
                    '    Next a
                    'End If
                    'Inc PtrLigne
            
                    '--- lecture de la fiche d'atelier ---
'                    If RechercheFichesAtelier(NumFicheAtelier) = TROUVE Then
'
'                        With TTempEnrFichesAtelier
'
'                            '--- quantité et désignation ---
'                            TModificationsAvantImpression(PtrLigne).Colonne1 = TCriteresImpression(3) & UN_ESPACE & .Designation
'                            TModificationsAvantImpression(PtrLigne).Colonne2 = ""
'                            Inc PtrLigne
'
'                            '--- matière ---
'                            TModificationsAvantImpression(PtrLigne).Colonne1 = "Matière : " & .Matiere
'                            TModificationsAvantImpression(PtrLigne).Colonne2 = ""
'                            Inc PtrLigne
'
'                            '--- plans ---
'                            If .NumPlans <> "" Then
'                                If .NbrPlans = 0 Or .NbrPlans = 1 Then
'                                    TModificationsAvantImpression(PtrLigne).Colonne1 = "SUIVANT LE PLAN N° " & .NumPlans
'                                Else
'                                    TModificationsAvantImpression(PtrLigne).Colonne1 = "SUIVANT LES PLANS N° " & .NumPlans
'                                End If
'                                If .PlansARendre = True Then
'                                    TModificationsAvantImpression(PtrLigne).Colonne1 = TModificationsAvantImpression(PtrLigne).Colonne1 & "  EN RETOUR"
'                                End If
'                                TModificationsAvantImpression(PtrLigne).Colonne2 = ""
'                                Inc PtrLigne
'                            End If

'                        End With
                    
'                    End If
            
                Case Else
            
            End Select
            
        Case GESTION_GRILLES.GG_COMPRESSION_GRILLE
            '--- compression des données ---
            PtrLigne = 1
            For a = 1 To NbrLignesModificationsAvantImpression
                If TModificationsAvantImpression(a).Colonne1 <> "" Or _
                    TModificationsAvantImpression(a).Colonne2 <> "" Then
                        TCopieModificationsAvantImpression(PtrLigne) = TModificationsAvantImpression(a)
                        Inc PtrLigne
                End If
            Next a
            For a = 1 To NbrLignesModificationsAvantImpression
                TModificationsAvantImpression(a) = TCopieModificationsAvantImpression(a)
            Next a
        
        Case GESTION_GRILLES.GG_CONSTRUCTION_GRILLE
            '--- construction de la grille ---
            With MSHFGModificationsAvantImpression
                
                '--- mémorisation des valeurs ligne, colonne ---
                MemLigne = .Row
                MemColonne = .Col
                .FocusRect = flexFocusNone
                .Redraw = False
                
                For a = 1 To NbrLignesModificationsAvantImpression
                
                    .Row = a
                    
                    If TModificationsAvantImpression(a).Colonne1 = "" And _
                       TModificationsAvantImpression(a).Colonne2 = "" Then
                        TModificationsAvantImpression(a) = FicheVide
                        For b = 1 To NbrColonnesModificationsAvantImpression
                            .Col = b
                            .Text = ""
                        Next b
                    Else
                        
'                        Select Case NumFeuilleAppel
'
'                            Case FEUILLES.F_DEVIS
'                                '--- feuilles des devis ---
'                                .Col = COLONNES_MODIFICATIONS_AVANT_IMPRESSION.C_COLONNE_1
'                                .Text = TModificationsAvantImpression(a).Colonne1
'                                .Col = COLONNES_MODIFICATIONS_AVANT_IMPRESSION.C_COLONNE_2
'                                .Text = TModificationsAvantImpression(a).Colonne2
'
'                            Case FEUILLES.F_FICHES_ATELIER
'                                '--- feuilles des fiches d'atelier
'                                .Col = COLONNES_MODIFICATIONS_AVANT_IMPRESSION.C_COLONNE_1
'                                .Text = TModificationsAvantImpression(a).Colonne1
'
'                            Case Else
'
'                        End Select
                    
                    End If
               
                Next a
            
                '--- restitution des valeurs ligne, colonne ---
                .Redraw = True
                .Row = MemLigne
                .Col = MemColonne
                .FocusRect = flexFocusHeavy
            
            End With
        
        Case Else
            
    End Select

End Sub

'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Rôle      : Déplace une ligne dans la grille des détails
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Private Sub DeplacementLigne()

    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next

    '--- déclaration ---
    Dim a As Integer, _
            PtrLigne As Integer
    Dim TFicheDepart As EnrModificationsAvantImpression
    
    '--- déclaration de certaine variable ---
    ReDim TCopieModificationsAvantImpression(1 To NbrLignesModificationsAvantImpression) As EnrModificationsAvantImpression

    If LigneDepartDeplacement > 0 And LigneDepartDeplacement < NbrLignesModificationsAvantImpression And _
       LigneArriveeDeplacement > 0 And LigneArriveeDeplacement < NbrLignesModificationsAvantImpression And _
       LigneDepartDeplacement <> LigneArriveeDeplacement Then
        
        '--- affectation ---
        TFicheDepart = TModificationsAvantImpression(LigneDepartDeplacement)
    
        '--- suppression à la ligne de départ ---
        PtrLigne = 1
        For a = 1 To NbrLignesModificationsAvantImpression
            If a <> LigneDepartDeplacement Then
                TCopieModificationsAvantImpression(PtrLigne) = TModificationsAvantImpression(a)
                Inc PtrLigne
            End If
        Next a
        
        '--- transfert dans le tableau ---
        For a = 1 To NbrLignesModificationsAvantImpression
            TModificationsAvantImpression(a) = TCopieModificationsAvantImpression(a)
        Next a
        
        '--- fixation de l'arrivée en fonction du sens de déplacement ---
        If LigneArriveeDeplacement > LigneDepartDeplacement Then
            LigneArriveeDeplacement = Pred(LigneArriveeDeplacement)
        End If
        If LigneArriveeDeplacement < 1 Then LigneArriveeDeplacement = 1
        If LigneArriveeDeplacement > NbrLignesModificationsAvantImpression Then LigneArriveeDeplacement = NbrLignesModificationsAvantImpression
        
        '--- insertion à la ligne d'arrivée ---
        PtrLigne = 1
        For a = 1 To NbrLignesModificationsAvantImpression
            If a = LigneArriveeDeplacement Then
                TCopieModificationsAvantImpression(PtrLigne) = TFicheDepart
                Inc PtrLigne
            End If
            If PtrLigne <= NbrLignesModificationsAvantImpression Then
                TCopieModificationsAvantImpression(PtrLigne) = TModificationsAvantImpression(a)
                Inc PtrLigne
            End If
            If PtrLigne >= NbrLignesModificationsAvantImpression Then Exit For
        Next a
        
        '--- transfert dans le tableau ---
        For a = 1 To NbrLignesModificationsAvantImpression
            TModificationsAvantImpression(a) = TCopieModificationsAvantImpression(a)
        Next a
        
        '--- reconstruction de la grille ---
        GestionModificationsAvantImpression GG_CONSTRUCTION_GRILLE
    
    End If
    
End Sub

'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Rôle      : Permet l'édition des modifications avant impression
' Entrées : KeyAscii -> Code ASCII de la touche frappée
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Private Sub EditionModificationsAvantImpression(ByRef KeyAscii As Integer)

    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next
    
    '--- déclaration ---
    
    '--- édition uniquement sur les bonnes colonnes ---
    Select Case MSHFGModificationsAvantImpression.Col
    
        Case COLONNES_MODIFICATIONS_AVANT_IMPRESSION.C_COLONNE_1, _
                 COLONNES_MODIFICATIONS_AVANT_IMPRESSION.C_COLONNE_2
   
            With TBEditionModificationsAvantImpression
   
                '--- affiche le contrôle texte au bon endroit (dans la cellule) ---
                .Move MSHFGModificationsAvantImpression.Left + MSHFGModificationsAvantImpression.CellLeft, _
                           MSHFGModificationsAvantImpression.Top + MSHFGModificationsAvantImpression.CellTop, _
                           MSHFGModificationsAvantImpression.CellWidth, _
                           MSHFGModificationsAvantImpression.CellHeight

                '--- paramètres de contrôle texte en fonction de la cellule ---
                Select Case MSHFGModificationsAvantImpression.Col
                    Case COLONNES_MODIFICATIONS_AVANT_IMPRESSION.C_COLONNE_1: .Alignment = vbLeftJustify
                    Case COLONNES_MODIFICATIONS_AVANT_IMPRESSION.C_COLONNE_2: .Alignment = vbLeftJustify
                    Case Else
                End Select
    
                '--- analyse du caractère qui a été tapé ---
                Select Case KeyAscii
                    
                    Case 0 To Pred(vbKeyBack), Succ(vbKeyBack) To Pred(vbKeyReturn), Succ(vbKeyReturn) To vbKeySpace
                        '--- du code 0 à l'espace (sauf retour arrière, Entrée) cela signifie une modification du texte en cours ---
                        .Text = MSHFGModificationsAvantImpression.Text
                        .Tag = .Text
                        .SelStart = 1000
                        .Visible = True
                        .SetFocus
                    
                    Case vbKeyBack
                        '--- touche retour arrière ---
                        .Text = ""
                        .Tag = MSHFGModificationsAvantImpression.Text
                        .SelStart = 1
                        .Visible = True
                        .SetFocus
                        TBeditionmodificationsavantimpression_Change
             
                    Case vbKeyReturn
                        '--- touche Entrée ---
                        With MSHFGModificationsAvantImpression
'                            Select Case .Col
'
'                                Case COLONNES_MODIFICATIONS_AVANT_IMPRESSION.C_COLONNE_1
'                                    Select Case NumFeuilleAppel
'                                        Case FEUILLES.F_DEVIS
'                                            .Col = COLONNES_MODIFICATIONS_AVANT_IMPRESSION.C_COLONNE_2
'                                        Case FEUILLES.F_FICHES_ATELIER
'                                            If .Row < .Rows - 1 Then .Row = .Row + 1
'                                        Case Else
'                                    End Select
'
'                                Case COLONNES_MODIFICATIONS_AVANT_IMPRESSION.C_COLONNE_2
'                                    If .Row < .Rows - 1 Then .Row = .Row + 1
'                                    .Col = COLONNES_MODIFICATIONS_AVANT_IMPRESSION.C_COLONNE_1
'
'                                Case Else
'                            End Select
                        End With

                    Case Else
                        '--- tout autre élément signifie le remplacement du texte en cours ---
                        .Text = ""
                        .Tag = MSHFGModificationsAvantImpression.Text
                        .SelStart = 1
                        .Visible = True
                        .SetFocus
                        SendKeys Chr(KeyAscii)
    
                End Select
    
            End With
    
        Case Else
    
    End Select
    
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
' Rôle      : Effectue le paramètrage de la feuille
' Entrées :  NumFeuilleAppel_ -> Numéro de feuille ayant lancé l'appel
'                 CritereRecherche_ -> Critère de recherche
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Sub ParametrageFeuille(ByVal NumFeuilleAppel_ As Long, _
                                                   ByVal CritereRecherche_ As String)

    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next

    '--- affectation ---
    NumFeuilleAppel = NumFeuilleAppel_
    CritereRecherche = CritereRecherche_

    '--- initialisation des variables ---
    Select Case NumFeuilleAppel

        'Case FEUILLES.F_DEVIS
            '--- feuille des devis ---
            'NbrLignesModificationsAvantImpression = NBR_LIGNES_DETAILS_DEVIS
            'NbrColonnesModificationsAvantImpression = 2
            'TitreFeuille = "Modification des travaux du DEVIS avant impression"
        
        'Case FEUILLES.F_FICHES_ATELIER
            '--- feuille des fiches d'atelier ---
            'NbrLignesModificationsAvantImpression = NBR_LIGNES_TRAVAUX
            'NbrColonnesModificationsAvantImpression = 1
            'TitreFeuille = "Modification des travaux du BON DE LIVRAISON avant impression"
        
        Case Else

    End Select

    '--- titre de la feuille et des messages ---
    Me.Caption = TitreFeuille
    TitreMessages = INDICATIF_PROGRAMME & TitreFeuille
        
    '--- affectation ---
    RepereFeuilleCritereRecherche = Str(NumFeuilleAppel) & "-" & CritereRecherche
    
    '--- construction de la grille ---
    GestionModificationsAvantImpression GG_INITIALISATION_GRILLE
    If MemRepereFeuilleCritereRecherche <> RepereFeuilleCritereRecherche Then
        GestionModificationsAvantImpression GG_TRANSFERT_DONNEES_VERS_GRILLE
        MemRepereFeuilleCritereRecherche = RepereFeuilleCritereRecherche
    End If
    GestionModificationsAvantImpression GG_CONSTRUCTION_GRILLE

End Sub


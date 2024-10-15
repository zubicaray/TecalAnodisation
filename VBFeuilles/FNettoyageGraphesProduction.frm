VERSION 5.00
Begin VB.Form FNettoyageGraphesProduction 
   Caption         =   "Nettoyage des graphes de production"
   ClientHeight    =   12675
   ClientLeft      =   4770
   ClientTop       =   1230
   ClientWidth     =   19725
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   12675
   ScaleWidth      =   19725
   Begin VB.PictureBox PBRenseignementsFenetre 
      Align           =   1  'Align Top
      BackColor       =   &H00FF0000&
      Height          =   375
      Left            =   0
      ScaleHeight     =   315
      ScaleWidth      =   19665
      TabIndex        =   4
      Top             =   0
      Width           =   19725
      Begin VB.Label LRenseignementsFenetre 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00008000&
         BackStyle       =   0  'Transparent
         Caption         =   "NETTOYAGE DES GRAPHES DE PRODUCTION"
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
   Begin VB.PictureBox PBBoutons 
      Align           =   2  'Align Bottom
      DrawStyle       =   2  'Dot
      DrawWidth       =   16891
      Height          =   1095
      Left            =   0
      ScaleHeight     =   1035
      ScaleWidth      =   19665
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   11580
      Width           =   19725
      Begin VB.CommandButton CBActualiser 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Caption         =   "Actualiser"
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
         Left            =   14940
         MaskColor       =   &H00FF00FF&
         Style           =   1  'Graphical
         TabIndex        =   7
         ToolTipText     =   " Actualiser les données "
         Top             =   105
         UseMaskColor    =   -1  'True
         Width           =   1515
      End
      Begin VB.CommandButton CBNettoyageGraphesProduction 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Nettoyage des graphes de plus de 1 an"
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
         Left            =   10080
         MaskColor       =   &H00FF00FF&
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   75
         UseMaskColor    =   -1  'True
         Width           =   4635
      End
      Begin VB.CommandButton CBQuitter 
         BackColor       =   &H00FFFFFF&
         Cancel          =   -1  'True
         Caption         =   "&Quitter"
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
         Left            =   16680
         MaskColor       =   &H00FF00FF&
         Style           =   1  'Graphical
         TabIndex        =   2
         ToolTipText     =   " Quitter cette fenêtre "
         Top             =   105
         UseMaskColor    =   -1  'True
         Width           =   1515
      End
      Begin VB.Shape SFocus 
         BorderColor     =   &H000000FF&
         BorderWidth     =   5
         Height          =   405
         Left            =   8220
         Top             =   150
         Visible         =   0   'False
         Width           =   420
      End
   End
   Begin VB.FileListBox FLBFichiersTracabilite 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6810
      Left            =   1440
      Pattern         =   "*.TRA"
      TabIndex        =   0
      Top             =   720
      Width           =   5415
   End
   Begin VB.Label LEtatsNettoyage 
      Alignment       =   2  'Center
      BackColor       =   &H0000FF00&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1920
      TabIndex        =   6
      Top             =   360
      Width           =   1155
   End
End
Attribute VB_Name = "FNettoyageGraphesProduction"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Rôle                    : Fenêtre de nettoyage des graphes de la production
' Nom                    : FNettoyageGraphesProduction.frm
' Date de création : 24/10/2011
' Détails                :
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

'--- déclarations obligatoires ---
Option Explicit

'--- options générales ---
Option Base 1
DefVar A-Z

'--- constantes privées ---
Private Const TITRE_FENETRE As String = "NETTOYAGE DES GRAPHES DE PRODUCTION"
Private Const TITRE_MESSAGES As String = TITRE_FENETRE

Private Const CHEMIN_FICHIERS = "\\Srv2003\Graphes de production"
Private Const EXTENSION_FICHIERS = "*.TRA"

'--- énumérations privées ---

'--- types privées ---

'--- variables privées ---
Private PremiereActivation As Boolean
Private InterdireEvenements As Boolean                                    'pour interdire certains évènements

'--- tableaux privés ---

'--- variables publiques ---
Public NumFenetre As Long                                                          'numéro de la fenêtre lorsqu'elle devient active

Private Sub Command1_Click()
    
End Sub

Private Sub CBActualiser_Click()
    
    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next

    With FLBFichiersTracabilite
        .Path = CHEMIN_FICHIERS
        .Pattern = EXTENSION_FICHIERS
        .Refresh
    End With

End Sub

Private Sub CBActualiser_GotFocus()
    
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

Private Sub CBActualiser_LostFocus()
    On Error Resume Next
    SFocus.Visible = False
End Sub

Private Sub CBNettoyageGraphesProduction_Click()
    
    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next
    
    '--- déclaration ---
    
    '--- affectation ---
    Dim a As Long                                     'pour les boucles FOR...NEXT
    Dim CptFichiers  As Long                    'compteur de fichiers
    
    Dim Texte As String                            'représente un texte quelconque
    
    Dim DateFichier As Date                     'date du fichier
    Dim DateReferenceTemps As Date    'date de référence dans le temps
    
    '--- calcul de la date de référence ---
    DateReferenceTemps = DateAdd("yyyy", -1, Now)
    
    With FLBFichiersTracabilite

        For a = 0 To .ListCount - 1
    
            '--- pointer le fichier ---
            .ListIndex = a
    
            '--- extraction de la date du fichier ---
            DateFichier = FileDateTime(.Path & "\" & .FileName)
            
            '--- destruction des fichiers ---
            If DateFichier < DateReferenceTemps Then
                
                '--- suppression di fichiers ---
                Kill .Path & "\" & .FileName

                '--- incrémentation du compteur de fichiers ---
                Inc CptFichiers

                '--- affectation du texte pour le message ---
                Texte = "SUPPRESSION DU FICHIER = " & .Path & "\" & .FileName

                '--- affichage du nom du fichier détruit ---
                AffichageTexte LEtatsNettoyage, Texte, COULEURS.ROUGE_4, COULEURS.JAUNE_3
                
                '--- rafriachissement des évènements ---
                DoEvents
            
            End If
    
        Next a
    
    End With

    '--- actualisation ---
    Call CBActualiser_Click

    '--- affichage du champ vide de l'état du nettoyage ---
    Texte = "-"
    AffichageTexte LEtatsNettoyage, Texte, COULEURS.VERT_3, COULEURS.NOIR

    '--- RAZ de la variable demandant l'entretien des graphes de production ---
    EntretienGraphesProduction = False
    
    '--- message de remarque ---
    Select Case CptFichiers
        Case 0: Texte = "AUCUN FICHIER DE SUPPRIMER"
        Case 1: Texte = "1 FICHIER SUPPRIME AVEC SUCCES"
        Case Else: Texte = CptFichiers & " FICHIERS SUPPRIMES AVEC SUCCES"
    End Select

    '--- affichage du message ---
    Bidon = AppelFenetre(F_MESSAGE, _
                                  TITRE_MESSAGES, _
                                  vbCrLf & vbCrLf & vbCrLf & Texte & vbCrLf & vbCrLf, _
                                  TYPES_MESSAGES.T_REMARQUE, _
                                  TYPES_BOUTONS.T_CONFIRMER, _
                                  EMPLACEMENT_FOCUS.E_SUR_CONFIRMER)

End Sub

Private Sub CBNettoyageGraphesProduction_GotFocus()
    
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

Private Sub CBNettoyageGraphesProduction_LostFocus()
    On Error Resume Next
    SFocus.Visible = False
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

Private Sub Form_Activate()
    
    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next
    
    '--- renseigne la fenêtre principale ---
    RenseigneFPrincipale
    
    '--- placement du focus ---
    If PremiereActivation = False Then
        Me.Refresh
        Call CBActualiser_Click
        PremiereActivation = True
    End If

End Sub

Private Sub PBBoutons_Resize()
    
    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next
    
    '--- calculs de l'emplacement des boutons ---
    CBQuitter.Left = PBBoutons.ScaleWidth - MARGES.M_BORD_DROIT - CBQuitter.Width
    CBActualiser.Left = CBQuitter.Left - MARGES.M_ENTRE_BOUTONS - CBActualiser.Width
    CBNettoyageGraphesProduction.Left = CBActualiser.Left - MARGES.M_ENTRE_BOUTONS - CBNettoyageGraphesProduction.Width

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
' Rôle      : Initialise la fenêtre (chargement ou en vue de la rendre visible)
' Entrées :
' Retours :
' Détails  :
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

End Sub

'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Rôle      : Décharge la fenêtre
' Entrées :
' Retours :
' Détails  :
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Private Sub DechargeFenetre()
    
    '--- aiguillage en cas d'erreur ---
    On Error Resume Next
    
    '--- affectation ---
    PremiereActivation = False
    
    '--- curseur souris par défaut ---
    SourisEnAttente False

    '--- neutralisation du timer ---
    
    '--- déchargement de la fenêtre ---
    Me.Visible = False
    Unload Me
    Set OccFNettoyageGraphesProduction = Nothing
    
End Sub

'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Rôle      : Change le curseur de la souris en fonction de l'attente
' Entrées : AttenteOuiNon -> TRUE   = Curseur en forme de sablier
'                                             FALSE = Curseur par défaut
' Retours :
' Détails  :
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

Private Sub PBRenseignementsFenetre_Resize()
    
    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next
    
    '--- déclaration ---
    
    With PBRenseignementsFenetre
        
        '--- calculs des emplacements ---
        LRenseignementsFenetre.Left = .ScaleLeft
        LRenseignementsFenetre.Top = .ScaleTop + 30
        LRenseignementsFenetre.Width = .ScaleWidth
        LRenseignementsFenetre.Height = .ScaleHeight
    
        '--- calculs de l'emplacement de la barre des états ---
        LEtatsNettoyage.Left = 0
        LEtatsNettoyage.Top = .Height
        LEtatsNettoyage.Width = .ScaleWidth + Screen.TwipsPerPixelX
    
        '--- calculs de l'emplacement de la liste des fichiers de traçabilité ---
        FLBFichiersTracabilite.Left = 0
        FLBFichiersTracabilite.Top = .Height + LEtatsNettoyage.Height + 2 * Screen.TwipsPerPixelX
        FLBFichiersTracabilite.Height = Me.ScaleHeight - .Height - LEtatsNettoyage.Height - PBBoutons.Height - Screen.TwipsPerPixelY
        FLBFichiersTracabilite.Width = .ScaleWidth
    
    End With

End Sub

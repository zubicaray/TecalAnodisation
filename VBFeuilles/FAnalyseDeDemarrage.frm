VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form FAnalyseDeDemarrage 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   " Analyse de d�marrage"
   ClientHeight    =   11895
   ClientLeft      =   1665
   ClientTop       =   2385
   ClientWidth     =   20820
   ControlBox      =   0   'False
   Icon            =   "FAnalyseDeDemarrage.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   11895
   ScaleWidth      =   20820
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox PBFond 
      AutoSize        =   -1  'True
      Height          =   1830
      Left            =   0
      Picture         =   "FAnalyseDeDemarrage.frx":000C
      ScaleHeight     =   1770
      ScaleWidth      =   20715
      TabIndex        =   3
      Top             =   0
      Width           =   20775
      Begin VB.Label LTitre 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "TECAL VERBRUGGE"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   48
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   1110
         Index           =   1
         Left            =   2100
         TabIndex        =   6
         Top             =   660
         Width           =   16575
      End
      Begin VB.Label LTitre 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Ligne d'anodisation"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   36
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   1110
         Index           =   0
         Left            =   2100
         TabIndex        =   4
         Top             =   0
         Width           =   16575
      End
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFGAnalyseDemarrage 
      Height          =   8895
      Left            =   0
      TabIndex        =   2
      Top             =   1860
      Width           =   20775
      _ExtentX        =   36645
      _ExtentY        =   15690
      _Version        =   393216
      GridColor       =   12632256
      WordWrap        =   -1  'True
      ScrollBars      =   0
      AllowUserResizing=   2
      BandDisplay     =   1
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
      _Band(0).Cols   =   2
      _Band(0).GridLinesBand=   1
      _Band(0).TextStyleBand=   0
      _Band(0).TextStyleHeader=   0
   End
   Begin VB.PictureBox PBBoutons 
      Align           =   2  'Align Bottom
      Height          =   1095
      Left            =   0
      Picture         =   "FAnalyseDeDemarrage.frx":7766E
      ScaleHeight     =   1035
      ScaleWidth      =   20760
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   10800
      Width           =   20820
      Begin VB.CommandButton CBQuitter 
         BackColor       =   &H00C0FFFF&
         Cancel          =   -1  'True
         Caption         =   "&Quitter"
         DownPicture     =   "FAnalyseDeDemarrage.frx":79D08
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   795
         Left            =   18540
         MaskColor       =   &H00FF00FF&
         Picture         =   "FAnalyseDeDemarrage.frx":7A40A
         Style           =   1  'Graphical
         TabIndex        =   5
         TabStop         =   0   'False
         ToolTipText     =   " Quitter cette fen�tre "
         Top             =   120
         UseMaskColor    =   -1  'True
         Visible         =   0   'False
         Width           =   2115
      End
      Begin VB.CommandButton CBContinuer 
         BackColor       =   &H00C0E0FF&
         Caption         =   "&Continuer l'analyse de d�marrage"
         DownPicture     =   "FAnalyseDeDemarrage.frx":7AB0C
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   795
         Left            =   120
         Picture         =   "FAnalyseDeDemarrage.frx":7B20E
         Style           =   1  'Graphical
         TabIndex        =   1
         ToolTipText     =   " Continuer l'analyse de d�marrage "
         Top             =   120
         UseMaskColor    =   -1  'True
         Visible         =   0   'False
         Width           =   18255
      End
      Begin VB.Shape SFocus 
         BorderColor     =   &H000000FF&
         BorderWidth     =   5
         Height          =   345
         Left            =   60
         Top             =   90
         Visible         =   0   'False
         Width           =   540
      End
   End
End
Attribute VB_Name = "FAnalyseDeDemarrage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' R�le                    : Fen�tre d'analyse de d�marrage
' Nom                    : FAnalyseDeDemarrage
' Date de cr�ation : 25/10/2000
' D�tails                :
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

'--- d�clarations obligatoires ---
Option Explicit

'--- options g�n�rales ---
Option Base 1
DefVar A-Z

'--- constantes priv�es ---
Private Const NBR_COLONNES_ANALYSE_DEMARRAGE  As Integer = 3
Private Const NBR_LIGNES_ANALYSE_DEMARRAGE  As Integer = 50

Private Const TITRE_FENETRE As String = "Analyse de d�marrage"
Private Const TITRE_MESSAGES As String = TITRE_FENETRE

'--- �num�rations priv�es ---
Private Enum COLONNES_ANALYSE_DEMARRAGE
    C_NUM_LIGNES = 0
    C_LIBELLE_MESSAGE = 1
    C_NUM_ERREUR = 2
    C_LIBELLE_ERREUR = 3
End Enum

'--- �num�rations publique ---
Public Enum TYPES_AFFICHAGE_ANALYSE_DEMARRAGE
    AFFICHAGE_LIBELLE = 0
    AFFICHAGE_ANALYSE = 1
End Enum

'--- variables priv�es ---
Private PremiereActivation As Boolean
Private DemandeArret As Boolean                'demande d'arr�t du d�filement dans la grille
Private NumLigneAnalyse As Integer            'num�ro de la ligne analys�

'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' R�le      : Affiche le r�sultat du contr�le d'une fonction au d�marrage du programme
' Entr�es :
' Retours :
' D�tails  :
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Sub ControleFonction(ByVal TypeAffichageAnalyseDemarrage As TYPES_AFFICHAGE_ANALYSE_DEMARRAGE, _
                                              ByVal LibelleOuResultatControle As Variant)
    
    '--- aiguillage en cas d'erreur ---
    On Error Resume Next

    '--- constantes priv�es ---
    Const NBR_LIGNES_AVANT_DECALAGE As Integer = 30
    
    '--- d�claration ---
    Static PtrDuControle As Integer
    Dim TLignesDecodees As Variant
    
    '--- forcer le premier plan ---
    Me.ZOrder
    
    '--- invisibilit� du bouton continuer ---
    Me.CBQuitter.Visible = False
    Me.CBContinuer.Visible = False
    
    With Me.MSHFGAnalyseDemarrage
    
        '--- affectation ---
        .Redraw = False
        .Row = NumLigneAnalyse
    
        '--- affichage du libell� ou du r�sultat du contr�le ---
        If TypeAffichageAnalyseDemarrage = AFFICHAGE_LIBELLE Then
        
            '--- affichage du libell� du message ---
            .Col = COLONNES_ANALYSE_DEMARRAGE.C_LIBELLE_MESSAGE
            .CellBackColor = COULEURS.VERT_2
            .CellForeColor = COULEURS.NOIR
            .Text = LibelleOuResultatControle
            
            '--- rafraichissement de l'affichage ---
            .Redraw = True
            .Refresh
        
        Else
                
            '--- affichage du r�sultat du contr�le ---
            Select Case LibelleOuResultatControle
        
                Case ""
                    '--- pas d'incident sur le contr�le ---
                    .Col = COLONNES_ANALYSE_DEMARRAGE.C_NUM_ERREUR
                    .CellBackColor = COULEURS.VERT_2
                    .CellForeColor = COULEURS.NOIR
                    .Text = "Bon"
                    
                    '--- affichage du libell� du message d'erreur ---
                    .Col = COLONNES_ANALYSE_DEMARRAGE.C_LIBELLE_ERREUR
                    .CellBackColor = COULEURS.VERT_2
                    .CellForeColor = COULEURS.NOIR
                    .Text = ""
                    
                    '--- rafraichissement de l'affichage ---
                    .Redraw = True
                    .Refresh
            
                Case Else
                    '--- changement de la couleur du libell� du message ---
                    .Col = COLONNES_ANALYSE_DEMARRAGE.C_LIBELLE_MESSAGE
                    .CellBackColor = COULEURS.ROUGE_3
                    .CellForeColor = COULEURS.JAUNE_3
                    
                    '--- affichage du num�ro de l'erreur ---
                    .Col = COLONNES_ANALYSE_DEMARRAGE.C_NUM_ERREUR
                    .CellBackColor = COULEURS.ROUGE_3
                    .CellForeColor = COULEURS.JAUNE_3
                    If IsNumeric(LibelleOuResultatControle) = True Then
                        .Text = LibelleOuResultatControle
                    Else
                        .Text = "-"
                    End If
                    
                    '--- affichage du libell� du message d'erreur ---
                    .Col = COLONNES_ANALYSE_DEMARRAGE.C_LIBELLE_ERREUR
                    .CellBackColor = COULEURS.ROUGE_3
                    .CellForeColor = COULEURS.JAUNE_3
                    If IsNumeric(LibelleOuResultatControle) = True Then
                        .Text = Error(LibelleOuResultatControle)
                    Else
                        .Text = LibelleOuResultatControle
                    
                        '--- calcul du nombre de lignes dans le texte � afficher ---
                        TLignesDecodees = Split(.Text, vbCrLf)
                        
                        '--- agrandissement de la ligne ---
                        .RowHeight(NumLigneAnalyse) = .RowHeight(NumLigneAnalyse) * (UBound(TLignesDecodees) + 1)
                    
                    End If
           
                    '--- visibilit� du bouton continuer ---
                    Me.CBQuitter.Visible = True
                    With Me.CBContinuer
                        .Visible = True
                        .SetFocus
                    End With
            
                    '--- rafraichissement de l'affichage ---
                    .Redraw = True
                    .Refresh
            
                    '--- arr�t complet de l'affichage ---
                    DemandeArret = True
                    Do While DemandeArret
                        DoEvents
                    Loop
           
            End Select
                    
            '--- colonne et ligne visible ---
            .LeftCol = COLONNES_ANALYSE_DEMARRAGE.C_LIBELLE_MESSAGE
            If .RowIsVisible(NumLigneAnalyse + 2) = False Then
                .TopRow = NumLigneAnalyse - NBR_LIGNES_AVANT_DECALAGE + 1
                .Refresh
            End If
        
            '--- incr�mentation ---
            Inc NumLigneAnalyse
        
        End If

    End With

End Sub

Private Sub CBContinuer_Click()
    
    '--- aiguillage en cas d'erreur ---
    On Error Resume Next

    '--- affectation ---
    DemandeArret = False

End Sub

Private Sub CBContinuer_GotFocus()
    
    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next

    '--- d�placement du focus sur le bouton ---
    With SFocus
        .Left = ActiveControl.Left
        .Top = ActiveControl.Top
        .Height = ActiveControl.Height
        .Width = ActiveControl.Width
        .Visible = True
    End With

End Sub

Private Sub CBContinuer_LostFocus()
    On Error Resume Next
    SFocus.Visible = False
End Sub

Private Sub CBQuitter_Click()
    On Error Resume Next
    End
End Sub

Private Sub CBQuitter_GotFocus()
    
    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next

    '--- d�placement du focus sur le bouton ---
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

    '--- aiguillage en cas d'erreur ---
    On Error Resume Next
    
    '--- affectation ---
    If PremiereActivation = False Then
        PremiereActivation = True
    End If

End Sub

Private Sub Form_Load()
    
    '--- aiguillage en cas d'erreur ---
    On Error Resume Next

    '--- centrage de la fenetre ---
    Centrefenetre Me, , , 200

    '--- gestion de la grille d'analyse de d�marrage ---
    GestionAnalyseDemarrage GG_INITIALISATION

    '--- affectation ---
    NumLigneAnalyse = 1

End Sub

'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' R�le      : Gestion de l'analyse de d�marrage
' Entr�es : EtatSouhaite -> Fonction de l'�num�ration GESTION_GRILLES
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Private Sub GestionAnalyseDemarrage(ByVal EtatSouhaite As GESTION_GRILLES)
    
    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next
    
    '--- constantes priv�es ---

    '--- d�claration ---
    Dim a As Integer
    
    Select Case EtatSouhaite
    
        Case GESTION_GRILLES.GG_INITIALISATION
            '--- initialisation de la grille de l'analyse de d�marrage ---
            With MSHFGAnalyseDemarrage
                
                .Redraw = False
                
                .Clear

                .FixedCols = 1
                .FixedRows = 1
                .Rows = NBR_LIGNES_ANALYSE_DEMARRAGE + .FixedRows
                .Cols = NBR_COLONNES_ANALYSE_DEMARRAGE + .FixedCols
                .Row = 0
        
                '--- couleurs ---
                .ForeColorFixed = COULEURS.BLANC
                .BackColorFixed = COULEURS.VERT_6
        
                '--- param�trages de chaque colonne ---
                .Col = COLONNES_ANALYSE_DEMARRAGE.C_NUM_LIGNES
                .ColWidth(.Col) = 3 * EPAISSEUR_CARACTERE: .Text = ""
                .Col = COLONNES_ANALYSE_DEMARRAGE.C_LIBELLE_MESSAGE
                .ColWidth(.Col) = 60 * EPAISSEUR_CARACTERE: .Text = "Analyse en cours"
                .Col = COLONNES_ANALYSE_DEMARRAGE.C_NUM_ERREUR
                .ColWidth(.Col) = 15 * EPAISSEUR_CARACTERE: .Text = "Etat / N� d'erreur"
                .Col = COLONNES_ANALYSE_DEMARRAGE.C_LIBELLE_ERREUR
                .ColWidth(.Col) = 70 * EPAISSEUR_CARACTERE + 30: .Text = "Libell� de l'erreur"

                '--- centrage des titres ---
                .Row = 0
                For a = 1 To Pred(.Cols)
                    .Col = a
                    .CellAlignment = flexAlignCenterCenter
                Next a
                
                '--- alignement des donn�es ---
                .ColAlignment(COLONNES_ANALYSE_DEMARRAGE.C_NUM_LIGNES) = flexAlignRightCenter
                .ColAlignment(COLONNES_ANALYSE_DEMARRAGE.C_LIBELLE_MESSAGE) = flexAlignLeftCenter
                .ColAlignment(COLONNES_ANALYSE_DEMARRAGE.C_NUM_ERREUR) = flexAlignCenterCenter
                .ColAlignment(COLONNES_ANALYSE_DEMARRAGE.C_LIBELLE_ERREUR) = flexAlignLeftCenter
        
                '--- N� de lignes, vidage des champs ---
                .Col = 0
                For a = 1 To NBR_LIGNES_ANALYSE_DEMARRAGE
                    .Row = a: .Text = CStr(a)
                Next a
            
                '--- fixer le curseur ---
                .Row = 1
                .Col = COLONNES_ANALYSE_DEMARRAGE.C_LIBELLE_MESSAGE
                
                .Redraw = True
            
            End With
    
        Case Else

    End Select

End Sub

Private Sub PBBoutons_Resize()
    
    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next
    
    '--- recalcul du focus apr�s d�placement ---
    With SFocus
        If .Visible = True Then
            .Left = ActiveControl.Left
            .Top = ActiveControl.Top
            .Height = ActiveControl.Height
            .Width = ActiveControl.Width
        End If
    End With

End Sub


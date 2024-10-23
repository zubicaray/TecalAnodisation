VERSION 5.00
Begin VB.Form FInformationsDefautsCommunicationAutomate 
   Caption         =   "INFORMATIONS SUR LES DEFAUTS DE COMMUNICATION AVEC UN AUTOMATE"
   ClientHeight    =   9795
   ClientLeft      =   2475
   ClientTop       =   3105
   ClientWidth     =   18015
   HelpContextID   =   80
   Icon            =   "FInformationsDefautsCommunicationAutomate.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   9795
   ScaleWidth      =   18015
   Begin VB.PictureBox PBRenseignementsFenetre 
      Align           =   1  'Align Top
      BackColor       =   &H00FF0000&
      Height          =   375
      Left            =   0
      Picture         =   "FInformationsDefautsCommunicationAutomate.frx":014A
      ScaleHeight     =   315
      ScaleWidth      =   17955
      TabIndex        =   3
      Top             =   0
      Width           =   18015
      Begin VB.Label LRenseignementsFenetre 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00008000&
         BackStyle       =   0  'Transparent
         Caption         =   "MAINTENANCE GEREE"
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
      ScaleWidth      =   17955
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   8700
      Width           =   18015
      Begin VB.CommandButton CBQuitter 
         BackColor       =   &H00FFFFFF&
         Cancel          =   -1  'True
         Caption         =   "&Quitter"
         DownPicture     =   "FInformationsDefautsCommunicationAutomate.frx":24A8C
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
         Left            =   16080
         MaskColor       =   &H00FF00FF&
         Picture         =   "FInformationsDefautsCommunicationAutomate.frx":2518E
         Style           =   1  'Graphical
         TabIndex        =   2
         ToolTipText     =   " Quitter cette fen�tre "
         Top             =   120
         UseMaskColor    =   -1  'True
         Width           =   1515
      End
      Begin VB.Shape SFocus 
         BorderColor     =   &H000000FF&
         BorderWidth     =   5
         Height          =   405
         Left            =   2460
         Top             =   240
         Visible         =   0   'False
         Width           =   600
      End
   End
   Begin VB.PictureBox IMGCodesErreurs 
      AutoSize        =   -1  'True
      Height          =   6375
      Left            =   1260
      Picture         =   "FInformationsDefautsCommunicationAutomate.frx":25890
      ScaleHeight     =   6315
      ScaleWidth      =   18240
      TabIndex        =   1
      Top             =   1980
      Width           =   18300
   End
End
Attribute VB_Name = "FInformationsDefautsCommunicationAutomate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' R�le                    : Fen�tre affichant les informations sur les d�fauts de communication avec un automate
' Nom                    : FInformationsDefautsCommunicationAutomate.frm
' Date de cr�ation : 02/10/2007
' D�tails                :
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

'--- d�clarations obligatoires ---
Option Explicit

'--- options g�n�rales ---
Option Base 1
DefVar A-Z
    
'--- constantes priv�es ---
Private Const TITRE_FENETRE As String = "Informations sur les d�fauts de communication avec un automate"
Private Const TITRE_MESSAGES As String = TITRE_FENETRE

'--- �num�rations priv�es ---

'--- types priv�es ---
    
'--- variables priv�es ---
Private PremiereActivation As Boolean

'--- variables publiques ---
Public NumFenetre As Long                                   'num�ro de la fen�tre lorsqu'elle devient active

Private Sub CBQuitter_Click()
    On Error Resume Next
    DechargeFenetre
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

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    
    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next
    
    '--- gestion des touches communes ---
    Call OccFSynoptique.GestionTouches(KeyCode, Shift)
    
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
    
    '--- d�claration ---
    
    '--- calculs de l'emplacement des boutons ---
    CBQuitter.Left = PBBoutons.ScaleWidth - MARGES.M_BORD_DROIT - CBQuitter.Width
    
    '--- centrage de l'image ---
    With IMGCodesErreurs
        .Left = (Me.ScaleWidth - .Width) / 2
        .Top = (Me.ScaleHeight - PBBoutons.Height - .Height) / 2
    End With
    
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

'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' R�le      : Initialise la fen�tre (chargement ou en vue de la rendre visible)
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Sub InitialisationFenetre()
    
    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next

    '--- d�claration ---

    '--- affectation ---
  
    '--- divers sur la Fenetre ---
    With Me
        .Caption = UCase(TITRE_FENETRE)
        .Picture = ImgFondDeFenetre
        .WindowState = vbMaximized
    End With
    
    '--- renseignements de la fen�tre ---
    LRenseignementsFenetre.Caption = UCase(TITRE_FENETRE)
    
    '--- fond de l'image des boutons ---
    PBBoutons.Picture = ImgFondDesBoutons
    
End Sub

'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' R�le      : Effectue le param�trage de la Fenetre
' Entr�es :
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Sub ParametrageFenetre()
    
End Sub

'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' R�le      : D�charge la Fenetre
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Private Sub DechargeFenetre()
    
    '--- aiguillage en cas d'erreur ---
    On Error Resume Next
    
    '--- affectation ---
    PremiereActivation = False
    
    '--- curseur souris par d�faut ---
    SourisEnAttente False

    '--- d�chargement de la fen�tre ---
    Me.Visible = False
    Unload Me
    Set OccFInformationsDefautsCommunicationAutomate = Nothing
    
End Sub

'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' R�le      : Change le curseur de la souris en fonction de l'attente
' Entr�es : AttenteOuiNon -> TRUE   = Curseur en forme de sablier
'                                             FALSE = Curseur par d�faut
' Retours :
' D�tails  :
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

Private Sub PBRenseignementsFenetre_Resize()
    
    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next
    
    '--- d�claration ---
    
    '--- calculs des emplacements ---
    With PBRenseignementsFenetre
        LRenseignementsFenetre.Left = .ScaleLeft
        LRenseignementsFenetre.Top = .ScaleTop + 30
        LRenseignementsFenetre.Width = .ScaleWidth
        LRenseignementsFenetre.Height = .ScaleHeight
    End With

End Sub


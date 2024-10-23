VERSION 5.00
Begin VB.Form FMessage 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   5085
   ClientLeft      =   7320
   ClientTop       =   5595
   ClientWidth     =   6630
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "MS Serif"
      Size            =   6.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5085
   ScaleWidth      =   6630
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox PBBoutons 
      Align           =   2  'Align Bottom
      DrawStyle       =   2  'Dot
      DrawWidth       =   16891
      Height          =   1155
      Left            =   0
      ScaleHeight     =   1095
      ScaleWidth      =   6570
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   3930
      Width           =   6630
      Begin VB.CommandButton CBAnnuler 
         Caption         =   "&Annuler"
         DownPicture     =   "FMessage.frx":0000
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   5040
         MaskColor       =   &H00FF00FF&
         Picture         =   "FMessage.frx":0702
         Style           =   1  'Graphical
         TabIndex        =   13
         ToolTipText     =   " Annuler les dernières modifications "
         Top             =   120
         UseMaskColor    =   -1  'True
         Width           =   1155
      End
      Begin VB.CommandButton CBNon 
         Cancel          =   -1  'True
         Caption         =   "&NON"
         DownPicture     =   "FMessage.frx":0E04
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   3660
         MaskColor       =   &H00FF00FF&
         Picture         =   "FMessage.frx":1506
         Style           =   1  'Graphical
         TabIndex        =   12
         ToolTipText     =   " Ne pas approuver le message "
         Top             =   120
         UseMaskColor    =   -1  'True
         Width           =   1155
      End
      Begin VB.CommandButton CBConfirmer 
         Caption         =   "&Confirmer"
         DownPicture     =   "FMessage.frx":1C08
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   360
         MaskColor       =   &H00FF00FF&
         Picture         =   "FMessage.frx":230A
         Style           =   1  'Graphical
         TabIndex        =   14
         ToolTipText     =   " Confirmer la lecture du message "
         Top             =   120
         UseMaskColor    =   -1  'True
         Width           =   1155
      End
      Begin VB.CommandButton CBOui 
         Caption         =   "&OUI"
         DownPicture     =   "FMessage.frx":2A0C
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   1740
         MaskColor       =   &H00FF00FF&
         Picture         =   "FMessage.frx":310E
         Style           =   1  'Graphical
         TabIndex        =   11
         ToolTipText     =   " Approuver le message "
         Top             =   120
         UseMaskColor    =   -1  'True
         Width           =   1155
      End
      Begin VB.Shape SAnnuler 
         BorderColor     =   &H000000FF&
         BorderWidth     =   7
         Height          =   855
         Left            =   5040
         Top             =   120
         Visible         =   0   'False
         Width           =   1155
      End
      Begin VB.Shape SOui 
         BorderColor     =   &H000000FF&
         BorderWidth     =   7
         Height          =   855
         Left            =   1740
         Top             =   120
         Visible         =   0   'False
         Width           =   1155
      End
      Begin VB.Shape SNon 
         BorderColor     =   &H000000FF&
         BorderWidth     =   7
         Height          =   855
         Left            =   3660
         Top             =   120
         Visible         =   0   'False
         Width           =   1155
      End
      Begin VB.Shape SConfirmer 
         BorderColor     =   &H000000FF&
         BorderWidth     =   7
         Height          =   855
         Left            =   360
         Top             =   120
         Visible         =   0   'False
         Width           =   1155
      End
   End
   Begin VB.PictureBox PBCommentaires 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3015
      Left            =   0
      ScaleHeight     =   2955
      ScaleWidth      =   6555
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   900
      Width           =   6615
      Begin VB.Label LCommentaires 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   315
         Index           =   7
         Left            =   120
         TabIndex        =   8
         Top             =   2580
         Width           =   6315
         WordWrap        =   -1  'True
      End
      Begin VB.Label LCommentaires 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   315
         Index           =   6
         Left            =   120
         TabIndex        =   7
         Top             =   2220
         Width           =   6315
         WordWrap        =   -1  'True
      End
      Begin VB.Label LCommentaires 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   315
         Index           =   5
         Left            =   120
         TabIndex        =   6
         Top             =   1860
         Width           =   6315
         WordWrap        =   -1  'True
      End
      Begin VB.Label LCommentaires 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   315
         Index           =   0
         Left            =   120
         TabIndex        =   5
         Top             =   60
         Width           =   6315
         WordWrap        =   -1  'True
      End
      Begin VB.Label LCommentaires 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   315
         Index           =   1
         Left            =   120
         TabIndex        =   4
         Top             =   420
         Width           =   6315
         WordWrap        =   -1  'True
      End
      Begin VB.Label LCommentaires 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   315
         Index           =   2
         Left            =   120
         TabIndex        =   3
         Top             =   780
         Width           =   6315
         WordWrap        =   -1  'True
      End
      Begin VB.Label LCommentaires 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   315
         Index           =   3
         Left            =   120
         TabIndex        =   2
         Top             =   1140
         Width           =   6315
         WordWrap        =   -1  'True
      End
      Begin VB.Label LCommentaires 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   315
         Index           =   4
         Left            =   120
         TabIndex        =   1
         Top             =   1500
         Width           =   6315
         WordWrap        =   -1  'True
      End
   End
   Begin VB.Label LTitreMessage 
      BackColor       =   &H00E0E0E0&
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
      ForeColor       =   &H00C00000&
      Height          =   315
      Left            =   0
      TabIndex        =   15
      Top             =   600
      Width           =   6615
   End
   Begin VB.Label LTitreGeneral 
      Alignment       =   2  'Center
      BackColor       =   &H0000FF00&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Titre "
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   615
      Left            =   0
      TabIndex        =   9
      Top             =   0
      Width           =   6615
   End
End
Attribute VB_Name = "FMessage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Rôle                    : Affichage des messages
' Nom                    : FMessage.frm
' Date de création : 30/10/2000
' Détails                :
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

'--- déclarations obligatoires ---
Option Explicit

'--- options générales ---
Option Base 1
DefVar A-Z

'--- constantes privées ---
Private Const TITRE_FENETRE As String = "Message"
Private Const TITRE_MESSAGES As String = TITRE_FENETRE

'--- énumérations privées ---
Public Enum TYPES_MESSAGES
    T_REMARQUE = 0
    T_AVERTISSEMENT = 1
    T_ATTENTION = 2
End Enum

Public Enum TYPES_BOUTONS
    T_OUI_NON = 0
    T_OUI_NON_ANNULER = 1
    T_CONFIRMER = 2
End Enum

Public Enum EMPLACEMENT_FOCUS
    E_SUR_CONFIRMER = 0
    E_SUR_OUI = 0
    E_SUR_NON = 1
    E_SUR_ANNULER = 2
End Enum

'--- variables privées ---
Private PremiereActivation As Boolean
Private Attente As Boolean                                'TRUE = attente opérateur en cours, FALSE = un évènement est intervenu

Private TypesBoutons As Integer                      'indique les types de boutons choisis
Private ChoixFocus As Integer                           'choix du focus pour les boutons

'--- variables publiques ---
Public NumFenetre As Long                                'numéro de la fenêtre lorsqu'elle devient active
Public VariableRetourneefenetre As Long          'VBOK ou VBYes ou VBNo

Private Sub CBAnnuler_Click()
    
    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next
    
    '--- affectation ---
    VariableRetourneefenetre = vbCancel
    Attente = False

End Sub

Private Sub CBAnnuler_GotFocus()
    
    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next
    
    '--- affichage du focus ---
    Select Case TypesBoutons
        Case T_OUI_NON_ANNULER
            SOui.Visible = False
            SNon.Visible = False
            SAnnuler.Visible = True
        Case Else
    End Select

End Sub

Private Sub CBConfirmer_Click()
    
    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next
    
    '--- affectation ---
    VariableRetourneefenetre = vbOK
    Attente = False

End Sub

Private Sub CBConfirmer_GotFocus()
    
    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next
    
    SConfirmer.Visible = True

End Sub

Private Sub CBNon_Click()
    
    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next
    
    '--- affectation ---
    VariableRetourneefenetre = vbNo
    Attente = False

End Sub

Private Sub CBNon_GotFocus()
    
    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next
    
    '--- affichage du focus ---
    Select Case TypesBoutons
        Case T_OUI_NON
            SOui.Visible = False
            SNon.Visible = True
        Case T_OUI_NON_ANNULER
            SOui.Visible = False
            SAnnuler.Visible = False
            SNon.Visible = True
        Case Else
    End Select
    
End Sub

Private Sub CBOui_Click()
    
    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next
    
    '--- affectation ---
    VariableRetourneefenetre = vbYes
    Attente = False

End Sub

Private Sub CBOui_GotFocus()
    
    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next
    
    '--- affichage du focus ---
    Select Case TypesBoutons
        Case T_OUI_NON
            SNon.Visible = False
            SOui.Visible = True
        Case T_OUI_NON_ANNULER
            SAnnuler.Visible = False
            SNon.Visible = False
            SOui.Visible = True
        Case Else
    End Select

End Sub

Private Sub Form_Activate()

    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next
    
    '--- renseigne la fenêtre principale ---
    RenseigneFPrincipale
    
    If PremiereActivation = False Then
        
        '--- rafraichisement de la fenêtre ---
        Me.Refresh
        
        '--- placement du focus ---
        Select Case TypesBoutons
        
            Case TYPES_BOUTONS.T_OUI_NON
                '--- boutons OUI et NON ---
                Select Case ChoixFocus
                    Case EMPLACEMENT_FOCUS.E_SUR_OUI: CBOui.SetFocus
                    Case EMPLACEMENT_FOCUS.E_SUR_NON: CBNon.SetFocus
                    Case Else
                End Select
            
            Case TYPES_BOUTONS.T_OUI_NON_ANNULER
                '--- boutons OUI et NON et ANNULER ---
                Select Case ChoixFocus
                    Case EMPLACEMENT_FOCUS.E_SUR_OUI: CBOui.SetFocus
                    Case EMPLACEMENT_FOCUS.E_SUR_NON: CBNon.SetFocus
                    Case EMPLACEMENT_FOCUS.E_SUR_ANNULER: CBAnnuler.SetFocus
                    Case Else
                End Select
    
            Case TYPES_BOUTONS.T_CONFIRMER
                '--- bouton confirmer ---
                CBConfirmer.SetFocus
        
            Case Else
    
        End Select
        
        '--- anti-rebond ---
        PremiereActivation = True
        
        '--- lancement de l'attente ---
        AttenteEvenement
    
    End If

End Sub

Private Sub Form_Load()
    
    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next
    
    '--- divers ---
    Centrefenetre Me, TITRE_MESSAGES

End Sub

'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Rôle      : Gére l'attente d'un évènement
' Entrées :
' Retours :
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Private Sub AttenteEvenement()
    
    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next

    '--- affectation ---
    Attente = True

    Do While Attente
        DoEvents
    Loop

    '--- cacher la fenêtre ---
    Me.Hide
    
    '--- affectation ---
    PremiereActivation = False

End Sub

'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Rôle      : Effectue le paramètrage de la fenêtre
' Entrées :
' Retours :
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Sub ParametrageFenetre(ByVal TitreMessage As String, _
                                                   ByVal LibelleMessage As String, _
                                                   ByVal TypeMessage As Integer, _
                                                   ByVal TypesBoutons_ As Integer, _
                                                   ByVal ChoixFocus_ As Integer)
                                                   
    '--- aiguillage en cas d'erreur ---
    On Error Resume Next

    '--- déclaration ---
    Dim a As Integer
    Dim PositionDuCrLf As Integer
    Dim PositionFinControle As Integer
    Dim Controle As String
    Dim Commentaire As String
    
    '--- affectation ---
    TypesBoutons = TypesBoutons_
    ChoixFocus = ChoixFocus_
    
    With FMessage
    
        '--- affichage et centrage du titre général --
        With .LTitreGeneral
        
            Select Case TypeMessage
            
                Case TYPES_MESSAGES.T_REMARQUE
                    '--- titre = remarque ---
                    .BackColor = COULEURS.VERT_3
                    .ForeColor = COULEURS.NOIR
                    .Caption = " REMARQUE "
            
                Case TYPES_MESSAGES.T_AVERTISSEMENT
                    '--- titre = avertissement ---
                    .BackColor = COULEURS.ORANGE_3
                    .ForeColor = COULEURS.NOIR
                    .Caption = " AVERTISSEMENT "
                                
                Case TYPES_MESSAGES.T_ATTENTION
                    '--- titre = attention ---
                    .BackColor = COULEURS.ROUGE_3
                    .ForeColor = COULEURS.JAUNE_3
                    .Caption = " ATTENTION "

                Case Else

            End Select

        End With

        '--- titre du message ---
        LTitreMessage.Caption = UN_ESPACE & TitreMessage
        
        '--- affectation ---
        LibelleMessage = LibelleMessage & vbCrLf
    
        '--- affichage du message ---
        For a = LCommentaires.LBound To LCommentaires.UBound

            '--- affectation ---
            PositionDuCrLf = InStr(LibelleMessage, vbCrLf)
        
            '--- recherche des valeurs ---
            Select Case PositionDuCrLf
        
                Case 0: Exit For
            
                Case 1
                    LibelleMessage = Mid(LibelleMessage, PositionDuCrLf + 2)
        
                Case Else
                    If PositionDuCrLf > 0 Then
                    
                        '--- affectation ---
                        Commentaire = Left(LibelleMessage, Pred(PositionDuCrLf))
                                       
                        '--- recherche de la présence d'un contrôle ---
                        PositionFinControle = InStr(Commentaire, "|")
                        With .LCommentaires(a)
                            If PositionFinControle > 0 Then
                        
                                '--- affectation ---
                                Controle = Left(Commentaire, Pred(PositionFinControle))
                                Commentaire = Mid(Commentaire, Succ(PositionFinControle))
                            
                                '--- décodage du contrôle ---
                                If InStr(Controle, "g") > 0 Then .Alignment = vbLeftJustify                     'alignement à gauche
                                If InStr(Controle, "c") > 0 Then .Alignment = vbCenter                            'alignement au centre
                                If InStr(Controle, "d") > 0 Then .Alignment = vbRightJustify                   'alignement à droite
                                If InStr(Controle, "m") > 0 Then Commentaire = UCase(Commentaire)  'commentaire en majuscules
                                If InStr(Controle, "s") > 0 Then .Font.Underline = True                           'soulignement
                        
                            End If
                            .Caption = Commentaire
                        End With
                    
                        '--- affectation ---
                        LibelleMessage = Mid(LibelleMessage, PositionDuCrLf + 2)
                
                    Else
                        Exit For
                    End If
    
            End Select
    
        Next a
    
        '--- types de boutons ---
        Select Case TypesBoutons
        
            Case TYPES_BOUTONS.T_OUI_NON
                '--- boutons OUI et NON ---
                .CBAnnuler.Visible = False
                .CBConfirmer.Visible = False
                
                With .CBOui
                    .Left = 1080
                    .Visible = True
                End With
                SOui.Left = .CBOui.Left
                With .CBNon
                    .Left = 4350
                    .Visible = True
                End With
                SNon.Left = .CBNon.Left
                
            Case TYPES_BOUTONS.T_OUI_NON_ANNULER
                '--- boutons OUI et NON et ANNULER ---
                .CBConfirmer.Visible = False
                
                With .CBOui
                    .Left = 770
                    .Visible = True
                End With
                SOui.Left = .CBOui.Left
                With .CBNon
                    .Left = 2700
                    .Visible = True
                End With
                SNon.Left = .CBNon.Left
                With .CBAnnuler
                    .Left = 4630
                    .Visible = True
                End With
                SAnnuler.Left = .CBAnnuler.Left
    
            Case TYPES_BOUTONS.T_CONFIRMER
                '--- bouton confirmer ---
                .CBAnnuler.Visible = False
                .CBOui.Visible = False
                .CBNon.Visible = False
                
                With .CBConfirmer
                    .Left = 2700
                    .Visible = True
                End With
                SConfirmer.Left = .CBConfirmer.Left
                SConfirmer.Visible = True
        
            Case Else
    
        End Select
    
    End With
    
End Sub


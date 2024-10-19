VERSION 5.00
Begin VB.Form FNavigateur 
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Navigateur"
   ClientHeight    =   4410
   ClientLeft      =   675
   ClientTop       =   1350
   ClientWidth     =   4920
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   Picture         =   "FNavigateur.frx":0000
   ScaleHeight     =   4410
   ScaleWidth      =   4920
   Begin VB.CommandButton CBMenu 
      BackColor       =   &H00FFC0C0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   335
      Index           =   23
      Left            =   2460
      Style           =   1  'Graphical
      TabIndex        =   23
      Top             =   4020
      Width           =   2415
   End
   Begin VB.CommandButton CBMenu 
      BackColor       =   &H00FFC0C0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   335
      Index           =   22
      Left            =   2460
      Style           =   1  'Graphical
      TabIndex        =   22
      Top             =   3660
      Width           =   2415
   End
   Begin VB.CommandButton CBMenu 
      BackColor       =   &H00FFC0C0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   335
      Index           =   21
      Left            =   2460
      Style           =   1  'Graphical
      TabIndex        =   21
      Top             =   3300
      Width           =   2415
   End
   Begin VB.CommandButton CBMenu 
      BackColor       =   &H00FFC0C0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   335
      Index           =   20
      Left            =   2460
      Style           =   1  'Graphical
      TabIndex        =   20
      Top             =   2940
      Width           =   2415
   End
   Begin VB.CommandButton CBMenu 
      BackColor       =   &H00FFC0C0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   335
      Index           =   19
      Left            =   2460
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   2580
      Width           =   2415
   End
   Begin VB.CommandButton CBMenu 
      BackColor       =   &H00FFC0C0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   335
      Index           =   18
      Left            =   2460
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   2220
      Width           =   2415
   End
   Begin VB.CommandButton CBMenu 
      BackColor       =   &H00FFC0C0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   335
      Index           =   17
      Left            =   2460
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   1860
      Width           =   2415
   End
   Begin VB.CommandButton CBMenu 
      BackColor       =   &H00FFC0C0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   335
      Index           =   16
      Left            =   2460
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   1500
      Width           =   2415
   End
   Begin VB.CommandButton CBMenu 
      BackColor       =   &H00FFC0C0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   335
      Index           =   15
      Left            =   2460
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   1140
      Width           =   2415
   End
   Begin VB.CommandButton CBMenu 
      BackColor       =   &H00FFC0C0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   335
      Index           =   14
      Left            =   2460
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   780
      Width           =   2415
   End
   Begin VB.CommandButton CBMenu 
      BackColor       =   &H00FFC0C0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   335
      Index           =   13
      Left            =   2460
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   420
      Width           =   2415
   End
   Begin VB.CommandButton CBMenu 
      BackColor       =   &H00FFC0C0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   335
      Index           =   12
      Left            =   2460
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   60
      Width           =   2415
   End
   Begin VB.CommandButton CBMenu 
      BackColor       =   &H00FFC0C0&
      Caption         =   "&Paramètres"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   335
      Index           =   11
      Left            =   60
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   4020
      Width           =   1935
   End
   Begin VB.CommandButton CBMenu 
      BackColor       =   &H00FFC0C0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   335
      Index           =   10
      Left            =   60
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   3660
      Width           =   1935
   End
   Begin VB.CommandButton CBMenu 
      BackColor       =   &H00FFC0C0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   335
      Index           =   9
      Left            =   60
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   3300
      Width           =   1935
   End
   Begin VB.CommandButton CBMenu 
      BackColor       =   &H00FFC0C0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   335
      Index           =   8
      Left            =   60
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   2940
      Width           =   1935
   End
   Begin VB.CommandButton CBMenu 
      BackColor       =   &H00FFC0C0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   335
      Index           =   7
      Left            =   60
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   2580
      Width           =   1935
   End
   Begin VB.CommandButton CBMenu 
      BackColor       =   &H00FFC0C0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   335
      Index           =   6
      Left            =   60
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   2220
      Width           =   1935
   End
   Begin VB.CommandButton CBMenu 
      BackColor       =   &H00FFC0C0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   335
      Index           =   5
      Left            =   60
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   1860
      Width           =   1935
   End
   Begin VB.CommandButton CBMenu 
      BackColor       =   &H00FFC0C0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   335
      Index           =   4
      Left            =   60
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   1500
      Width           =   1935
   End
   Begin VB.CommandButton CBMenu 
      BackColor       =   &H00FFC0C0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   335
      Index           =   3
      Left            =   60
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   1140
      Width           =   1935
   End
   Begin VB.CommandButton CBMenu 
      BackColor       =   &H00FFC0C0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   335
      Index           =   2
      Left            =   60
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   780
      Width           =   1935
   End
   Begin VB.CommandButton CBMenu 
      BackColor       =   &H00FFC0C0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   335
      Index           =   1
      Left            =   60
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   420
      Width           =   1935
   End
   Begin VB.CommandButton CBMenu 
      BackColor       =   &H00FFC0C0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   335
      Index           =   0
      Left            =   60
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   60
      Width           =   1935
   End
   Begin VB.Line LLigneSousMenu 
      BorderColor     =   &H000000FF&
      BorderWidth     =   3
      Index           =   11
      X1              =   2220
      X2              =   2460
      Y1              =   4200
      Y2              =   4200
   End
   Begin VB.Line LColonneSousMenu 
      BorderColor     =   &H000000FF&
      BorderWidth     =   3
      Index           =   10
      X1              =   2220
      X2              =   2220
      Y1              =   3840
      Y2              =   4200
   End
   Begin VB.Line LLigneMenuPrincipal 
      BorderColor     =   &H000000FF&
      BorderWidth     =   3
      Index           =   11
      X1              =   1980
      X2              =   2220
      Y1              =   4200
      Y2              =   4200
   End
   Begin VB.Line LLigneSousMenu 
      BorderColor     =   &H000000FF&
      BorderWidth     =   3
      Index           =   10
      X1              =   2220
      X2              =   2460
      Y1              =   3840
      Y2              =   3840
   End
   Begin VB.Line LColonneSousMenu 
      BorderColor     =   &H000000FF&
      BorderWidth     =   3
      Index           =   9
      X1              =   2220
      X2              =   2220
      Y1              =   3480
      Y2              =   3840
   End
   Begin VB.Line LLigneMenuPrincipal 
      BorderColor     =   &H000000FF&
      BorderWidth     =   3
      Index           =   10
      X1              =   1980
      X2              =   2220
      Y1              =   3840
      Y2              =   3840
   End
   Begin VB.Line LLigneMenuPrincipal 
      BorderColor     =   &H000000FF&
      BorderWidth     =   3
      Index           =   9
      X1              =   1980
      X2              =   2220
      Y1              =   3480
      Y2              =   3480
   End
   Begin VB.Line LLigneMenuPrincipal 
      BorderColor     =   &H000000FF&
      BorderWidth     =   3
      Index           =   8
      X1              =   1980
      X2              =   2220
      Y1              =   3120
      Y2              =   3120
   End
   Begin VB.Line LLigneMenuPrincipal 
      BorderColor     =   &H000000FF&
      BorderWidth     =   3
      Index           =   7
      X1              =   1980
      X2              =   2220
      Y1              =   2760
      Y2              =   2760
   End
   Begin VB.Line LLigneMenuPrincipal 
      BorderColor     =   &H000000FF&
      BorderWidth     =   3
      Index           =   6
      X1              =   1980
      X2              =   2220
      Y1              =   2400
      Y2              =   2400
   End
   Begin VB.Line LLigneMenuPrincipal 
      BorderColor     =   &H000000FF&
      BorderWidth     =   3
      Index           =   5
      X1              =   1980
      X2              =   2220
      Y1              =   2040
      Y2              =   2040
   End
   Begin VB.Line LLigneMenuPrincipal 
      BorderColor     =   &H000000FF&
      BorderWidth     =   3
      Index           =   4
      X1              =   1980
      X2              =   2220
      Y1              =   1680
      Y2              =   1680
   End
   Begin VB.Line LLigneMenuPrincipal 
      BorderColor     =   &H000000FF&
      BorderWidth     =   3
      Index           =   3
      X1              =   1980
      X2              =   2220
      Y1              =   1320
      Y2              =   1320
   End
   Begin VB.Line LLigneMenuPrincipal 
      BorderColor     =   &H000000FF&
      BorderWidth     =   3
      Index           =   2
      X1              =   1980
      X2              =   2220
      Y1              =   960
      Y2              =   960
   End
   Begin VB.Line LLigneMenuPrincipal 
      BorderColor     =   &H000000FF&
      BorderWidth     =   3
      Index           =   1
      X1              =   1980
      X2              =   2220
      Y1              =   600
      Y2              =   600
   End
   Begin VB.Line LLigneMenuPrincipal 
      BorderColor     =   &H000000FF&
      BorderWidth     =   3
      Index           =   0
      X1              =   1980
      X2              =   2220
      Y1              =   240
      Y2              =   240
   End
   Begin VB.Line LColonneSousMenu 
      BorderColor     =   &H000000FF&
      BorderWidth     =   3
      Index           =   8
      X1              =   2220
      X2              =   2220
      Y1              =   3120
      Y2              =   3480
   End
   Begin VB.Line LColonneSousMenu 
      BorderColor     =   &H000000FF&
      BorderWidth     =   3
      Index           =   7
      X1              =   2220
      X2              =   2220
      Y1              =   2760
      Y2              =   3120
   End
   Begin VB.Line LColonneSousMenu 
      BorderColor     =   &H000000FF&
      BorderWidth     =   3
      Index           =   6
      X1              =   2220
      X2              =   2220
      Y1              =   2400
      Y2              =   2760
   End
   Begin VB.Line LColonneSousMenu 
      BorderColor     =   &H000000FF&
      BorderWidth     =   3
      Index           =   5
      X1              =   2220
      X2              =   2220
      Y1              =   2040
      Y2              =   2400
   End
   Begin VB.Line LColonneSousMenu 
      BorderColor     =   &H000000FF&
      BorderWidth     =   3
      Index           =   4
      X1              =   2220
      X2              =   2220
      Y1              =   1680
      Y2              =   2040
   End
   Begin VB.Line LColonneSousMenu 
      BorderColor     =   &H000000FF&
      BorderWidth     =   3
      Index           =   3
      X1              =   2220
      X2              =   2220
      Y1              =   1320
      Y2              =   1680
   End
   Begin VB.Line LColonneSousMenu 
      BorderColor     =   &H000000FF&
      BorderWidth     =   3
      Index           =   2
      X1              =   2220
      X2              =   2220
      Y1              =   960
      Y2              =   1320
   End
   Begin VB.Line LColonneSousMenu 
      BorderColor     =   &H000000FF&
      BorderWidth     =   3
      Index           =   1
      X1              =   2220
      X2              =   2220
      Y1              =   600
      Y2              =   960
   End
   Begin VB.Line LColonneSousMenu 
      BorderColor     =   &H000000FF&
      BorderWidth     =   3
      Index           =   0
      X1              =   2220
      X2              =   2220
      Y1              =   240
      Y2              =   600
   End
   Begin VB.Line LLigneSousMenu 
      BorderColor     =   &H000000FF&
      BorderWidth     =   3
      Index           =   9
      X1              =   2220
      X2              =   2460
      Y1              =   3480
      Y2              =   3480
   End
   Begin VB.Line LLigneSousMenu 
      BorderColor     =   &H000000FF&
      BorderWidth     =   3
      Index           =   8
      X1              =   2220
      X2              =   2460
      Y1              =   3120
      Y2              =   3120
   End
   Begin VB.Line LLigneSousMenu 
      BorderColor     =   &H000000FF&
      BorderWidth     =   3
      Index           =   7
      X1              =   2220
      X2              =   2460
      Y1              =   2760
      Y2              =   2760
   End
   Begin VB.Line LLigneSousMenu 
      BorderColor     =   &H000000FF&
      BorderWidth     =   3
      Index           =   6
      X1              =   2220
      X2              =   2460
      Y1              =   2400
      Y2              =   2400
   End
   Begin VB.Line LLigneSousMenu 
      BorderColor     =   &H000000FF&
      BorderWidth     =   3
      Index           =   5
      X1              =   2220
      X2              =   2460
      Y1              =   2040
      Y2              =   2040
   End
   Begin VB.Line LLigneSousMenu 
      BorderColor     =   &H000000FF&
      BorderWidth     =   3
      Index           =   4
      X1              =   2220
      X2              =   2460
      Y1              =   1680
      Y2              =   1680
   End
   Begin VB.Line LLigneSousMenu 
      BorderColor     =   &H000000FF&
      BorderWidth     =   3
      Index           =   3
      X1              =   2220
      X2              =   2460
      Y1              =   1320
      Y2              =   1320
   End
   Begin VB.Line LLigneSousMenu 
      BorderColor     =   &H000000FF&
      BorderWidth     =   3
      Index           =   2
      X1              =   2220
      X2              =   2460
      Y1              =   960
      Y2              =   960
   End
   Begin VB.Line LLigneSousMenu 
      BorderColor     =   &H000000FF&
      BorderWidth     =   3
      Index           =   1
      X1              =   2220
      X2              =   2460
      Y1              =   600
      Y2              =   600
   End
   Begin VB.Line LLigneSousMenu 
      BorderColor     =   &H000000FF&
      BorderWidth     =   3
      Index           =   0
      X1              =   2220
      X2              =   2460
      Y1              =   240
      Y2              =   240
   End
End
Attribute VB_Name = "FNavigateur"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Rôle                    : Représentation graphique de l'ensemble des menus
' Nom                    : FNavigateur.frm
' Date de création : 30/03/1999
' Détails                :
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

'--- déclarations obligatoires ---
Option Explicit

'--- options générales ---
Option Base 1
DefVar A-Z

'--- constantes privées ---
Private Const DEBUT_MENU_PRINCIPAL As Integer = 0
Private Const FIN_MENU_PRINCIPAL As Integer = 11
Private Const DEBUT_SOUS_MENU As Integer = 12
Private Const FIN_SOUS_MENU As Integer = 23

Private Const LARGEUR_MINI_FEUILLE As Long = 2150
Private Const LARGEUR_MAXI_FEUILLE As Long = 5025

'--- variables privées ---
Private PremiereActivation As Boolean

'--- variables publiques ---
Public NumFeuille As Long    'numéro de la feuille lorsqu'elle devient active

'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Rôle      : Construction graphique du sous menu
' Entrées : IdxPosition -> Index de position
' Détails  :
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Private Sub ConstructionSousMenu(ByVal IdxPosition As Integer)
    
    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next
        
    '--- déclaration ---
    Dim a As Integer
    Dim NbrLignesSousMenu As Integer

    '--- affectation du nombre de lignes du sous menu ---
    NbrLignesSousMenu = CalculNbrLignesSousMenu(IdxPosition)

    '--- affichage et effacement des lignes du menu principal ---
    For a = DEBUT_MENU_PRINCIPAL To FIN_MENU_PRINCIPAL
        With LLigneMenuPrincipal(a)
            If a = IdxPosition And NbrLignesSousMenu > 0 Then
                If .Visible = False Then
                    .Visible = True
                    .Refresh
                End If
            Else
                If .Visible = True Then
                    .Visible = False
                    .Refresh
                End If
            End If
        End With
    Next a

    '--- affichage et effacement des objets concernés du sous menu ---
    For a = DEBUT_SOUS_MENU To FIN_SOUS_MENU
        
        '--- lignes du sous menu ---
        With LLigneSousMenu(a - DEBUT_SOUS_MENU)
            If (a - DEBUT_SOUS_MENU) < NbrLignesSousMenu Then
                If .Visible = False Then
                    .Visible = True
                    .Refresh
                End If
            Else
                If .Visible = True Then
                    .Visible = False
                    .Refresh
                End If
            End If
        End With
        
        '--- boutons de commande du sous menu ---
        With CBMenu(a)
            If (a - DEBUT_SOUS_MENU + 1) <= NbrLignesSousMenu Then
                If .Visible = False Then
                    .Visible = True
                    .Refresh
                End If
            Else
                If .Visible = True Then
                    .Visible = False
                    .Refresh
                End If
            End If
        End With
         
        '--- colonnes du sous menu ---
        If a - DEBUT_SOUS_MENU < (FIN_SOUS_MENU - DEBUT_SOUS_MENU) Then
            With LColonneSousMenu(a - DEBUT_SOUS_MENU)
                If (a - DEBUT_SOUS_MENU < IdxPosition And NbrLignesSousMenu > 0) Or _
                   (a - DEBUT_SOUS_MENU < Pred(NbrLignesSousMenu)) Then
                    If .Visible = False Then
                        .Visible = True
                        .Refresh
                    End If
                Else
                    If .Visible = True Then
                        .Visible = False
                        .Refresh
                    End If
                End If
            End With
        End If
        
    Next a
    
    '--- modification des dimensions de la feuille ---
    Me.WindowState = vbNormal
    Select Case Me.WindowState
        Case vbNormal
            If NbrLignesSousMenu > 0 Then
                Me.Width = LARGEUR_MAXI_FEUILLE
            Else
                Me.Width = LARGEUR_MINI_FEUILLE
            End If
            Me.Refresh
        Case Else
    End Select

End Sub

'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Rôle      : Analyse générale du menu
' Entrées : IdxPosition    -> Index de position
'                TypeFonction -> FALSE = Affichage des textes à l'intérieur des boutons du sous menu
'                                           TRUE  = Appel de la fonction concernée
' Détails  :
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Private Sub AnalyseGeneraleMenu(ByVal IdxPosition As Integer, _
                                                          ByVal TypeFonction As Boolean)
    
    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next

    '--- déclaration ---
    Dim a As Integer, _
            NbrLignesSousMenu As Integer, _
            IdxChoix As Integer
    
    '--- affectation du nombre de lignes du sous menu ---
    NbrLignesSousMenu = CalculNbrLignesSousMenu(IdxPosition)

    If TypeFonction = False Then
    
        '--- affichage des textes à l'intérieur des boutons du sous menu ---
        If NbrLignesSousMenu > 0 And _
           (IdxPosition >= DEBUT_MENU_PRINCIPAL And IdxPosition <= FIN_MENU_PRINCIPAL) Then
        
            For a = DEBUT_SOUS_MENU To DEBUT_SOUS_MENU + Pred(NbrLignesSousMenu)
            
                With CBMenu(a)
                
                    '--- affectation ---
                    IdxChoix = a - DEBUT_SOUS_MENU + 1
                
                   Select Case IdxPosition
    
                        Case 0
                            .Caption = Choose(IdxChoix, _
                                                           "Paramètres du logiciel", _
                                                           "Données prédéfinies", _
                                                           "Horaires de l'atelier", _
                                                           "Personnel", _
                                                           "Articles pour les devis", _
                                                           "Bains", _
                                                           "Jours chômés payés de l'année", _
                                                           "Tarifs clients", _
                                                           "", _
                                                           "", _
                                                           "", _
                                                           "", _
                                                           "")

                        Case 1
                            .Caption = Choose(IdxChoix, "", "", "", "", "", "", "", "", "", "", "", "")
                        
                        Case 2
                            .Caption = Choose(IdxChoix, "", "", "", "", "", "", "", "", "", "", "")
                    
                        Case 3
                            .Caption = Choose(IdxChoix, "", "", "", "", "", "", "", "", "", "", "", "")
                    
                        Case 4
                            .Caption = Choose(IdxChoix, "", "", "", "", "", "", "", "", "", "", "", "")
                    
                        Case 5
                            .Caption = Choose(IdxChoix, "Sous-traitants", "Achats", "", "", "", "", "", "", "", "", "", "")
                        
                        Case 6
                            .Caption = Choose(IdxChoix, "Personnel", "Bains", "", "", "", "", "", "", "", "", "", "")
                    
                        Case 7
                            .Caption = Choose(IdxChoix, "Ordinaires", "Spéciaux", "", "", "", "", "", "", "", "", "", "")
                    
                        Case 8
                            .Caption = Choose(IdxChoix, "Du personnel", "Bains", "", "", "", "", "", "", "", "", "", "")
                    
                        Case 9
                            .Caption = Choose(IdxChoix, "Etiquettes d'expédition", "", "", "", "", "", "", "", "", "", "", "")
                    
                        Case 10
                            .Caption = Choose(IdxChoix, "", "", "", "", "", "", "", "", "", "", "", "")
                    
                        Case 11
                            .Caption = Choose(IdxChoix, "", "", "", "", "", "", "", "", "", "", "", "")
                    
                        Case Else
                            Exit For
                
                    End Select
           
                End With
        
            Next a
    
        End If
    
    Else
    
        '--- appel de la fonction concernée ---
        Select Case MemMenuPrincipalNavigateur
        
            Case 0
                '--- menu principal sur gestion des paramètres ---
                If NbrLignesSousMenu > 0 Then
                    Select Case MemSousMenuNavigateur
                        Case 1: AppelFeuille FEUILLES.F_PARAMETRES_LOGICIEL
                        Case 2: AppelFeuille FEUILLES.F_DONNEES_PREDEFINIES
                        Case 3
                        Case 4
                        Case 5
                        Case 6
                        Case 7
                        Case 8
                        Case 9
                        Case 10
                        Case 11
                        Case Else
                    End Select
                End If
        
            Case 1
                '--- menu principal sur clients / prospects ---
                'AppelFeuille FEUILLES.F_CLIENTS_PROSPECTS
            
            Case 2
                '--- menu principal sur fournisseurs ---
                'AppelFeuille FEUILLES.F_FOURNISSEURS
            
            Case 3
                '--- menu principal sur
        
            Case 4
                '--- menu principal sur
            
            Case 5
                '--- menu principal sur
                If NbrLignesSousMenu > 0 Then
                    Select Case MemSousMenuNavigateur
                        Case 1
                        Case 2
                        Case 3
                        Case 4
                        Case 5
                        Case 6
                        Case 7
                        Case 8
                        Case 9
                        Case 10
                        Case 11
                        Case 12
                        Case Else
                    End Select
                End If
            
            Case 6
                '--- menu principal sur
                If NbrLignesSousMenu > 0 Then
                    Select Case MemSousMenuNavigateur
                        Case 1
                        Case 2
                        Case 3
                        Case 4
                        Case 5
                        Case 6
                        Case 7
                        Case 8
                        Case 9
                        Case 10
                        Case 11
                        Case 12
                        Case Else
                    End Select
                End If
            
            Case 7
                '--- menu principal sur
                If NbrLignesSousMenu > 0 Then
                    Select Case MemSousMenuNavigateur
                        Case 1
                        Case 2
                        Case 3
                        Case 4
                        Case 5
                        Case 6
                        Case 7
                        Case 8
                        Case 9
                        Case 10
                        Case 11
                        Case 12
                        Case Else
                    End Select
                End If
            
            Case 8
                '--- menu principal sur
                If NbrLignesSousMenu > 0 Then
                    Select Case MemSousMenuNavigateur
                        Case 1
                        Case 2
                        Case 3
                        Case 4
                        Case 5
                        Case 6
                        Case 7
                        Case 8
                        Case 9
                        Case 10
                        Case 11
                        Case 12
                        Case Else
                    End Select
                End If
            
            Case 9
                '--- menu principal sur divers ---
                If NbrLignesSousMenu > 0 Then
                    Select Case MemSousMenuNavigateur
                        Case 1
                        Case 2
                        Case 3
                        Case 4
                        Case 5
                        Case 6
                        Case 7
                        Case 8
                        Case 9
                        Case 10
                        Case 11
                        Case 12
                        Case Else
                    End Select
                End If
            
            Case 10
            Case 11
        
            Case Else
    
        End Select

    End If

End Sub

'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Rôle      : Calcul le nombre de lignes du sous menu en fonction de la sélection du menu principal
' Entrées : IdxPosition                           -> Index de position
' Retours : CalculNbrLignesSousMenu -> Représente le nombre de lignes du sous menu
' Détails  :
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Private Function CalculNbrLignesSousMenu(ByVal IdxPosition As Integer) As Integer
    
    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next

    '--- affectation du nombre de lignes du sous menu ---
    Select Case IdxPosition
        Case 0: CalculNbrLignesSousMenu = 8
        Case 1: CalculNbrLignesSousMenu = 0
        Case 2: CalculNbrLignesSousMenu = 0
        Case 3: CalculNbrLignesSousMenu = 0
        Case 4: CalculNbrLignesSousMenu = 0
        Case 5: CalculNbrLignesSousMenu = 2
        Case 6: CalculNbrLignesSousMenu = 2
        Case 7: CalculNbrLignesSousMenu = 2
        Case 8: CalculNbrLignesSousMenu = 2
        Case 9: CalculNbrLignesSousMenu = 1
        Case 10: CalculNbrLignesSousMenu = 0
        Case 11: CalculNbrLignesSousMenu = 0
        Case Else: CalculNbrLignesSousMenu = 0
    End Select
    
End Function

Private Sub CBMenu_Click(Index As Integer)
    
    '--- affectation ---
    If Index >= DEBUT_MENU_PRINCIPAL And Index <= FIN_MENU_PRINCIPAL Then
        MemMenuPrincipalNavigateur = Index
        MemSousMenuNavigateur = 0
    Else
        MemSousMenuNavigateur = Index - DEBUT_SOUS_MENU + 1
    End If
    
    '--- analyse générale du menu ---
    AnalyseGeneraleMenu MemMenuPrincipalNavigateur, True

End Sub

Private Sub CBMenu_GotFocus(Index As Integer)
    
    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next
    
    '--- sortie directe si la feuille doit être réduite ---
    If Me.WindowState = vbMinimized Then Exit Sub
    
    If PremiereActivation = True Then
    
        '--- changement de la couleur de fond ---
        With CBMenu(Index)
            .BackColor = COULEURS.JAUNE_2
            .Refresh
        End With
            
        '--- construction du sous menu ---
        If Index >= DEBUT_MENU_PRINCIPAL And Index <= FIN_MENU_PRINCIPAL Then
            ConstructionSousMenu Index
            AnalyseGeneraleMenu Index, False
        End If
    
    Else
        
        '--- affectation ---
        PremiereActivation = True
    
    End If

End Sub

Private Sub CBMenu_LostFocus(Index As Integer)
    
    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next
    
    '--- changement de la couleur de fond ---
    With CBMenu(Index)
        .BackColor = COULEURS.BLEU_1
        .Refresh
    End With

End Sub

Private Sub Form_Activate()
    
    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next
    
    '--- renseigne la feuille principale ---
    RenseigneFPrincipale
    
    '--- mise en place du focus mémorisé ---
    If PremiereActivation = False Then
        
        '--- construction du menu ---
        ConstructionSousMenu (MemMenuPrincipalNavigateur)
        AnalyseGeneraleMenu MemMenuPrincipalNavigateur, False
        If MemSousMenuNavigateur = 0 Then
            If MemMenuPrincipalNavigateur = 0 Then
                With CBMenu(MemMenuPrincipalNavigateur)
                    .BackColor = COULEURS.JAUNE_2
                    .Refresh
                End With
            Else
                CBMenu(MemMenuPrincipalNavigateur).SetFocus
            End If
        Else
            CBMenu(DEBUT_SOUS_MENU + MemSousMenuNavigateur - 1).SetFocus
        End If
    
    End If

End Sub

Private Sub Form_GotFocus()

    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next
       
    '--- mise en place du focus mémorisé ---
    ConstructionSousMenu (MemMenuPrincipalNavigateur)
    AnalyseGeneraleMenu MemMenuPrincipalNavigateur, False
    If MemSousMenuNavigateur = 0 Then
        If MemMenuPrincipalNavigateur = 0 Then
            With CBMenu(MemMenuPrincipalNavigateur)
                .BackColor = COULEURS.BLEU_1
                .Refresh
            End With
        Else
            CBMenu(MemMenuPrincipalNavigateur).SetFocus
        End If
    Else
        CBMenu(DEBUT_SOUS_MENU + MemSousMenuNavigateur - 1).SetFocus
    End If

End Sub

Private Sub Form_Load()
    
    '--- image de fond ---
    'Me.Picture = ImgFondDeFenetreJPG
    
    '--- affectation ---

End Sub


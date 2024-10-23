VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Begin VB.Form FRedresseurs 
   BackColor       =   &H00C0C0C0&
   Caption         =   "REDRESSEURS"
   ClientHeight    =   12990
   ClientLeft      =   1125
   ClientTop       =   480
   ClientWidth     =   19215
   BeginProperty Font 
      Name            =   "Marlett"
      Size            =   8.25
      Charset         =   2
      Weight          =   500
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FRedresseurs.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   12990
   ScaleWidth      =   19215
   Begin VB.PictureBox PBBoutons 
      Align           =   2  'Align Bottom
      DrawStyle       =   2  'Dot
      DrawWidth       =   16891
      Height          =   990
      Left            =   0
      ScaleHeight     =   930
      ScaleWidth      =   19155
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   12000
      Width           =   19215
      Begin VB.Timer TimerEtatsRedresseurs 
         Interval        =   500
         Left            =   1380
         Top             =   480
      End
      Begin VB.PictureBox PBOutilsDeplacementFenetre 
         BackColor       =   &H00E0E0E0&
         Height          =   915
         Left            =   0
         ScaleHeight     =   855
         ScaleWidth      =   1155
         TabIndex        =   10
         Top             =   0
         Visible         =   0   'False
         Width           =   1215
         Begin VB.VScrollBar VSDeplacementFenetre 
            Height          =   855
            LargeChange     =   300
            Left            =   900
            SmallChange     =   100
            TabIndex        =   13
            Top             =   0
            Width           =   255
         End
         Begin VB.HScrollBar HSDeplacementFenetre 
            Height          =   240
            LargeChange     =   300
            Left            =   0
            SmallChange     =   100
            TabIndex        =   12
            Top             =   615
            Width           =   900
         End
         Begin VB.CommandButton CBAgrandirFenetre 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Agrandir"
            DownPicture     =   "FRedresseurs.frx":014A
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   615
            Left            =   0
            MaskColor       =   &H00FF00FF&
            Picture         =   "FRedresseurs.frx":02F4
            Style           =   1  'Graphical
            TabIndex        =   11
            ToolTipText     =   " Agrandissement de la fenêtre "
            Top             =   0
            UseMaskColor    =   -1  'True
            Width           =   900
         End
      End
      Begin VB.CommandButton CBQuitter 
         Cancel          =   -1  'True
         Caption         =   "Echap=&QUITTER"
         DownPicture     =   "FRedresseurs.frx":049E
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   13440
         MaskColor       =   &H00FF00FF&
         Picture         =   "FRedresseurs.frx":0BA0
         Style           =   1  'Graphical
         TabIndex        =   61
         ToolTipText     =   " Quitter cette fenêtre "
         Top             =   120
         UseMaskColor    =   -1  'True
         Width           =   1575
      End
      Begin VB.Shape SFocus 
         BorderColor     =   &H000000FF&
         BorderWidth     =   5
         Height          =   285
         Left            =   1380
         Top             =   120
         Visible         =   0   'False
         Width           =   300
      End
   End
   Begin VB.PictureBox PBRenseignementsFenetre 
      Align           =   1  'Align Top
      BackColor       =   &H00FF0000&
      Height          =   315
      Left            =   0
      ScaleHeight     =   255
      ScaleWidth      =   19155
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   0
      Width           =   19215
      Begin VB.Label LRenseignementsFenetre 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00008000&
         BackStyle       =   0  'Transparent
         Caption         =   "REDRESSEURS"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   1440
         TabIndex        =   8
         Top             =   0
         Width           =   9315
         WordWrap        =   -1  'True
      End
   End
   Begin VB.PictureBox PBDeplacementFenetre 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   11565
      Index           =   0
      Left            =   0
      ScaleHeight     =   11565
      ScaleWidth      =   19215
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   315
      Width           =   19215
      Begin VB.PictureBox PBDeplacementFenetre 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   11205
         Index           =   1
         Left            =   0
         ScaleHeight     =   11205
         ScaleWidth      =   21450
         TabIndex        =   14
         TabStop         =   0   'False
         Top             =   0
         Width           =   21450
         Begin VB.PictureBox PBCadreLectureValeurs 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0FFFF&
            ForeColor       =   &H80000008&
            Height          =   6015
            Index           =   0
            Left            =   15360
            ScaleHeight     =   5985
            ScaleWidth      =   3525
            TabIndex        =   156
            Top             =   360
            Width           =   3555
            Begin LigneChromeURANIE5.OCXRedresseur OCXRedresseurs 
               Height          =   2865
               Index           =   1
               Left            =   180
               TabIndex        =   157
               Top             =   180
               Width           =   1785
               _ExtentX        =   3149
               _ExtentY        =   5054
               Modele          =   1
            End
            Begin VB.Label LEtatsR4 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H0080FFFF&
               BorderStyle     =   1  'Fixed Single
               Caption         =   "4"
               BeginProperty Font 
                  Name            =   "Microsoft Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   255
               Index           =   1
               Left            =   2040
               TabIndex        =   199
               Top             =   2700
               Visible         =   0   'False
               Width           =   315
            End
            Begin VB.Label LEtatsR3 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H0080FFFF&
               BorderStyle     =   1  'Fixed Single
               Caption         =   "3"
               BeginProperty Font 
                  Name            =   "Microsoft Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   255
               Index           =   1
               Left            =   2040
               TabIndex        =   194
               Top             =   2460
               Width           =   315
            End
            Begin VB.Label LEtatsR2 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H0080FFFF&
               BorderStyle     =   1  'Fixed Single
               Caption         =   "2"
               BeginProperty Font 
                  Name            =   "Microsoft Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   255
               Index           =   1
               Left            =   2040
               TabIndex        =   189
               Top             =   2220
               Width           =   315
            End
            Begin VB.Label LEtatsR1 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H0080FFFF&
               BorderStyle     =   1  'Fixed Single
               Caption         =   "1"
               BeginProperty Font 
                  Name            =   "Microsoft Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   255
               Index           =   1
               Left            =   2040
               TabIndex        =   184
               Top             =   1980
               Width           =   315
            End
            Begin VB.Label LIntensiteR4 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               BorderStyle     =   1  'Fixed Single
               BeginProperty Font 
                  Name            =   "Microsoft Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FF0000&
               Height          =   255
               Index           =   1
               Left            =   2340
               TabIndex        =   175
               Top             =   2700
               Visible         =   0   'False
               Width           =   1095
            End
            Begin VB.Label LNumCharges 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               BorderStyle     =   1  'Fixed Single
               BeginProperty Font 
                  Name            =   "Microsoft Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   255
               Index           =   5
               Left            =   2340
               TabIndex        =   174
               Top             =   660
               Width           =   855
            End
            Begin VB.Label LLibelles 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H0080C0FF&
               BorderStyle     =   1  'Fixed Single
               Caption         =   "5"
               BeginProperty Font 
                  Name            =   "Microsoft Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   255
               Index           =   9
               Left            =   2340
               TabIndex        =   173
               Top             =   420
               Width           =   855
            End
            Begin VB.Label LLibelles 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H0080C0FF&
               BorderStyle     =   1  'Fixed Single
               Caption         =   "POSTE"
               BeginProperty Font 
                  Name            =   "Microsoft Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   255
               Index           =   8
               Left            =   2340
               TabIndex        =   172
               Top             =   180
               Width           =   855
            End
            Begin VB.Label LLibelles 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H0080FFFF&
               BorderStyle     =   1  'Fixed Single
               Caption         =   "TEMPS TOTAL DU CYCLE"
               BeginProperty Font 
                  Name            =   "Microsoft Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   255
               Index           =   1
               Left            =   300
               TabIndex        =   171
               Top             =   3240
               Width           =   2955
            End
            Begin VB.Label LLibelles 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H0080FFFF&
               BorderStyle     =   1  'Fixed Single
               Caption         =   "TEMPS RESTANT DU CYCLE"
               BeginProperty Font 
                  Name            =   "Microsoft Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   255
               Index           =   2
               Left            =   300
               TabIndex        =   170
               Top             =   4140
               Width           =   2955
            End
            Begin VB.Label LLibelles 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H0080FFFF&
               BorderStyle     =   1  'Fixed Single
               Caption         =   "COUCHE EN COURS"
               BeginProperty Font 
                  Name            =   "Microsoft Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   255
               Index           =   5
               Left            =   300
               TabIndex        =   169
               Top             =   4740
               Width           =   2955
            End
            Begin VB.Label LLibelles 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H0080FFFF&
               BorderStyle     =   1  'Fixed Single
               Caption         =   "PHASE EN COURS"
               BeginProperty Font 
                  Name            =   "Microsoft Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   255
               Index           =   3
               Left            =   300
               TabIndex        =   168
               Top             =   5340
               Width           =   2955
            End
            Begin VB.Label LTempsTotalCycle 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               BorderStyle     =   1  'Fixed Single
               BeginProperty Font 
                  Name            =   "Microsoft Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FF0000&
               Height          =   255
               Index           =   1
               Left            =   300
               TabIndex        =   167
               Top             =   3480
               Width           =   2955
            End
            Begin VB.Label LTempsRestantCycle 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               BorderStyle     =   1  'Fixed Single
               BeginProperty Font 
                  Name            =   "Microsoft Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FF0000&
               Height          =   255
               Index           =   1
               Left            =   300
               TabIndex        =   166
               Top             =   4380
               Width           =   2955
            End
            Begin VB.Label LCoucheEnCours 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               BorderStyle     =   1  'Fixed Single
               BeginProperty Font 
                  Name            =   "Microsoft Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FF0000&
               Height          =   255
               Index           =   1
               Left            =   300
               TabIndex        =   165
               Top             =   4980
               Width           =   2955
            End
            Begin VB.Label LPhaseEnCours 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               BorderStyle     =   1  'Fixed Single
               BeginProperty Font 
                  Name            =   "Microsoft Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FF0000&
               Height          =   255
               Index           =   1
               Left            =   300
               TabIndex        =   164
               Top             =   5580
               Width           =   2955
            End
            Begin VB.Label LLibelles 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H0080FFFF&
               BorderStyle     =   1  'Fixed Single
               Caption         =   "TENSION"
               BeginProperty Font 
                  Name            =   "Microsoft Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   255
               Index           =   4
               Left            =   2220
               TabIndex        =   163
               Top             =   1080
               Width           =   1095
            End
            Begin VB.Label LTension 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               BorderStyle     =   1  'Fixed Single
               BeginProperty Font 
                  Name            =   "Microsoft Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FF0000&
               Height          =   255
               Index           =   1
               Left            =   2220
               TabIndex        =   162
               Top             =   1320
               Width           =   1095
            End
            Begin VB.Label LIntensiteR1 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               BorderStyle     =   1  'Fixed Single
               BeginProperty Font 
                  Name            =   "Microsoft Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FF0000&
               Height          =   255
               Index           =   1
               Left            =   2340
               TabIndex        =   161
               Top             =   1980
               Width           =   1095
            End
            Begin VB.Label LIntensiteR2 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               BorderStyle     =   1  'Fixed Single
               BeginProperty Font 
                  Name            =   "Microsoft Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FF0000&
               Height          =   255
               Index           =   1
               Left            =   2340
               TabIndex        =   160
               Top             =   2220
               Width           =   1095
            End
            Begin VB.Label LLibelles 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H0080FFFF&
               BorderStyle     =   1  'Fixed Single
               Caption         =   "I / Redresseur"
               BeginProperty Font 
                  Name            =   "Microsoft Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   255
               Index           =   63
               Left            =   2040
               TabIndex        =   159
               Top             =   1740
               Width           =   1395
            End
            Begin VB.Label LIntensiteR3 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               BorderStyle     =   1  'Fixed Single
               BeginProperty Font 
                  Name            =   "Microsoft Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FF0000&
               Height          =   255
               Index           =   1
               Left            =   2340
               TabIndex        =   158
               Top             =   2460
               Width           =   1095
            End
         End
         Begin VB.PictureBox PBCadreCommandesReseauRS485 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFC0C0&
            ForeColor       =   &H80000008&
            Height          =   1035
            Index           =   5
            Left            =   240
            ScaleHeight     =   1005
            ScaleWidth      =   3525
            TabIndex        =   136
            Top             =   9420
            Width           =   3555
            Begin VB.CommandButton CBArretRedresseur 
               BackColor       =   &H00FFFFC0&
               Caption         =   "ARRET"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Index           =   5
               Left            =   120
               MaskColor       =   &H00FFFFFF&
               Style           =   1  'Graphical
               TabIndex        =   143
               Top             =   120
               Width           =   1575
            End
            Begin VB.CommandButton CBMarcheRedresseur 
               BackColor       =   &H00FFFFC0&
               Caption         =   "MARCHE"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Index           =   5
               Left            =   1860
               Style           =   1  'Graphical
               TabIndex        =   138
               Top             =   120
               Width           =   1575
            End
            Begin VB.CommandButton CBExclusionRedresseur 
               BackColor       =   &H00FFFFC0&
               Caption         =   "EXCLUSION du redresseur"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Index           =   5
               Left            =   120
               Style           =   1  'Graphical
               TabIndex        =   137
               ToolTipText     =   " "
               Top             =   540
               Width           =   3315
            End
         End
         Begin VB.PictureBox PBCadreCommandesReseauRS485 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFC0C0&
            ForeColor       =   &H80000008&
            Height          =   1035
            Index           =   4
            Left            =   4020
            ScaleHeight     =   1005
            ScaleWidth      =   3525
            TabIndex        =   133
            Top             =   9420
            Width           =   3555
            Begin VB.CommandButton CBArretRedresseur 
               BackColor       =   &H00FFFFC0&
               Caption         =   "ARRET"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Index           =   4
               Left            =   120
               MaskColor       =   &H00FFFFFF&
               Style           =   1  'Graphical
               TabIndex        =   142
               Top             =   120
               Width           =   1575
            End
            Begin VB.CommandButton CBMarcheRedresseur 
               BackColor       =   &H00FFFFC0&
               Caption         =   "MARCHE"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Index           =   4
               Left            =   1860
               Style           =   1  'Graphical
               TabIndex        =   135
               Top             =   120
               Width           =   1575
            End
            Begin VB.CommandButton CBExclusionRedresseur 
               BackColor       =   &H00FFFFC0&
               Caption         =   "EXCLUSION du redresseur"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Index           =   4
               Left            =   120
               Style           =   1  'Graphical
               TabIndex        =   134
               Top             =   540
               Width           =   3315
            End
         End
         Begin VB.PictureBox PBCadreCommandesReseauRS485 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFC0C0&
            ForeColor       =   &H80000008&
            Height          =   1035
            Index           =   3
            Left            =   7800
            ScaleHeight     =   1005
            ScaleWidth      =   3525
            TabIndex        =   130
            Top             =   9420
            Width           =   3555
            Begin VB.CommandButton CBArretRedresseur 
               BackColor       =   &H00FFFFC0&
               Caption         =   "ARRET"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Index           =   3
               Left            =   120
               MaskColor       =   &H00FFFFFF&
               Style           =   1  'Graphical
               TabIndex        =   141
               Top             =   120
               Width           =   1575
            End
            Begin VB.CommandButton CBMarcheRedresseur 
               BackColor       =   &H00FFFFC0&
               Caption         =   "MARCHE"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Index           =   3
               Left            =   1860
               Style           =   1  'Graphical
               TabIndex        =   132
               Top             =   120
               Width           =   1575
            End
            Begin VB.CommandButton CBExclusionRedresseur 
               BackColor       =   &H00FFFFC0&
               Caption         =   "EXCLUSION du redresseur"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Index           =   3
               Left            =   120
               Style           =   1  'Graphical
               TabIndex        =   131
               Top             =   540
               Width           =   3315
            End
         End
         Begin VB.PictureBox PBCadreCommandesReseauRS485 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFC0C0&
            ForeColor       =   &H80000008&
            Height          =   1035
            Index           =   2
            Left            =   11580
            ScaleHeight     =   1005
            ScaleWidth      =   3525
            TabIndex        =   127
            Top             =   9420
            Width           =   3555
            Begin VB.CommandButton CBArretRedresseur 
               BackColor       =   &H00FFFFC0&
               Caption         =   "ARRET"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Index           =   2
               Left            =   120
               MaskColor       =   &H00FFFFFF&
               Style           =   1  'Graphical
               TabIndex        =   140
               Top             =   120
               Width           =   1575
            End
            Begin VB.CommandButton CBMarcheRedresseur 
               BackColor       =   &H00FFFFC0&
               Caption         =   "MARCHE"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Index           =   2
               Left            =   1860
               Style           =   1  'Graphical
               TabIndex        =   129
               Top             =   120
               Width           =   1575
            End
            Begin VB.CommandButton CBExclusionRedresseur 
               BackColor       =   &H00FFFFC0&
               Caption         =   "EXCLUSION du redresseur"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Index           =   2
               Left            =   120
               Style           =   1  'Graphical
               TabIndex        =   128
               Top             =   540
               Width           =   3315
            End
         End
         Begin VB.PictureBox PBCadreCommandesReseauRS485 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFC0C0&
            ForeColor       =   &H80000008&
            Height          =   1035
            Index           =   1
            Left            =   15360
            ScaleHeight     =   1005
            ScaleWidth      =   3525
            TabIndex        =   124
            Top             =   9420
            Width           =   3555
            Begin VB.CommandButton CBArretRedresseur 
               BackColor       =   &H00FFFFC0&
               Caption         =   "ARRET"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Index           =   1
               Left            =   120
               MaskColor       =   &H00FFFFFF&
               Style           =   1  'Graphical
               TabIndex        =   139
               Top             =   120
               Width           =   1575
            End
            Begin VB.CommandButton CBMarcheRedresseur 
               BackColor       =   &H00FFFFC0&
               Caption         =   "MARCHE"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Index           =   1
               Left            =   1860
               MaskColor       =   &H00FFFFFF&
               Style           =   1  'Graphical
               TabIndex        =   126
               Top             =   120
               Width           =   1575
            End
            Begin VB.CommandButton CBExclusionRedresseur 
               BackColor       =   &H00FFFFC0&
               Caption         =   "EXCLUSION du redresseur"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Index           =   1
               Left            =   120
               Style           =   1  'Graphical
               TabIndex        =   125
               Top             =   540
               Width           =   3315
            End
         End
         Begin VB.PictureBox PBCadreModificationValeurs 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0E0FF&
            ForeColor       =   &H80000008&
            Height          =   1755
            Index           =   5
            Left            =   240
            ScaleHeight     =   1725
            ScaleWidth      =   3525
            TabIndex        =   108
            Top             =   7020
            Width           =   3555
            Begin VB.CommandButton CBSensAjoutTempsDeBain 
               BackColor       =   &H00FFFF80&
               Caption         =   "EN +"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   345
               Index           =   5
               Left            =   1680
               Style           =   1  'Graphical
               TabIndex        =   212
               Top             =   150
               Width           =   615
            End
            Begin VB.CommandButton CBTransfertModificationsVersAPI 
               BackColor       =   &H00C0C0FF&
               Caption         =   "Transférer"
               Enabled         =   0   'False
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Index           =   5
               Left            =   1920
               Style           =   1  'Graphical
               TabIndex        =   110
               ToolTipText     =   " Transfère les valeurs dans l'automate "
               Top             =   1260
               Width           =   1515
            End
            Begin VB.CommandButton CBAnnulerTransfertModifications 
               BackColor       =   &H00C0FFFF&
               Caption         =   "Annuler"
               Enabled         =   0   'False
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Index           =   5
               Left            =   120
               Style           =   1  'Graphical
               TabIndex        =   109
               ToolTipText     =   " Annule l'entrée des données "
               Top             =   1260
               Width           =   1515
            End
            Begin MSMask.MaskEdBox MEBTempsTotalCycle 
               Height          =   315
               Index           =   5
               Left            =   2400
               TabIndex        =   111
               Top             =   180
               Width           =   975
               _ExtentX        =   1720
               _ExtentY        =   556
               _Version        =   393216
               ForeColor       =   16711680
               MaxLength       =   5
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Mask            =   "##:##"
               PromptChar      =   "_"
            End
            Begin MSMask.MaskEdBox MEBIntensite 
               Height          =   315
               Index           =   5
               Left            =   2400
               TabIndex        =   112
               Top             =   660
               Width           =   975
               _ExtentX        =   1720
               _ExtentY        =   556
               _Version        =   393216
               ForeColor       =   16711680
               MaxLength       =   5
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Mask            =   "#####"
               PromptChar      =   "_"
            End
            Begin VB.Label LLibelles 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BackColor       =   &H0080C0FF&
               BackStyle       =   0  'Transparent
               Caption         =   "Nouvelle intensité (A)"
               BeginProperty Font 
                  Name            =   "Microsoft Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   240
               Index           =   39
               Left            =   60
               TabIndex        =   114
               Top             =   720
               Width           =   1905
            End
            Begin VB.Label LLibelles 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H0080C0FF&
               BackStyle       =   0  'Transparent
               Caption         =   "Ajout au temps de bain (mm:ss)"
               BeginProperty Font 
                  Name            =   "Microsoft Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   420
               Index           =   38
               Left            =   60
               TabIndex        =   113
               Top             =   120
               Width           =   1485
            End
            Begin VB.Shape SDecoration 
               BackColor       =   &H000040C0&
               BackStyle       =   1  'Opaque
               Height          =   615
               Index           =   5
               Left            =   -300
               Top             =   1140
               Width           =   4035
            End
         End
         Begin VB.PictureBox PBCadreModificationValeurs 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0E0FF&
            ForeColor       =   &H80000008&
            Height          =   1755
            Index           =   4
            Left            =   4020
            ScaleHeight     =   1725
            ScaleWidth      =   3525
            TabIndex        =   102
            Top             =   7020
            Width           =   3555
            Begin VB.CommandButton CBSensAjoutTempsDeBain 
               BackColor       =   &H00FFFF80&
               Caption         =   "EN +"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   345
               Index           =   4
               Left            =   1680
               Style           =   1  'Graphical
               TabIndex        =   211
               Top             =   150
               Width           =   615
            End
            Begin VB.CommandButton CBAnnulerTransfertModifications 
               BackColor       =   &H00C0FFFF&
               Caption         =   "Annuler"
               Enabled         =   0   'False
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Index           =   4
               Left            =   120
               Style           =   1  'Graphical
               TabIndex        =   104
               ToolTipText     =   " Annule l'entrée des données "
               Top             =   1260
               Width           =   1515
            End
            Begin VB.CommandButton CBTransfertModificationsVersAPI 
               BackColor       =   &H00C0C0FF&
               Caption         =   "Transférer"
               Enabled         =   0   'False
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Index           =   4
               Left            =   1920
               Style           =   1  'Graphical
               TabIndex        =   103
               ToolTipText     =   " Transfère les valeurs dans l'automate "
               Top             =   1260
               Width           =   1515
            End
            Begin MSMask.MaskEdBox MEBTempsTotalCycle 
               Height          =   315
               Index           =   4
               Left            =   2400
               TabIndex        =   105
               Top             =   180
               Width           =   975
               _ExtentX        =   1720
               _ExtentY        =   556
               _Version        =   393216
               ForeColor       =   16711680
               MaxLength       =   5
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Mask            =   "##:##"
               PromptChar      =   "_"
            End
            Begin MSMask.MaskEdBox MEBIntensite 
               Height          =   315
               Index           =   4
               Left            =   2400
               TabIndex        =   106
               Top             =   660
               Width           =   975
               _ExtentX        =   1720
               _ExtentY        =   556
               _Version        =   393216
               ForeColor       =   16711680
               MaxLength       =   5
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Mask            =   "#####"
               PromptChar      =   "_"
            End
            Begin VB.Label LLibelles 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H0080C0FF&
               BackStyle       =   0  'Transparent
               Caption         =   "Ajout au temps de bain (mm:ss)"
               BeginProperty Font 
                  Name            =   "Microsoft Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   420
               Index           =   12
               Left            =   60
               TabIndex        =   204
               Top             =   120
               Width           =   1485
            End
            Begin VB.Shape SDecoration 
               BackColor       =   &H000040C0&
               BackStyle       =   1  'Opaque
               Height          =   615
               Index           =   4
               Left            =   -300
               Top             =   1140
               Width           =   4035
            End
            Begin VB.Label LLibelles 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BackColor       =   &H0080C0FF&
               BackStyle       =   0  'Transparent
               Caption         =   "Nouvelle intensité (A)"
               BeginProperty Font 
                  Name            =   "Microsoft Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   240
               Index           =   36
               Left            =   60
               TabIndex        =   107
               Top             =   720
               Width           =   1905
            End
         End
         Begin VB.PictureBox PBCadreLectureValeurs 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0FFC0&
            ForeColor       =   &H80000008&
            Height          =   6015
            Index           =   5
            Left            =   240
            ScaleHeight     =   5985
            ScaleWidth      =   3525
            TabIndex        =   80
            Top             =   360
            Width           =   3555
            Begin LigneChromeURANIE5.OCXRedresseur OCXRedresseurs 
               Height          =   2865
               Index           =   5
               Left            =   180
               TabIndex        =   101
               Top             =   180
               Width           =   1785
               _ExtentX        =   3149
               _ExtentY        =   5054
               Modele          =   1
            End
            Begin VB.Label LTempsAjouteSurIFaible 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               BorderStyle     =   1  'Fixed Single
               BeginProperty Font 
                  Name            =   "Microsoft Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FF0000&
               Height          =   255
               Index           =   5
               Left            =   300
               TabIndex        =   216
               Top             =   3720
               Width           =   2955
            End
            Begin VB.Label LEtatsR4 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H0080FF80&
               BorderStyle     =   1  'Fixed Single
               Caption         =   "4"
               BeginProperty Font 
                  Name            =   "Microsoft Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   255
               Index           =   5
               Left            =   2040
               TabIndex        =   203
               Top             =   2700
               Width           =   315
            End
            Begin VB.Label LEtatsR3 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H0080FF80&
               BorderStyle     =   1  'Fixed Single
               Caption         =   "3"
               BeginProperty Font 
                  Name            =   "Microsoft Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   255
               Index           =   5
               Left            =   2040
               TabIndex        =   198
               Top             =   2460
               Width           =   315
            End
            Begin VB.Label LEtatsR2 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H0080FF80&
               BorderStyle     =   1  'Fixed Single
               Caption         =   "2"
               BeginProperty Font 
                  Name            =   "Microsoft Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   255
               Index           =   5
               Left            =   2040
               TabIndex        =   193
               Top             =   2220
               Width           =   315
            End
            Begin VB.Label LEtatsR1 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H0080FF80&
               BorderStyle     =   1  'Fixed Single
               Caption         =   "1"
               BeginProperty Font 
                  Name            =   "Microsoft Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   255
               Index           =   5
               Left            =   2040
               TabIndex        =   188
               Top             =   1980
               Width           =   315
            End
            Begin VB.Label LLibelles 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H0080FF80&
               BorderStyle     =   1  'Fixed Single
               Caption         =   "I / Redresseur"
               BeginProperty Font 
                  Name            =   "Microsoft Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   255
               Index           =   67
               Left            =   2040
               TabIndex        =   183
               Top             =   1740
               Width           =   1395
            End
            Begin VB.Label LIntensiteR4 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               BorderStyle     =   1  'Fixed Single
               BeginProperty Font 
                  Name            =   "Microsoft Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FF0000&
               Height          =   255
               Index           =   5
               Left            =   2340
               TabIndex        =   179
               Top             =   2700
               Width           =   1095
            End
            Begin VB.Label LIntensiteR3 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               BorderStyle     =   1  'Fixed Single
               BeginProperty Font 
                  Name            =   "Microsoft Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FF0000&
               Height          =   255
               Index           =   5
               Left            =   2340
               TabIndex        =   155
               Top             =   2460
               Width           =   1095
            End
            Begin VB.Label LIntensiteR2 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               BorderStyle     =   1  'Fixed Single
               BeginProperty Font 
                  Name            =   "Microsoft Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FF0000&
               Height          =   255
               Index           =   5
               Left            =   2340
               TabIndex        =   151
               Top             =   2220
               Width           =   1095
            End
            Begin VB.Label LIntensiteR1 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               BorderStyle     =   1  'Fixed Single
               BeginProperty Font 
                  Name            =   "Microsoft Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FF0000&
               Height          =   255
               Index           =   5
               Left            =   2340
               TabIndex        =   147
               Top             =   1980
               Width           =   1095
            End
            Begin VB.Label LTension 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               BorderStyle     =   1  'Fixed Single
               BeginProperty Font 
                  Name            =   "Microsoft Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FF0000&
               Height          =   255
               Index           =   5
               Left            =   2220
               TabIndex        =   95
               Top             =   1320
               Width           =   1095
            End
            Begin VB.Label LLibelles 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H0080FF80&
               BorderStyle     =   1  'Fixed Single
               Caption         =   "TENSION"
               BeginProperty Font 
                  Name            =   "Microsoft Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   255
               Index           =   56
               Left            =   2220
               TabIndex        =   94
               Top             =   1080
               Width           =   1095
            End
            Begin VB.Label LPhaseEnCours 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               BorderStyle     =   1  'Fixed Single
               BeginProperty Font 
                  Name            =   "Microsoft Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FF0000&
               Height          =   255
               Index           =   5
               Left            =   300
               TabIndex        =   93
               Top             =   5580
               Width           =   2955
            End
            Begin VB.Label LCoucheEnCours 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               BorderStyle     =   1  'Fixed Single
               BeginProperty Font 
                  Name            =   "Microsoft Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FF0000&
               Height          =   255
               Index           =   5
               Left            =   300
               TabIndex        =   92
               Top             =   4980
               Width           =   2955
            End
            Begin VB.Label LTempsRestantCycle 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               BorderStyle     =   1  'Fixed Single
               BeginProperty Font 
                  Name            =   "Microsoft Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FF0000&
               Height          =   255
               Index           =   5
               Left            =   300
               TabIndex        =   91
               Top             =   4380
               Width           =   2955
            End
            Begin VB.Label LTempsTotalCycle 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               BorderStyle     =   1  'Fixed Single
               BeginProperty Font 
                  Name            =   "Microsoft Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FF0000&
               Height          =   255
               Index           =   5
               Left            =   300
               TabIndex        =   90
               Top             =   3480
               Width           =   2955
            End
            Begin VB.Label LLibelles 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H0080FF80&
               BorderStyle     =   1  'Fixed Single
               Caption         =   "PHASE EN COURS"
               BeginProperty Font 
                  Name            =   "Microsoft Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   255
               Index           =   55
               Left            =   300
               TabIndex        =   89
               Top             =   5340
               Width           =   2955
            End
            Begin VB.Label LLibelles 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H0080FF80&
               BorderStyle     =   1  'Fixed Single
               Caption         =   "COUCHE EN COURS"
               BeginProperty Font 
                  Name            =   "Microsoft Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   255
               Index           =   54
               Left            =   300
               TabIndex        =   88
               Top             =   4740
               Width           =   2955
            End
            Begin VB.Label LLibelles 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H0080FF80&
               BorderStyle     =   1  'Fixed Single
               Caption         =   "TEMPS RESTANT DU CYCLE"
               BeginProperty Font 
                  Name            =   "Microsoft Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   255
               Index           =   53
               Left            =   300
               TabIndex        =   87
               Top             =   4140
               Width           =   2955
            End
            Begin VB.Label LLibelles 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H0080FF80&
               BorderStyle     =   1  'Fixed Single
               Caption         =   "TEMPS TOTAL DU CYCLE"
               BeginProperty Font 
                  Name            =   "Microsoft Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   255
               Index           =   52
               Left            =   300
               TabIndex        =   86
               Top             =   3240
               Width           =   2955
            End
            Begin VB.Label LLibelles 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H0080C0FF&
               BorderStyle     =   1  'Fixed Single
               Caption         =   "POSTES"
               BeginProperty Font 
                  Name            =   "Microsoft Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   255
               Index           =   51
               Left            =   2220
               TabIndex        =   85
               Top             =   180
               Width           =   1095
            End
            Begin VB.Label LLibelles 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H0080C0FF&
               BorderStyle     =   1  'Fixed Single
               Caption         =   "15"
               BeginProperty Font 
                  Name            =   "Microsoft Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   255
               Index           =   50
               Left            =   2220
               TabIndex        =   84
               Top             =   420
               Width           =   555
            End
            Begin VB.Label LLibelles 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H0080C0FF&
               BorderStyle     =   1  'Fixed Single
               Caption         =   "16"
               BeginProperty Font 
                  Name            =   "Microsoft Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   255
               Index           =   49
               Left            =   2760
               TabIndex        =   83
               Top             =   420
               Width           =   555
            End
            Begin VB.Label LNumCharges 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               BorderStyle     =   1  'Fixed Single
               BeginProperty Font 
                  Name            =   "Microsoft Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   255
               Index           =   15
               Left            =   2220
               TabIndex        =   82
               Top             =   660
               Width           =   555
            End
            Begin VB.Label LNumCharges 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               BorderStyle     =   1  'Fixed Single
               BeginProperty Font 
                  Name            =   "Microsoft Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   255
               Index           =   16
               Left            =   2760
               TabIndex        =   81
               Top             =   660
               Width           =   555
            End
         End
         Begin VB.PictureBox PBCadreLectureValeurs 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0FFC0&
            ForeColor       =   &H80000008&
            Height          =   6015
            Index           =   4
            Left            =   4020
            ScaleHeight     =   5985
            ScaleWidth      =   3525
            TabIndex        =   64
            Top             =   360
            Width           =   3555
            Begin LigneChromeURANIE5.OCXRedresseur OCXRedresseurs 
               Height          =   2865
               Index           =   4
               Left            =   180
               TabIndex        =   100
               Top             =   180
               Width           =   1785
               _ExtentX        =   3149
               _ExtentY        =   5054
               Modele          =   1
            End
            Begin VB.Label LTempsAjouteSurIFaible 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               BorderStyle     =   1  'Fixed Single
               BeginProperty Font 
                  Name            =   "Microsoft Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FF0000&
               Height          =   255
               Index           =   4
               Left            =   300
               TabIndex        =   215
               Top             =   3720
               Width           =   2955
            End
            Begin VB.Label LEtatsR4 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H0080FF80&
               BorderStyle     =   1  'Fixed Single
               Caption         =   "4"
               BeginProperty Font 
                  Name            =   "Microsoft Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   255
               Index           =   4
               Left            =   2040
               TabIndex        =   202
               Top             =   2700
               Width           =   315
            End
            Begin VB.Label LEtatsR3 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H0080FF80&
               BorderStyle     =   1  'Fixed Single
               Caption         =   "3"
               BeginProperty Font 
                  Name            =   "Microsoft Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   255
               Index           =   4
               Left            =   2040
               TabIndex        =   197
               Top             =   2460
               Width           =   315
            End
            Begin VB.Label LEtatsR2 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H0080FF80&
               BorderStyle     =   1  'Fixed Single
               Caption         =   "2"
               BeginProperty Font 
                  Name            =   "Microsoft Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   255
               Index           =   4
               Left            =   2040
               TabIndex        =   192
               Top             =   2220
               Width           =   315
            End
            Begin VB.Label LEtatsR1 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H0080FF80&
               BorderStyle     =   1  'Fixed Single
               Caption         =   "1"
               BeginProperty Font 
                  Name            =   "Microsoft Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   255
               Index           =   4
               Left            =   2040
               TabIndex        =   187
               Top             =   1980
               Width           =   315
            End
            Begin VB.Label LLibelles 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H0080FF80&
               BorderStyle     =   1  'Fixed Single
               Caption         =   "I / Redresseur"
               BeginProperty Font 
                  Name            =   "Microsoft Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   255
               Index           =   66
               Left            =   2040
               TabIndex        =   182
               Top             =   1740
               Width           =   1395
            End
            Begin VB.Label LIntensiteR4 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               BorderStyle     =   1  'Fixed Single
               BeginProperty Font 
                  Name            =   "Microsoft Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FF0000&
               Height          =   255
               Index           =   4
               Left            =   2340
               TabIndex        =   178
               Top             =   2700
               Width           =   1095
            End
            Begin VB.Label LIntensiteR3 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               BorderStyle     =   1  'Fixed Single
               BeginProperty Font 
                  Name            =   "Microsoft Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FF0000&
               Height          =   255
               Index           =   4
               Left            =   2340
               TabIndex        =   154
               Top             =   2460
               Width           =   1095
            End
            Begin VB.Label LIntensiteR2 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               BorderStyle     =   1  'Fixed Single
               BeginProperty Font 
                  Name            =   "Microsoft Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FF0000&
               Height          =   255
               Index           =   4
               Left            =   2340
               TabIndex        =   150
               Top             =   2220
               Width           =   1095
            End
            Begin VB.Label LIntensiteR1 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               BorderStyle     =   1  'Fixed Single
               BeginProperty Font 
                  Name            =   "Microsoft Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FF0000&
               Height          =   255
               Index           =   4
               Left            =   2340
               TabIndex        =   146
               Top             =   1980
               Width           =   1095
            End
            Begin VB.Label LTension 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               BorderStyle     =   1  'Fixed Single
               BeginProperty Font 
                  Name            =   "Microsoft Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FF0000&
               Height          =   255
               Index           =   4
               Left            =   2220
               TabIndex        =   79
               Top             =   1320
               Width           =   1095
            End
            Begin VB.Label LLibelles 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H0080FF80&
               BorderStyle     =   1  'Fixed Single
               Caption         =   "TENSION"
               BeginProperty Font 
                  Name            =   "Microsoft Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   255
               Index           =   48
               Left            =   2220
               TabIndex        =   78
               Top             =   1080
               Width           =   1095
            End
            Begin VB.Label LPhaseEnCours 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               BorderStyle     =   1  'Fixed Single
               BeginProperty Font 
                  Name            =   "Microsoft Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FF0000&
               Height          =   255
               Index           =   4
               Left            =   300
               TabIndex        =   77
               Top             =   5580
               Width           =   2955
            End
            Begin VB.Label LCoucheEnCours 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               BorderStyle     =   1  'Fixed Single
               BeginProperty Font 
                  Name            =   "Microsoft Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FF0000&
               Height          =   255
               Index           =   4
               Left            =   300
               TabIndex        =   76
               Top             =   4980
               Width           =   2955
            End
            Begin VB.Label LTempsRestantCycle 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               BorderStyle     =   1  'Fixed Single
               BeginProperty Font 
                  Name            =   "Microsoft Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FF0000&
               Height          =   255
               Index           =   4
               Left            =   300
               TabIndex        =   75
               Top             =   4380
               Width           =   2955
            End
            Begin VB.Label LTempsTotalCycle 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               BorderStyle     =   1  'Fixed Single
               BeginProperty Font 
                  Name            =   "Microsoft Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FF0000&
               Height          =   255
               Index           =   4
               Left            =   300
               TabIndex        =   74
               Top             =   3480
               Width           =   2955
            End
            Begin VB.Label LLibelles 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H0080FF80&
               BorderStyle     =   1  'Fixed Single
               Caption         =   "PHASE EN COURS"
               BeginProperty Font 
                  Name            =   "Microsoft Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   255
               Index           =   47
               Left            =   300
               TabIndex        =   73
               Top             =   5340
               Width           =   2955
            End
            Begin VB.Label LLibelles 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H0080FF80&
               BorderStyle     =   1  'Fixed Single
               Caption         =   "COUCHE EN COURS"
               BeginProperty Font 
                  Name            =   "Microsoft Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   255
               Index           =   46
               Left            =   300
               TabIndex        =   72
               Top             =   4740
               Width           =   2955
            End
            Begin VB.Label LLibelles 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H0080FF80&
               BorderStyle     =   1  'Fixed Single
               Caption         =   "TEMPS RESTANT DU CYCLE"
               BeginProperty Font 
                  Name            =   "Microsoft Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   255
               Index           =   45
               Left            =   300
               TabIndex        =   71
               Top             =   4140
               Width           =   2955
            End
            Begin VB.Label LLibelles 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H0080FF80&
               BorderStyle     =   1  'Fixed Single
               Caption         =   "TEMPS TOTAL DU CYCLE"
               BeginProperty Font 
                  Name            =   "Microsoft Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   255
               Index           =   44
               Left            =   300
               TabIndex        =   70
               Top             =   3240
               Width           =   2955
            End
            Begin VB.Label LLibelles 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H0080C0FF&
               BorderStyle     =   1  'Fixed Single
               Caption         =   "POSTES"
               BeginProperty Font 
                  Name            =   "Microsoft Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   255
               Index           =   43
               Left            =   2220
               TabIndex        =   69
               Top             =   180
               Width           =   1095
            End
            Begin VB.Label LLibelles 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H0080C0FF&
               BorderStyle     =   1  'Fixed Single
               Caption         =   "13"
               BeginProperty Font 
                  Name            =   "Microsoft Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   255
               Index           =   42
               Left            =   2220
               TabIndex        =   68
               Top             =   420
               Width           =   555
            End
            Begin VB.Label LLibelles 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H0080C0FF&
               BorderStyle     =   1  'Fixed Single
               Caption         =   "14"
               BeginProperty Font 
                  Name            =   "Microsoft Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   255
               Index           =   41
               Left            =   2760
               TabIndex        =   67
               Top             =   420
               Width           =   555
            End
            Begin VB.Label LNumCharges 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               BorderStyle     =   1  'Fixed Single
               BeginProperty Font 
                  Name            =   "Microsoft Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   255
               Index           =   13
               Left            =   2220
               TabIndex        =   66
               Top             =   660
               Width           =   555
            End
            Begin VB.Label LNumCharges 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               BorderStyle     =   1  'Fixed Single
               BeginProperty Font 
                  Name            =   "Microsoft Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   255
               Index           =   14
               Left            =   2760
               TabIndex        =   65
               Top             =   660
               Width           =   555
            End
         End
         Begin VB.PictureBox PBCadreLectureValeurs 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0FFC0&
            ForeColor       =   &H80000008&
            Height          =   6015
            Index           =   2
            Left            =   11580
            ScaleHeight     =   5985
            ScaleWidth      =   3525
            TabIndex        =   41
            Top             =   360
            Width           =   3555
            Begin LigneChromeURANIE5.OCXRedresseur OCXRedresseurs 
               Height          =   2865
               Index           =   2
               Left            =   180
               TabIndex        =   98
               Top             =   180
               Width           =   1785
               _ExtentX        =   3149
               _ExtentY        =   5054
               Modele          =   1
            End
            Begin VB.Label LTempsAjouteSurIFaible 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               BorderStyle     =   1  'Fixed Single
               BeginProperty Font 
                  Name            =   "Microsoft Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FF0000&
               Height          =   255
               Index           =   2
               Left            =   300
               TabIndex        =   213
               Top             =   3720
               Width           =   2955
            End
            Begin VB.Label LEtatsR4 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H0080FF80&
               BorderStyle     =   1  'Fixed Single
               Caption         =   "4"
               BeginProperty Font 
                  Name            =   "Microsoft Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   255
               Index           =   2
               Left            =   2040
               TabIndex        =   200
               Top             =   2700
               Width           =   315
            End
            Begin VB.Label LEtatsR3 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H0080FF80&
               BorderStyle     =   1  'Fixed Single
               Caption         =   "3"
               BeginProperty Font 
                  Name            =   "Microsoft Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   255
               Index           =   2
               Left            =   2040
               TabIndex        =   195
               Top             =   2460
               Width           =   315
            End
            Begin VB.Label LEtatsR2 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H0080FF80&
               BorderStyle     =   1  'Fixed Single
               Caption         =   "2"
               BeginProperty Font 
                  Name            =   "Microsoft Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   255
               Index           =   2
               Left            =   2040
               TabIndex        =   190
               Top             =   2220
               Width           =   315
            End
            Begin VB.Label LEtatsR1 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H0080FF80&
               BorderStyle     =   1  'Fixed Single
               Caption         =   "1"
               BeginProperty Font 
                  Name            =   "Microsoft Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   255
               Index           =   2
               Left            =   2040
               TabIndex        =   185
               Top             =   1980
               Width           =   315
            End
            Begin VB.Label LLibelles 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H0080FF80&
               BorderStyle     =   1  'Fixed Single
               Caption         =   "I / Redresseur"
               BeginProperty Font 
                  Name            =   "Microsoft Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   255
               Index           =   64
               Left            =   2040
               TabIndex        =   180
               Top             =   1740
               Width           =   1395
            End
            Begin VB.Label LIntensiteR4 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               BorderStyle     =   1  'Fixed Single
               BeginProperty Font 
                  Name            =   "Microsoft Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FF0000&
               Height          =   255
               Index           =   2
               Left            =   2340
               TabIndex        =   176
               Top             =   2700
               Width           =   1095
            End
            Begin VB.Label LIntensiteR3 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               BorderStyle     =   1  'Fixed Single
               BeginProperty Font 
                  Name            =   "Microsoft Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FF0000&
               Height          =   255
               Index           =   2
               Left            =   2340
               TabIndex        =   152
               Top             =   2460
               Width           =   1095
            End
            Begin VB.Label LIntensiteR2 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               BorderStyle     =   1  'Fixed Single
               BeginProperty Font 
                  Name            =   "Microsoft Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FF0000&
               Height          =   255
               Index           =   2
               Left            =   2340
               TabIndex        =   148
               Top             =   2220
               Width           =   1095
            End
            Begin VB.Label LIntensiteR1 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               BorderStyle     =   1  'Fixed Single
               BeginProperty Font 
                  Name            =   "Microsoft Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FF0000&
               Height          =   255
               Index           =   2
               Left            =   2340
               TabIndex        =   144
               Top             =   1980
               Width           =   1095
            End
            Begin VB.Label LNumCharges 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               BorderStyle     =   1  'Fixed Single
               BeginProperty Font 
                  Name            =   "Microsoft Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   255
               Index           =   7
               Left            =   2760
               TabIndex        =   56
               Top             =   660
               Width           =   555
            End
            Begin VB.Label LNumCharges 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               BorderStyle     =   1  'Fixed Single
               BeginProperty Font 
                  Name            =   "Microsoft Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   255
               Index           =   6
               Left            =   2220
               TabIndex        =   55
               Top             =   660
               Width           =   555
            End
            Begin VB.Label LLibelles 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H0080C0FF&
               BorderStyle     =   1  'Fixed Single
               Caption         =   "7"
               BeginProperty Font 
                  Name            =   "Microsoft Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   255
               Index           =   19
               Left            =   2760
               TabIndex        =   54
               Top             =   420
               Width           =   555
            End
            Begin VB.Label LLibelles 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H0080C0FF&
               BorderStyle     =   1  'Fixed Single
               Caption         =   "6"
               BeginProperty Font 
                  Name            =   "Microsoft Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   255
               Index           =   20
               Left            =   2220
               TabIndex        =   53
               Top             =   420
               Width           =   555
            End
            Begin VB.Label LLibelles 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H0080C0FF&
               BorderStyle     =   1  'Fixed Single
               Caption         =   "POSTES"
               BeginProperty Font 
                  Name            =   "Microsoft Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   255
               Index           =   21
               Left            =   2220
               TabIndex        =   52
               Top             =   180
               Width           =   1095
            End
            Begin VB.Label LLibelles 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H0080FF80&
               BorderStyle     =   1  'Fixed Single
               Caption         =   "TEMPS TOTAL DU CYCLE"
               BeginProperty Font 
                  Name            =   "Microsoft Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   255
               Index           =   22
               Left            =   300
               TabIndex        =   51
               Top             =   3240
               Width           =   2955
            End
            Begin VB.Label LLibelles 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H0080FF80&
               BorderStyle     =   1  'Fixed Single
               Caption         =   "TEMPS RESTANT DU CYCLE"
               BeginProperty Font 
                  Name            =   "Microsoft Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   255
               Index           =   23
               Left            =   300
               TabIndex        =   50
               Top             =   4140
               Width           =   2955
            End
            Begin VB.Label LLibelles 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H0080FF80&
               BorderStyle     =   1  'Fixed Single
               Caption         =   "COUCHE EN COURS"
               BeginProperty Font 
                  Name            =   "Microsoft Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   255
               Index           =   24
               Left            =   300
               TabIndex        =   49
               Top             =   4740
               Width           =   2955
            End
            Begin VB.Label LLibelles 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H0080FF80&
               BorderStyle     =   1  'Fixed Single
               Caption         =   "PHASE EN COURS"
               BeginProperty Font 
                  Name            =   "Microsoft Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   255
               Index           =   25
               Left            =   300
               TabIndex        =   48
               Top             =   5340
               Width           =   2955
            End
            Begin VB.Label LTempsTotalCycle 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               BorderStyle     =   1  'Fixed Single
               BeginProperty Font 
                  Name            =   "Microsoft Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FF0000&
               Height          =   255
               Index           =   2
               Left            =   300
               TabIndex        =   47
               Top             =   3480
               Width           =   2955
            End
            Begin VB.Label LTempsRestantCycle 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               BorderStyle     =   1  'Fixed Single
               BeginProperty Font 
                  Name            =   "Microsoft Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FF0000&
               Height          =   255
               Index           =   2
               Left            =   300
               TabIndex        =   46
               Top             =   4380
               Width           =   2955
            End
            Begin VB.Label LCoucheEnCours 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               BorderStyle     =   1  'Fixed Single
               BeginProperty Font 
                  Name            =   "Microsoft Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FF0000&
               Height          =   255
               Index           =   2
               Left            =   300
               TabIndex        =   45
               Top             =   4980
               Width           =   2955
            End
            Begin VB.Label LPhaseEnCours 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               BorderStyle     =   1  'Fixed Single
               BeginProperty Font 
                  Name            =   "Microsoft Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FF0000&
               Height          =   255
               Index           =   2
               Left            =   300
               TabIndex        =   44
               Top             =   5580
               Width           =   2955
            End
            Begin VB.Label LLibelles 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H0080FF80&
               BorderStyle     =   1  'Fixed Single
               Caption         =   "TENSION"
               BeginProperty Font 
                  Name            =   "Microsoft Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   255
               Index           =   26
               Left            =   2220
               TabIndex        =   43
               Top             =   1080
               Width           =   1095
            End
            Begin VB.Label LTension 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               BorderStyle     =   1  'Fixed Single
               BeginProperty Font 
                  Name            =   "Microsoft Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FF0000&
               Height          =   255
               Index           =   2
               Left            =   2220
               TabIndex        =   42
               Top             =   1320
               Width           =   1095
            End
         End
         Begin VB.PictureBox PBCadreLectureValeurs 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0FFC0&
            ForeColor       =   &H80000008&
            Height          =   6015
            Index           =   3
            Left            =   7800
            ScaleHeight     =   5985
            ScaleWidth      =   3525
            TabIndex        =   25
            Top             =   360
            Width           =   3555
            Begin LigneChromeURANIE5.OCXRedresseur OCXRedresseurs 
               Height          =   2865
               Index           =   3
               Left            =   180
               TabIndex        =   99
               Top             =   180
               Width           =   1785
               _ExtentX        =   3149
               _ExtentY        =   5054
               Modele          =   1
            End
            Begin VB.Label LTempsAjouteSurIFaible 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               BorderStyle     =   1  'Fixed Single
               BeginProperty Font 
                  Name            =   "Microsoft Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FF0000&
               Height          =   255
               Index           =   3
               Left            =   300
               TabIndex        =   214
               Top             =   3720
               Width           =   2955
            End
            Begin VB.Label LEtatsR4 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H0080FF80&
               BorderStyle     =   1  'Fixed Single
               Caption         =   "4"
               BeginProperty Font 
                  Name            =   "Microsoft Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   255
               Index           =   3
               Left            =   2040
               TabIndex        =   201
               Top             =   2700
               Width           =   315
            End
            Begin VB.Label LEtatsR3 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H0080FF80&
               BorderStyle     =   1  'Fixed Single
               Caption         =   "3"
               BeginProperty Font 
                  Name            =   "Microsoft Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   255
               Index           =   3
               Left            =   2040
               TabIndex        =   196
               Top             =   2460
               Width           =   315
            End
            Begin VB.Label LEtatsR2 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H0080FF80&
               BorderStyle     =   1  'Fixed Single
               Caption         =   "2"
               BeginProperty Font 
                  Name            =   "Microsoft Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   255
               Index           =   3
               Left            =   2040
               TabIndex        =   191
               Top             =   2220
               Width           =   315
            End
            Begin VB.Label LEtatsR1 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H0080FF80&
               BorderStyle     =   1  'Fixed Single
               Caption         =   "1"
               BeginProperty Font 
                  Name            =   "Microsoft Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   255
               Index           =   3
               Left            =   2040
               TabIndex        =   186
               Top             =   1980
               Width           =   315
            End
            Begin VB.Label LLibelles 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H0080FF80&
               BorderStyle     =   1  'Fixed Single
               Caption         =   "I / Redresseur"
               BeginProperty Font 
                  Name            =   "Microsoft Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   255
               Index           =   65
               Left            =   2040
               TabIndex        =   181
               Top             =   1740
               Width           =   1395
            End
            Begin VB.Label LIntensiteR4 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               BorderStyle     =   1  'Fixed Single
               BeginProperty Font 
                  Name            =   "Microsoft Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FF0000&
               Height          =   255
               Index           =   3
               Left            =   2340
               TabIndex        =   177
               Top             =   2700
               Width           =   1095
            End
            Begin VB.Label LIntensiteR3 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               BorderStyle     =   1  'Fixed Single
               BeginProperty Font 
                  Name            =   "Microsoft Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FF0000&
               Height          =   255
               Index           =   3
               Left            =   2340
               TabIndex        =   153
               Top             =   2460
               Width           =   1095
            End
            Begin VB.Label LIntensiteR2 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               BorderStyle     =   1  'Fixed Single
               BeginProperty Font 
                  Name            =   "Microsoft Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FF0000&
               Height          =   255
               Index           =   3
               Left            =   2340
               TabIndex        =   149
               Top             =   2220
               Width           =   1095
            End
            Begin VB.Label LIntensiteR1 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               BorderStyle     =   1  'Fixed Single
               BeginProperty Font 
                  Name            =   "Microsoft Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FF0000&
               Height          =   255
               Index           =   3
               Left            =   2340
               TabIndex        =   145
               Top             =   1980
               Width           =   1095
            End
            Begin VB.Label LNumCharges 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               BorderStyle     =   1  'Fixed Single
               BeginProperty Font 
                  Name            =   "Microsoft Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   255
               Index           =   9
               Left            =   2760
               TabIndex        =   40
               Top             =   660
               Width           =   555
            End
            Begin VB.Label LNumCharges 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               BorderStyle     =   1  'Fixed Single
               BeginProperty Font 
                  Name            =   "Microsoft Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   255
               Index           =   8
               Left            =   2220
               TabIndex        =   39
               Top             =   660
               Width           =   555
            End
            Begin VB.Label LLibelles 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H0080C0FF&
               BorderStyle     =   1  'Fixed Single
               Caption         =   "9"
               BeginProperty Font 
                  Name            =   "Microsoft Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   255
               Index           =   27
               Left            =   2760
               TabIndex        =   38
               Top             =   420
               Width           =   555
            End
            Begin VB.Label LLibelles 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H0080C0FF&
               BorderStyle     =   1  'Fixed Single
               Caption         =   "8"
               BeginProperty Font 
                  Name            =   "Microsoft Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   255
               Index           =   28
               Left            =   2220
               TabIndex        =   37
               Top             =   420
               Width           =   555
            End
            Begin VB.Label LLibelles 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H0080C0FF&
               BorderStyle     =   1  'Fixed Single
               Caption         =   "POSTES"
               BeginProperty Font 
                  Name            =   "Microsoft Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   255
               Index           =   29
               Left            =   2220
               TabIndex        =   36
               Top             =   180
               Width           =   1095
            End
            Begin VB.Label LLibelles 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H0080FF80&
               BorderStyle     =   1  'Fixed Single
               Caption         =   "TEMPS TOTAL DU CYCLE"
               BeginProperty Font 
                  Name            =   "Microsoft Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   255
               Index           =   30
               Left            =   300
               TabIndex        =   35
               Top             =   3240
               Width           =   2955
            End
            Begin VB.Label LLibelles 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H0080FF80&
               BorderStyle     =   1  'Fixed Single
               Caption         =   "TEMPS RESTANT DU CYCLE"
               BeginProperty Font 
                  Name            =   "Microsoft Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   255
               Index           =   31
               Left            =   300
               TabIndex        =   34
               Top             =   4140
               Width           =   2955
            End
            Begin VB.Label LLibelles 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H0080FF80&
               BorderStyle     =   1  'Fixed Single
               Caption         =   "COUCHE EN COURS"
               BeginProperty Font 
                  Name            =   "Microsoft Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   255
               Index           =   32
               Left            =   300
               TabIndex        =   33
               Top             =   4740
               Width           =   2955
            End
            Begin VB.Label LLibelles 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H0080FF80&
               BorderStyle     =   1  'Fixed Single
               Caption         =   "PHASE EN COURS"
               BeginProperty Font 
                  Name            =   "Microsoft Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   255
               Index           =   33
               Left            =   300
               TabIndex        =   32
               Top             =   5340
               Width           =   2955
            End
            Begin VB.Label LTempsTotalCycle 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               BorderStyle     =   1  'Fixed Single
               BeginProperty Font 
                  Name            =   "Microsoft Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FF0000&
               Height          =   255
               Index           =   3
               Left            =   300
               TabIndex        =   31
               Top             =   3480
               Width           =   2955
            End
            Begin VB.Label LTempsRestantCycle 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               BorderStyle     =   1  'Fixed Single
               BeginProperty Font 
                  Name            =   "Microsoft Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FF0000&
               Height          =   255
               Index           =   3
               Left            =   300
               TabIndex        =   30
               Top             =   4380
               Width           =   2955
            End
            Begin VB.Label LCoucheEnCours 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               BorderStyle     =   1  'Fixed Single
               BeginProperty Font 
                  Name            =   "Microsoft Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FF0000&
               Height          =   255
               Index           =   3
               Left            =   300
               TabIndex        =   29
               Top             =   4980
               Width           =   2955
            End
            Begin VB.Label LPhaseEnCours 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               BorderStyle     =   1  'Fixed Single
               BeginProperty Font 
                  Name            =   "Microsoft Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FF0000&
               Height          =   255
               Index           =   3
               Left            =   300
               TabIndex        =   28
               Top             =   5580
               Width           =   2955
            End
            Begin VB.Label LLibelles 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H0080FF80&
               BorderStyle     =   1  'Fixed Single
               Caption         =   "TENSION"
               BeginProperty Font 
                  Name            =   "Microsoft Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   255
               Index           =   34
               Left            =   2220
               TabIndex        =   27
               Top             =   1080
               Width           =   1095
            End
            Begin VB.Label LTension 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               BorderStyle     =   1  'Fixed Single
               BeginProperty Font 
                  Name            =   "Microsoft Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FF0000&
               Height          =   255
               Index           =   3
               Left            =   2220
               TabIndex        =   26
               Top             =   1320
               Width           =   1095
            End
         End
         Begin VB.PictureBox PBCadreModificationValeurs 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0E0FF&
            ForeColor       =   &H80000008&
            Height          =   1755
            Index           =   1
            Left            =   15360
            ScaleHeight     =   1725
            ScaleWidth      =   3525
            TabIndex        =   21
            Top             =   7020
            Width           =   3555
            Begin VB.CommandButton CBSensAjoutTempsDeBain 
               BackColor       =   &H00FFFF80&
               Caption         =   "EN +"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   345
               Index           =   1
               Left            =   1680
               Style           =   1  'Graphical
               TabIndex        =   208
               Top             =   150
               Width           =   615
            End
            Begin VB.CommandButton CBTransfertModificationsVersAPI 
               BackColor       =   &H00C0C0FF&
               Caption         =   "Transférer"
               Enabled         =   0   'False
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Index           =   1
               Left            =   1920
               Style           =   1  'Graphical
               TabIndex        =   23
               ToolTipText     =   " Transfère les valeurs dans l'automate "
               Top             =   1260
               Width           =   1515
            End
            Begin VB.CommandButton CBAnnulerTransfertModifications 
               BackColor       =   &H00C0FFFF&
               Caption         =   "Annuler"
               Enabled         =   0   'False
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Index           =   1
               Left            =   120
               Style           =   1  'Graphical
               TabIndex        =   22
               ToolTipText     =   " Annule l'entrée des données "
               Top             =   1260
               Width           =   1515
            End
            Begin MSMask.MaskEdBox MEBTempsTotalCycle 
               Height          =   315
               Index           =   1
               Left            =   2400
               TabIndex        =   0
               Top             =   180
               Width           =   975
               _ExtentX        =   1720
               _ExtentY        =   556
               _Version        =   393216
               ForeColor       =   16711680
               MaxLength       =   5
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Mask            =   "##:##"
               PromptChar      =   "_"
            End
            Begin MSMask.MaskEdBox MEBIntensite 
               Height          =   315
               Index           =   1
               Left            =   2400
               TabIndex        =   1
               Top             =   660
               Width           =   975
               _ExtentX        =   1720
               _ExtentY        =   556
               _Version        =   393216
               ForeColor       =   16711680
               PromptInclude   =   0   'False
               MaxLength       =   5
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Mask            =   "#####"
               PromptChar      =   "_"
            End
            Begin VB.Label LLibelles 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H0080C0FF&
               BackStyle       =   0  'Transparent
               Caption         =   "Ajout au temps de bain (mm:ss)"
               BeginProperty Font 
                  Name            =   "Microsoft Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   420
               Index           =   37
               Left            =   60
               TabIndex        =   207
               Top             =   120
               Width           =   1485
            End
            Begin VB.Shape SDecoration 
               BackColor       =   &H000040C0&
               BackStyle       =   1  'Opaque
               Height          =   615
               Index           =   1
               Left            =   -300
               Top             =   1140
               Width           =   4035
            End
            Begin VB.Label LLibelles 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BackColor       =   &H0080C0FF&
               BackStyle       =   0  'Transparent
               Caption         =   "Nouvelle intensité (A)"
               BeginProperty Font 
                  Name            =   "Microsoft Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   240
               Index           =   11
               Left            =   60
               TabIndex        =   24
               Top             =   720
               Width           =   1905
            End
         End
         Begin VB.PictureBox PBCadreModificationValeurs 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0E0FF&
            ForeColor       =   &H80000008&
            Height          =   1755
            Index           =   2
            Left            =   11580
            ScaleHeight     =   1725
            ScaleWidth      =   3525
            TabIndex        =   18
            Top             =   7020
            Width           =   3555
            Begin VB.CommandButton CBSensAjoutTempsDeBain 
               BackColor       =   &H00FFFF80&
               Caption         =   "EN +"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   345
               Index           =   2
               Left            =   1680
               Style           =   1  'Graphical
               TabIndex        =   209
               Top             =   150
               Width           =   615
            End
            Begin VB.CommandButton CBTransfertModificationsVersAPI 
               BackColor       =   &H00C0C0FF&
               Caption         =   "Transférer"
               Enabled         =   0   'False
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Index           =   2
               Left            =   1920
               Style           =   1  'Graphical
               TabIndex        =   20
               ToolTipText     =   " Transfère les valeurs dans l'automate "
               Top             =   1260
               Width           =   1515
            End
            Begin VB.CommandButton CBAnnulerTransfertModifications 
               BackColor       =   &H00C0FFFF&
               Caption         =   "Annuler"
               Enabled         =   0   'False
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Index           =   2
               Left            =   120
               Style           =   1  'Graphical
               TabIndex        =   19
               ToolTipText     =   " Annule l'entrée des données "
               Top             =   1260
               Width           =   1515
            End
            Begin MSMask.MaskEdBox MEBTempsTotalCycle 
               Height          =   315
               Index           =   2
               Left            =   2400
               TabIndex        =   2
               Top             =   180
               Width           =   975
               _ExtentX        =   1720
               _ExtentY        =   556
               _Version        =   393216
               ForeColor       =   16711680
               MaxLength       =   5
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Mask            =   "##:##"
               PromptChar      =   "_"
            End
            Begin MSMask.MaskEdBox MEBIntensite 
               Height          =   315
               Index           =   2
               Left            =   2400
               TabIndex        =   3
               Top             =   660
               Width           =   975
               _ExtentX        =   1720
               _ExtentY        =   556
               _Version        =   393216
               ForeColor       =   16711680
               MaxLength       =   5
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Mask            =   "#####"
               PromptChar      =   "_"
            End
            Begin VB.Label LLibelles 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H0080C0FF&
               BackStyle       =   0  'Transparent
               Caption         =   "Ajout au temps de bain (mm:ss)"
               BeginProperty Font 
                  Name            =   "Microsoft Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   420
               Index           =   14
               Left            =   60
               TabIndex        =   206
               Top             =   120
               Width           =   1485
            End
            Begin VB.Label LLibelles 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BackColor       =   &H0080C0FF&
               BackStyle       =   0  'Transparent
               Caption         =   "Nouvelle intensité (A)"
               BeginProperty Font 
                  Name            =   "Microsoft Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   240
               Index           =   15
               Left            =   60
               TabIndex        =   62
               Top             =   720
               Width           =   1905
            End
            Begin VB.Shape SDecoration 
               BackColor       =   &H000040C0&
               BackStyle       =   1  'Opaque
               Height          =   615
               Index           =   2
               Left            =   -300
               Top             =   1140
               Width           =   4035
            End
         End
         Begin VB.PictureBox PBCadreModificationValeurs 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0E0FF&
            ForeColor       =   &H80000008&
            Height          =   1755
            Index           =   3
            Left            =   7800
            ScaleHeight     =   1725
            ScaleWidth      =   3525
            TabIndex        =   15
            Top             =   7020
            Width           =   3555
            Begin VB.CommandButton CBSensAjoutTempsDeBain 
               BackColor       =   &H00FFFF80&
               Caption         =   "EN +"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   345
               Index           =   3
               Left            =   1680
               Style           =   1  'Graphical
               TabIndex        =   210
               Top             =   150
               Width           =   615
            End
            Begin VB.CommandButton CBTransfertModificationsVersAPI 
               BackColor       =   &H00C0C0FF&
               Caption         =   "Transférer"
               Enabled         =   0   'False
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Index           =   3
               Left            =   1920
               Style           =   1  'Graphical
               TabIndex        =   17
               ToolTipText     =   " Transfère les valeurs dans l'automate "
               Top             =   1260
               Width           =   1515
            End
            Begin VB.CommandButton CBAnnulerTransfertModifications 
               BackColor       =   &H00C0FFFF&
               Caption         =   "Annuler"
               Enabled         =   0   'False
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Index           =   3
               Left            =   120
               Style           =   1  'Graphical
               TabIndex        =   16
               ToolTipText     =   " Annule l'entrée des données "
               Top             =   1260
               Width           =   1515
            End
            Begin MSMask.MaskEdBox MEBTempsTotalCycle 
               Height          =   315
               Index           =   3
               Left            =   2400
               TabIndex        =   4
               Top             =   180
               Width           =   975
               _ExtentX        =   1720
               _ExtentY        =   556
               _Version        =   393216
               ForeColor       =   16711680
               MaxLength       =   5
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Mask            =   "##:##"
               PromptChar      =   "_"
            End
            Begin MSMask.MaskEdBox MEBIntensite 
               Height          =   315
               Index           =   3
               Left            =   2400
               TabIndex        =   5
               Top             =   660
               Width           =   975
               _ExtentX        =   1720
               _ExtentY        =   556
               _Version        =   393216
               ForeColor       =   16711680
               MaxLength       =   5
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Mask            =   "#####"
               PromptChar      =   "_"
            End
            Begin VB.Label LLibelles 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H0080C0FF&
               BackStyle       =   0  'Transparent
               Caption         =   "Ajout au temps de bain (mm:ss)"
               BeginProperty Font 
                  Name            =   "Microsoft Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   420
               Index           =   13
               Left            =   60
               TabIndex        =   205
               Top             =   120
               Width           =   1485
            End
            Begin VB.Label LLibelles 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BackColor       =   &H0080C0FF&
               BackStyle       =   0  'Transparent
               Caption         =   "Nouvelle intensité (A)"
               BeginProperty Font 
                  Name            =   "Microsoft Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   240
               Index           =   16
               Left            =   60
               TabIndex        =   63
               Top             =   720
               Width           =   1905
            End
            Begin VB.Shape SDecoration 
               BackColor       =   &H000040C0&
               BackStyle       =   1  'Opaque
               Height          =   615
               Index           =   3
               Left            =   -300
               Top             =   1140
               Width           =   4035
            End
         End
         Begin VB.Label LLibelles 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FF0000&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "COMMANDES du REDRESSEUR du CHROME 4"
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
            Height          =   555
            Index           =   62
            Left            =   240
            TabIndex        =   123
            Top             =   8880
            Width           =   3555
         End
         Begin VB.Label LLibelles 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FF0000&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "COMMANDES du REDRESSEUR du CHROME 3"
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
            Height          =   555
            Index           =   61
            Left            =   4020
            TabIndex        =   122
            Top             =   8880
            Width           =   3555
         End
         Begin VB.Label LLibelles 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FF0000&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "COMMANDES du REDRESSEUR du CHROME 2"
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
            Height          =   555
            Index           =   60
            Left            =   7800
            TabIndex        =   121
            Top             =   8880
            Width           =   3555
         End
         Begin VB.Label LLibelles 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FF0000&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "COMMANDES du REDRESSEUR du CHROME 1"
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
            Height          =   555
            Index           =   59
            Left            =   11580
            TabIndex        =   120
            Top             =   8880
            Width           =   3555
         End
         Begin VB.Label LLibelles 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FF0000&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "COMMANDES du REDRESSEUR de l'ATTAQUE ANODIQUE"
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
            Height          =   555
            Index           =   10
            Left            =   15360
            TabIndex        =   119
            Top             =   8880
            Width           =   3555
         End
         Begin VB.Label LLibelles 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H000040C0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "MODIFICATION du REDRESSEUR du CHROME 4"
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
            Height          =   555
            Index           =   58
            Left            =   240
            TabIndex        =   118
            Top             =   6480
            Width           =   3555
         End
         Begin VB.Label LLibelles 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H000040C0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "MODIFICATION du REDRESSEUR du CHROME 3"
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
            Height          =   555
            Index           =   40
            Left            =   4020
            TabIndex        =   117
            Top             =   6480
            Width           =   3555
         End
         Begin VB.Label LLibelles 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H000040C0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "MODIFICATION du REDRESSEUR du CHROME 2"
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
            Height          =   555
            Index           =   35
            Left            =   7800
            TabIndex        =   116
            Top             =   6480
            Width           =   3555
         End
         Begin VB.Label LLibelles 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H000040C0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "MODIFICATION du REDRESSEUR du CHROME 1"
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
            Height          =   555
            Index           =   18
            Left            =   11580
            TabIndex        =   115
            Top             =   6480
            Width           =   3555
         End
         Begin VB.Label LLibelles 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00004000&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "REDRESSEUR D'ATTAQUE"
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
            Index           =   17
            Left            =   15360
            TabIndex        =   97
            Top             =   120
            Width           =   3555
         End
         Begin VB.Label LLibelles 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00004000&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "REDRESSEUR CHROME 4"
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
            Index           =   57
            Left            =   240
            TabIndex        =   96
            Top             =   120
            Width           =   3555
         End
         Begin VB.Label LLibelles 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00004000&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "REDRESSEUR CHROME 1"
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
            Left            =   11580
            TabIndex        =   60
            Top             =   120
            Width           =   3555
         End
         Begin VB.Label LLibelles 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H000040C0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "MODIFICATION du REDRESSEUR de l'ATTAQUE ANODIQUE"
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
            Height          =   555
            Index           =   100
            Left            =   15360
            TabIndex        =   59
            Top             =   6480
            Width           =   3555
         End
         Begin VB.Label LLibelles 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00004000&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "REDRESSEUR CHROME 2"
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
            Index           =   6
            Left            =   7800
            TabIndex        =   58
            Top             =   120
            Width           =   3555
         End
         Begin VB.Label LLibelles 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00004000&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "REDRESSEUR CHROME 3"
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
            Index           =   7
            Left            =   4020
            TabIndex        =   57
            Top             =   120
            Width           =   3555
         End
      End
   End
End
Attribute VB_Name = "FRedresseurs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Rôle                    : Fenêtre gérant les redresseurs
' Nom                    : FRedresseurs.frm
' Date de création : 11/01/2005
' Détails                :
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

'--- déclarations obligatoires ---
Option Explicit

'--- options générales ---
Option Base 1
DefVar A-Z

'--- constantes privées ---
Private Const TITRE_FENETRE As String = "REDRESSEURS"
Private Const TITRE_MESSAGES As String = TITRE_FENETRE

Private Const TEXTE_EN_PLUS As String = "EN +"
Private Const TEXTE_EN_MOINS As String = "EN -"

'--- énumérations privées ---

'--- types privées ---

'--- variables privées ---
Private PremiereActivation As Boolean
Private InterdireEvenements As Boolean        'pour interdire certains évènements

'--- variables publiques ---
Public NumFenetre As Long                             'numéro de la fenêtre lorsqu'elle devient active
Public RedresseurEnCours As Integer             'redresseur en cours

'--- tableaux privés ---

'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Rôle      : Initialise la fenêtre (chargement ou en vue de la rendre visible)
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Sub InitialisationFenetre()
    
    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next

    '--- déclaration ---

    '--- affectation ---

    '--- divers sur la Fenetre ---
    With Me
        .Caption = UCase(TITRE_FENETRE)
        .WindowState = vbMaximized
    End With
    PBDeplacementFenetre(ZONES_DEPLACEMENT_Fenetre.Z_FILLE).Picture = ImgFondDeFenetreXP
    
    '--- renseignements de la fenêtre ---
    LRenseignementsFenetre.Caption = UCase(TITRE_FENETRE)
    
End Sub

'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Rôle      : Décharge la Fenetre
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Private Sub DechargeFenetre()
    
    '--- aiguillage en cas d'erreur ---
    On Error Resume Next
    
    '--- affectation ---
    PremiereActivation = False

    '--- curseur souris par défaut ---
    SourisEnAttente False

    '--- neutralisation du timer ---
    With TimerEtatsRedresseurs
        .Enabled = False
        .Interval = 0
    End With
    
    '--- déchargement de la fenêtre ---
    Me.Visible = False
    Unload Me
    Set OccFRedresseurs = Nothing
    
End Sub

'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Rôle      : Change le curseur de la souris en fonction de l'attente
' Entrées : AttenteOuiNon -> TRUE   = Curseur en forme de sablier
'                                            FALSE = Curseur par défaut
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
' Rôle      : Effectue le paramètrage de la Fenetre
' Entrées :
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Sub ParametrageFenetre()
    
    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next

    '--- visualisation des différents états des redresseurs ---
    EtatsRedresseurs

    '--- lancement du timer ---
    TimerEtatsRedresseurs.Enabled = True
                
End Sub

'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Rôle      : Visualisation des différents états des redresseurs
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Private Sub EtatsRedresseurs()
    
    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next
    
    '--- déclaration ---
    
    '--- affichage de la totalité des charges ---
    AffichageTotaliteCharges
    
    '--- affichage de la totalité des données des redresseurs ---
    AffichageDonneesRedresseurs
    
End Sub

Private Sub CBAgrandirFENETRE_Click()
    On Error Resume Next
    Me.WindowState = vbMaximized
End Sub

Private Sub CBAnnulerTransfertModifications_Click(Index As Integer)

    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next

    '--- interdire les évènements ---
    InterdireEvenements = True

    '--- vidage des champs ---
    MEBTempsTotalCycle(Index).Mask = ""
    MEBTempsTotalCycle(Index).Text = ""
    MEBTempsTotalCycle(Index).Mask = "##:##"
    
    MEBIntensite(Index).Mask = ""
    MEBIntensite(Index).Text = ""
    MEBIntensite(Index).Mask = "#####"

    '--- autoriser les évènements ---
    InterdireEvenements = False

    '--- gestion des boutons ---
    CBAnnulerTransfertModifications(Index).Enabled = False
    CBTransfertModificationsVersAPI(Index).Enabled = False

End Sub

Private Sub CBArretRedresseur_Click(Index As Integer)
    
    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next
    
    '--- constantes privées ---
    Const NOM_GROUPE As String = "REDRESSEURS"
    Const CODE_ARRET_REDRESSEUR As String = "10"
    
    '--- déclaration ---
    Dim ValeurRetourneeAPI As Long          'valeur retournée par une fonction concernant le dialogue avec l'automate
    Dim NomVariable As String                    'nom de la variable OPC
    
    If AppelFenetre(F_MESSAGE, _
                            TITRE_MESSAGES, _
                             MESSAGE_10, _
                            TYPES_MESSAGES.T_ATTENTION, _
                            TYPES_BOUTONS.T_OUI_NON, _
                            EMPLACEMENT_FOCUS.E_SUR_NON) = vbYes Then

        '--- transfert des valeurs ---
        If PROGRAMME_AVEC_AUTOMATE = True Then
        
            '--- curseur de la souris ---
            SourisEnAttente True
                    
            '--- affectation du nom de la variable ---
            NomVariable = Choose(Index, "DemandesDuPCRA", "DemandesDuPCR1", "DemandesDuPCR2", "DemandesDuPCR3", "DemandesDuPCR4")
                    
            '--- écriture dans l'automate ---
            ValeurRetourneeAPI = APIEcritureVariableNommee(NOM_GROUPE, NomVariable, CODE_ARRET_REDRESSEUR)
            If ValeurRetourneeAPI <> 0 Then
                Bidon = MessageErreur(TITRE_MESSAGES, MESSAGE_500)             'lancer un message d'erreur
            End If
                        
            '--- curseur de la souris ---
            SourisEnAttente False
    
        End If
        
    End If

End Sub

Private Sub CBExclusionRedresseur_Click(Index As Integer)
    
    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next
    
    '--- constantes privées ---
    Const NOM_GROUPE As String = "REDRESSEURS"
    Const CODE_EXCLUSION_REDRESSEUR As String = "14"
    
    '--- déclaration ---
    Dim ValeurRetourneeAPI As Long          'valeur retournée par une fonction concernant le dialogue avec l'automate
    Dim NomVariable As String                    'nom de la variable OPC
    
    If AppelFenetre(F_MESSAGE, _
                            TITRE_MESSAGES, _
                             MESSAGE_12, _
                            TYPES_MESSAGES.T_ATTENTION, _
                            TYPES_BOUTONS.T_OUI_NON, _
                            EMPLACEMENT_FOCUS.E_SUR_NON) = vbYes Then

        '--- transfert des valeurs ---
        If PROGRAMME_AVEC_AUTOMATE = True Then
        
            '--- curseur de la souris ---
            SourisEnAttente True
                    
            '--- affectation du nom de la variable ---
            NomVariable = Choose(Index, "DemandesDuPCRA", "DemandesDuPCR1", "DemandesDuPCR2", "DemandesDuPCR3", "DemandesDuPCR4")
                    
            '--- écriture dans l'automate ---
            ValeurRetourneeAPI = APIEcritureVariableNommee(NOM_GROUPE, NomVariable, CODE_EXCLUSION_REDRESSEUR)
            If ValeurRetourneeAPI <> 0 Then
                Bidon = MessageErreur(TITRE_MESSAGES, MESSAGE_500)             'lancer un message d'erreur
            End If
                        
            '--- curseur de la souris ---
            SourisEnAttente False
    
        End If
        
    End If

End Sub

Private Sub CBMarcheRedresseur_Click(Index As Integer)
    
    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next
    
    '--- constantes privées ---
    Const NOM_GROUPE As String = "REDRESSEURS"
    Const CODE_MARCHE_REDRESSEUR As String = "11"
    
    '--- déclaration ---
    Dim ValeurRetourneeAPI As Long          'valeur retournée par une fonction concernant le dialogue avec l'automate
    Dim NomVariable As String                    'nom de la variable OPC
    
    If AppelFenetre(F_MESSAGE, _
                            TITRE_MESSAGES, _
                             MESSAGE_11, _
                            TYPES_MESSAGES.T_ATTENTION, _
                            TYPES_BOUTONS.T_OUI_NON, _
                            EMPLACEMENT_FOCUS.E_SUR_NON) = vbYes Then

        '--- transfert des valeurs ---
        If PROGRAMME_AVEC_AUTOMATE = True Then
        
            '--- curseur de la souris ---
            SourisEnAttente True
                    
            '--- affectation du nom de la variable ---
            NomVariable = Choose(Index, "DemandesDuPCRA", "DemandesDuPCR1", "DemandesDuPCR2", "DemandesDuPCR3", "DemandesDuPCR4")
                    
            '--- écriture dans l'automate ---
            ValeurRetourneeAPI = APIEcritureVariableNommee(NOM_GROUPE, NomVariable, CODE_MARCHE_REDRESSEUR)
            If ValeurRetourneeAPI <> 0 Then
                Bidon = MessageErreur(TITRE_MESSAGES, MESSAGE_500)             'lancer un message d'erreur
            End If
                        
            '--- curseur de la souris ---
            SourisEnAttente False
    
        End If
        
    End If

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

Private Sub CBSensAjoutTempsDeBain_Click(Index As Integer)

    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next

    '--- constantes privées ---
    
    '--- déclaration ---

    With CBSensAjoutTempsDeBain(Index)
        If .Caption = TEXTE_EN_PLUS Then
            .Caption = TEXTE_EN_MOINS
            .BackColor = COULEURS.ROUGE_1
        Else
            .Caption = TEXTE_EN_PLUS
            .BackColor = COULEURS.CYAN_2
        End If
    End With

End Sub

Private Sub CBTransfertModificationsVersAPI_Click(Index As Integer)

    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next

    '--- constantes privées ---
    
    '--- déclaration ---
    Dim NumCharge As Integer                                           'numéro de charge
    Dim NbrTotalCouches As Integer                                  'nombre total de couches
    Dim AjoutTempsTotalCycleSecondes As Integer           'représente l'ajout au temps total en secondes
    Dim NouveauTempsCoucheTSecondes As Integer       'représente le nouveau temps total de la couche T en secondes
    Dim NouveauTempsTotalChromageAFaire As Integer   'représente le nouveau temps total de chromage à faire
    
    Dim NouvelleIntensite As Long                                     'représente la nouvelle intensité
    Dim ValeurRetourneeAPI As Long                                 'valeur retournée par une fonction concernant le dialogue avec l'automate
    
    Dim NouvelleIntensiteTexte As String                           'représente la nouvelle intensité en format texte
    Dim AjoutTempsTotalCycleTexte As String                   'représente l'ajout de temps total en format texte
    Dim NomGroupe As String                                             'représente un nom de groupe
    Dim NomElement As String                                            'représente un nom d'élément (variable nommée)
    
    '--- affectation du numéro de charge ---
    NumCharge = TEtatsRedresseurs(Index).NumCharge

    '--- analyse en fonction du numéro de charge ---
    If NumCharge >= CHARGES.C_NUM_MINI And NumCharge <= CHARGES.C_NUM_MAXI Then

        '--- calcul de la valeur du temps pour la dernière couche ---
        With TEtatsCharges(NumCharge)

            '--- affectation du nombre total de couches à faire ---
            NbrTotalCouches = TEtatsRedresseurs(Index).NbrTotalCouches
            
            If NbrTotalCouches >= 1 Then
                
                '--- affectation du nouveau temps en secondes tapé dans le champ d'édition ---
                AjoutTempsTotalCycleTexte = MEBTempsTotalCycle(Index).Text
                AjoutTempsTotalCycleTexte = Replace(AjoutTempsTotalCycleTexte, "_", "0")
                If AjoutTempsTotalCycleTexte = "" Then AjoutTempsTotalCycleTexte = "00:00"
                
                '--- affectation en numérique / affectation du signe ---
                AjoutTempsTotalCycleSecondes = CInt(Left(AjoutTempsTotalCycleTexte, 2)) * 60 + CInt(Right(AjoutTempsTotalCycleTexte, 2))
                If CBSensAjoutTempsDeBain(Index).Caption = TEXTE_EN_MOINS Then
                    AjoutTempsTotalCycleSecondes = -AjoutTempsTotalCycleSecondes
                End If
        
                '--- calcul du nouveau temps de la couche T (chromage) en secondes ---
                If NbrTotalCouches = 1 Then
                    NouveauTempsCoucheTSecondes = .TDetailsOF.TGammeChrome.TMonoCouche.T + AjoutTempsTotalCycleSecondes
                Else
                    NouveauTempsCoucheTSecondes = .TDetailsOF.TGammeChrome.TMultiCouches(NbrTotalCouches - 1).T + AjoutTempsTotalCycleSecondes
                End If
                
                '--- calcul du nouveau temps de chromage à faire ---
                NouveauTempsTotalChromageAFaire = .TDetailsOF.TempsTotalChromageAFaire + AjoutTempsTotalCycleSecondes
                
                '-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
                        
                '--- transfert des nouvelles valeurs dans l'automate ---
                If AppelFenetre(F_MESSAGE, _
                                         TITRE_MESSAGES, _
                                         MESSAGE_4, _
                                         TYPES_MESSAGES.T_ATTENTION, _
                                         TYPES_BOUTONS.T_OUI_NON, _
                                         EMPLACEMENT_FOCUS.E_SUR_NON) = vbYes Then
                
                    '--- affectation du groupe ---
                    NomGroupe = "CHARGE_0" & NumCharge
            
                    '*************************************************************************************************************************************
                    
                    If AjoutTempsTotalCycleTexte <> "00:00" Then
                    
                        '--- transfert dans l'automate du nouveau temps total du cycle ---
                        ValeurRetourneeAPI = APIEcritureVariableNommee(NomGroupe, "TempsTotalChromageAFaire", NouveauTempsTotalChromageAFaire)
                        If ValeurRetourneeAPI <> 0 Then
                            Bidon = MessageErreur(TITRE_MESSAGES, MESSAGE_500)
                        End If
                        
                        '--- transfert dans l'automate du temps de la couche T ---
                        Select Case NbrTotalCouches
                            Case 1: NomElement = "Mono_TempsCouche"
                            Case Else: NomElement = "Multi" & Right("00" & CStr(NbrTotalCouches - 1), 2) & "_TempsCouche"
                        End Select
                        ValeurRetourneeAPI = APIEcritureVariableNommee(NomGroupe, NomElement, NouveauTempsCoucheTSecondes)
                        If ValeurRetourneeAPI <> 0 Then
                            Bidon = MessageErreur(TITRE_MESSAGES, MESSAGE_500)
                        End If
                            
                    End If
                            
                    '*************************************************************************************************************************************
                
                    '--- affectation de la nouvelle intensité au format texte ---
                    NouvelleIntensiteTexte = MEBIntensite(Index).Text
                    NouvelleIntensiteTexte = Replace(NouvelleIntensiteTexte, "_", "")
                    
                    If NouvelleIntensiteTexte <> "" Then
                    
                        '--- analyse du champ intensité ---
                        If IsNumeric(NouvelleIntensiteTexte) = True Then
                    
                            '--- affectation de l'intensité ---
                            NouvelleIntensite = CLng(NouvelleIntensiteTexte)
                    
                            If NouvelleIntensite >= 0 And NouvelleIntensite <= TEtatsRedresseurs(Index).DefinitionRedresseur.IMaxiRedresseur Then
                        
                                '--- intensité de la couche T ---
                                Select Case NbrTotalCouches
                                    Case 1: NomElement = "Mono_IntensiteTempsCouche"
                                    Case Else: NomElement = "Multi" & Right("00" & CStr(NbrTotalCouches - 1), 2) & "_IntensiteTempsCouche"
                                End Select
                                ValeurRetourneeAPI = APIEcritureVariableNommee(NomGroupe, NomElement, NouvelleIntensite)
                                If ValeurRetourneeAPI <> 0 Then
                                    Bidon = MessageErreur(TITRE_MESSAGES, MESSAGE_500)
                                End If
    
                            Else
    
                                '--- lancer un message d'erreur ---
                                Bidon = MessageErreur(TITRE_MESSAGES, vbCrLf & vbCrLf & vbCrLf & "cs|La valeur de l'intensité n'est pas conforme.")
    
                            End If
                        
                        Else
                            
                            '--- lancer un message d'erreur ---
                            Bidon = MessageErreur(TITRE_MESSAGES, vbCrLf & vbCrLf & vbCrLf & _
                                                                                                    "cs|La valeur de l'intensité doit être supérieure ou" & vbCrLf & _
                                                                                                    "cs|égale à 0 et inférieure ou égale à " & TEtatsRedresseurs(Index).DefinitionRedresseur.IMaxiRedresseur & " A")
                    
                        End If
                
                    End If
                
                End If
            
            End If

        End With
    
    End If

End Sub

Private Sub Form_Activate()
    
    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next
    
    '--- renseigne la Fenetre principale ---
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
    
    '--- gestion des touches communes ---
    Call OccFSynoptique.GestionTouches(KeyCode, Shift)

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    On Error Resume Next
    If UnloadMode = vbFormControlMenu Then          'obligation de passer par le bouton quitter
        Cancel = True
        CBQuitter_Click
    End If
End Sub

Private Sub Form_Resize()
    
    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next
    
    '--- zone mére et fille du déplacement de la Fenetre ---
    PBDeplacementFenetre(ZONES_DEPLACEMENT_Fenetre.Z_MERE).Height = Abs(Me.ScaleHeight - PBRenseignementsFenetre.Height - PBBoutons.Height)
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
    PBDeplacementFenetre(ZONES_DEPLACEMENT_Fenetre.Z_FILLE).Left = -HSDeplacementFenetre.value
End Sub

Private Sub LRenseignementsFenetre_DblClick()
    On Error Resume Next
    If Me.WindowState = vbMaximized Then
        Me.WindowState = vbNormal
    Else
        Me.WindowState = vbMaximized
    End If
End Sub

Private Sub MEBIntensite_Change(Index As Integer)
    On Error Resume Next
    If InterdireEvenements = False Then
        CBAnnulerTransfertModifications(Index).Enabled = True
        CBTransfertModificationsVersAPI(Index).Enabled = True
    End If
End Sub

Private Sub MEBIntensite_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    On Error Resume Next
    FiltreToucheFonction KeyCode, Shift
End Sub

Private Sub MEBTempsTotalCycle_Change(Index As Integer)
    On Error Resume Next
    If InterdireEvenements = False Then
        CBAnnulerTransfertModifications(Index).Enabled = True
        CBTransfertModificationsVersAPI(Index).Enabled = True
    End If
End Sub

Private Sub MEBTempsTotalCycle_GotFocus(Index As Integer)
    On Error Resume Next
    With ActiveControl
        .SelStart = 0          'met en surbrillance la sélection saisie
        .SelLength = Len(.Text)
    End With
End Sub

Private Sub MEBTempsTotalCycle_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    On Error Resume Next
    FiltreToucheFonction KeyCode, Shift
End Sub

Private Sub MEBTempsTotalCycle_LostFocus(Index As Integer)
    On Error Resume Next
    InterdireEvenements = True
    MEBTempsTotalCycle(Index).Text = Replace(MEBTempsTotalCycle(Index).Text, "_", "0")
    InterdireEvenements = False
End Sub

Private Sub PBBoutons_Resize()
    
    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next
    
    '--- déclaration ---
    
    '--- calculs de l'emplacement des boutons ---
    CBQuitter.Left = PBBoutons.ScaleWidth - MARGES.M_BORD_DROIT - CBQuitter.Width
    
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

Private Sub PBDeplacementFenetre_Resize(Index As Integer)
    
    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next
        
    If Index = ZONES_DEPLACEMENT_Fenetre.Z_MERE Then

        If Me.WindowState = vbMaximized Then
            
            '--- agrandir la zone fille ---
            With PBDeplacementFenetre(ZONES_DEPLACEMENT_Fenetre.Z_FILLE)
                
                .Left = PBDeplacementFenetre(ZONES_DEPLACEMENT_Fenetre.Z_MERE).ScaleLeft
                .Top = PBDeplacementFenetre(ZONES_DEPLACEMENT_Fenetre.Z_MERE).ScaleTop
                .Height = PBDeplacementFenetre(ZONES_DEPLACEMENT_Fenetre.Z_MERE).ScaleHeight
                .Width = PBDeplacementFenetre(ZONES_DEPLACEMENT_Fenetre.Z_MERE).ScaleWidth
            
                '--- agrandir en proportion de la zone fille ---
            
            End With
                   
        End If

    Else

    End If
            
    '--- valeur des curseurs ---
    If Me.WindowState <> vbMaximized And Me.WindowState <> vbMinimized Then
        HSDeplacementFenetre.Max = PBDeplacementFenetre(ZONES_DEPLACEMENT_Fenetre.Z_FILLE).Width - _
                                                         PBDeplacementFenetre(ZONES_DEPLACEMENT_Fenetre.Z_MERE).Width
        VSDeplacementFenetre.Max = PBDeplacementFenetre(ZONES_DEPLACEMENT_Fenetre.Z_FILLE).Height - _
                                                         PBDeplacementFenetre(ZONES_DEPLACEMENT_Fenetre.Z_MERE).Height
    End If
    
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

Private Sub TimerEtatsRedresseurs_Timer()
    
    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next
    
    '--- appel de la routine ---
    TimerEtatsRedresseurs.Enabled = False        'bloquage du timer
    EtatsRedresseurs
    TimerEtatsRedresseurs.Enabled = True         'rétablissement du timer
    
    '--- bip de passage dans la routine UNIQUEMENT POUR LES TESTS ---
    'If PROGRAMME_AVEC_AUTOMATE = False Then Beep

End Sub

Private Sub VSDeplacementFENETRE_Change()
    On Error Resume Next
    PBDeplacementFenetre(ZONES_DEPLACEMENT_Fenetre.Z_FILLE).Top = -VSDeplacementFenetre.value
End Sub

'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Rôle      : Affiche la totalité des charges
' Entrées :
' Retours :
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Private Sub AffichageTotaliteCharges()
   
    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next
    
    '--- constantes privées ---
    Const CROIX_DE_CONDAMNATION As String = "croix de condamnation"
    Const RECTANGLE_VERT As String = "rectangle vert"
    Const RECTANGLE_ROUGE As String = "rectangle rouge"
    Const RECTANGLE_BLANC As String = "rectangle blanc"
    
    '--- déclaration ---
    Dim a As Integer                            'pour les boucles FOR...NEXT
    Dim Texte As String                       'représente un texte quelconque
    
    '--- affichage pour les postes ---
    For a = POSTES.P_ATTAQUE To POSTES.P_Cr4_2
        
        Select Case a
        
            Case POSTES.P_ATTAQUE, POSTES.P_Cr1_1 To POSTES.P_Cr2_2, POSTES.P_Cr3_1 To POSTES.P_Cr4_2
                '--- postes concernés ---
                With TEtatsPostes(a)
        
                    If .Condamnation = True Then
                                
                        '--- affichage de la croix de condamnation ---
                        Texte = "X"
                        AffichageTexte Me.LNumCharges(a), Texte, COULEURS.BLANC, COULEURS.ROUGE_3
                        
                    Else
                        
                        '--- affichage du numéro des charges ---
                        If .NumCharge >= CHARGES.C_NUM_MINI And .NumCharge <= CHARGES.C_NUM_MAXI Then
                        
                            '--- affichage du numéro de charge ---
                            Texte = CStr(.NumCharge)
                            If TEtatsCharges(.NumCharge).TDetailsOF.LeTypeOF = T_REPIQUAGE Then
                                Texte = Texte & "R"
                            End If
                            AffichageTexte LNumCharges(a), Texte, COULEURS.JAUNE_2, COULEURS.NOIR
                                    
                        Else
                            
                            '--- vider le champ ---
                            AffichageTexte LNumCharges(a), "", COULEURS.BLANC, COULEURS.NOIR
                        
                        End If
        
                    End If
                
                End With
    
            Case Else
        End Select
    
    Next a
    
End Sub

'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Rôle      : Affiche la totalité des données des redresseurs
' Entrées :
' Retours :
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Private Sub AffichageDonneesRedresseurs()
    
    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next
    
    '--- constantes privées ---
    
    '--- déclaration ---
    Dim a As Integer                'pour les boucles FOR...NEXT
    Dim Texte As String           'repésente un texte quelconque

    For a = REDRESSEURS.R_ATTAQUE To REDRESSEURS.R_CHROME_4
    
        With TEtatsRedresseurs(a)
        
            '***********************************************************************************************************************************************
            '                                                                               SUR LE DESSIN DU REDRESSEUR
            '***********************************************************************************************************************************************
            
            '--- mode du redresseur ---
            Select Case .ModeRedresseur
                Case MODES_REDRESSEUR.M_MANUEL: OCXRedresseurs(a).Mode = MODE_MANUEL
                Case MODES_REDRESSEUR.M_AUTOMATIQUE: OCXRedresseurs(a).Mode = MODE_AUTOMATIQUE
                Case Else: OCXRedresseurs(a).Mode = MODE_NON_DEFINI
            End Select
            
            '--- intensité ---
            If .EtatRedresseur = E_ARRET Then
                OCXRedresseurs(a).Intensite = 0
            Else
                OCXRedresseurs(a).Intensite = .I
            End If
            
            '--- ah ---
            OCXRedresseurs(a).Ah = .Ah
            
            '--- anodique / cathodique ---
            Select Case a
                Case REDRESSEURS.R_ATTAQUE: OCXRedresseurs(a).Sens = SENS_ANODIQUE
                Case Else: OCXRedresseurs(a).Sens = SENS_CATHODIQUE
            End Select

            '--- temps restant de la phase (99:59 possible) ---
            If .TempsPhaseEnCours > 0 And .TempsEcoulePhaseEnCours > 0 Then
                OCXRedresseurs(a).TempsRestantPhase = CTemps3(Abs(.TempsPhaseEnCours - .TempsEcoulePhaseEnCours))
            Else
                OCXRedresseurs(a).TempsRestantPhase = "-"
            End If
            
            '--- vu-mètre de la phase en cours ---
            Select Case .NumPhaseEnCours
                Case PHASES_GAMMES_CHROME.PH_T1 To PHASES_GAMMES_CHROME.PH_T: OCXRedresseurs(a).Phase = .NumPhaseEnCours
                Case Else: OCXRedresseurs(a).Phase = ETEINT
            End Select
            
            '--- état du redresseur ---
            Select Case .EtatRedresseur
                Case ETATS_REDRESSEUR.E_ARRET To ETATS_REDRESSEUR.E_DEFAUT: OCXRedresseurs(a).Etat = .EtatRedresseur
                Case Else:  OCXRedresseurs(a).Etat = ETAT_NON_DEFINI
            End Select
                
            '***********************************************************************************************************************************************
            '                                                                               SUR LES CHAMPS D'AFFICHAGE
            '***********************************************************************************************************************************************
                
            '--- tension ---
            Texte = Format(.U, FORMAT_TENSION_2_DECIMALES_UNITE)
            AffichageTexte LTension(a), Texte
            
            '--- temps total du cycle ---
            Texte = CTemps(.TempsTotalCycle)
            AffichageTexte LTempsTotalCycle(a), Texte
            
            '--- temps ajouté sur une intensité plus faible que prévue (panne de redresseur) ---
            Select Case a
                Case REDRESSEURS.R_CHROME_1 To REDRESSEURS.R_CHROME_4
                    If .TempsAjouteSurIFaible = 0 Then
                        AffichageTexte LTempsAjouteSurIFaible(a), "Pas de compensation de temps", COULEURS.BLANC, COULEURS.BLEU_3
                    Else
                        AffichageTexte LTempsAjouteSurIFaible(a), "Dont compensation = " & CTemps(.TempsAjouteSurIFaible), COULEURS.ROUGE_3, COULEURS.BLANC
                    End If
                Case Else
            End Select
            
            '--- temps restant du cycle ---
            Texte = CTemps(.TempsRestantCycle)
            AffichageTexte LTempsRestantCycle(a), Texte
                
            '--- couche en cours ---
            If .NbrTotalCouches > 1 Then
                Texte = "COUCHE " & Succ(.NumCoucheEnCours) & " / " & .NbrTotalCouches
            Else
                Texte = "MONO-COUCHE"
            End If
            AffichageTexte LCoucheEnCours(a), Texte
                
            '--- phase en cours ---
            Select Case .NumPhaseEnCours
                Case PHASES_GAMMES_CHROME.PH_T1 To PHASES_GAMMES_CHROME.PH_T
                    Texte = "PHASE " & .NumPhaseEnCours & " - " & CTemps(.TempsEcoulePhaseEnCours) & " / " & CTemps(.TempsPhaseEnCours)
                Case Else
                    Texte = "-"
            End Select
            AffichageTexte LPhaseEnCours(a), Texte
        
        End With

    Next a

End Sub


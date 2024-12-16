VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form FGestionRedresseurs 
   BackColor       =   &H00C0C0C0&
   Caption         =   "REDRESSEURS"
   ClientHeight    =   15630
   ClientLeft      =   300
   ClientTop       =   2280
   ClientWidth     =   27195
   BeginProperty Font 
      Name            =   "Marlett"
      Size            =   8.25
      Charset         =   2
      Weight          =   500
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FGestionRedresseurs.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   15630
   ScaleWidth      =   27195
   Begin VB.PictureBox PBRenseignementsFenetre 
      Align           =   1  'Align Top
      BackColor       =   &H00FF0000&
      ForeColor       =   &H0000FFFF&
      Height          =   375
      Left            =   0
      Picture         =   "FGestionRedresseurs.frx":014A
      ScaleHeight     =   315
      ScaleWidth      =   27135
      TabIndex        =   28
      Top             =   0
      Width           =   27195
      Begin VB.Label LRenseignementsFenetre 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "REDRESSEUR GEREE"
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
         Height          =   225
         Left            =   240
         TabIndex        =   29
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
      ScaleWidth      =   27135
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   14535
      Width           =   27195
      Begin VB.CommandButton CBQuitter 
         BackColor       =   &H00FFFFFF&
         Cancel          =   -1  'True
         Caption         =   "&Quitter"
         DownPicture     =   "FGestionRedresseurs.frx":24A8C
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
         Left            =   17460
         MaskColor       =   &H00FF00FF&
         Picture         =   "FGestionRedresseurs.frx":2518E
         Style           =   1  'Graphical
         TabIndex        =   37
         ToolTipText     =   " Quitter cette fenêtre "
         Top             =   120
         UseMaskColor    =   -1  'True
         Width           =   1515
      End
      Begin VB.PictureBox PBOutilsDeplacementFenetre 
         BackColor       =   &H00E0E0E0&
         Height          =   1035
         Left            =   0
         ScaleHeight     =   975
         ScaleWidth      =   1155
         TabIndex        =   33
         Top             =   0
         Visible         =   0   'False
         Width           =   1215
         Begin VB.HScrollBar HSDeplacementFenetre 
            Height          =   255
            LargeChange     =   300
            Left            =   0
            SmallChange     =   100
            TabIndex        =   36
            Top             =   720
            Width           =   915
         End
         Begin VB.VScrollBar VSDeplacementFenetre 
            Height          =   975
            LargeChange     =   300
            Left            =   900
            SmallChange     =   100
            TabIndex        =   35
            Top             =   0
            Width           =   255
         End
         Begin VB.CommandButton CBAgrandirFenetre 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Agrandir"
            DownPicture     =   "FGestionRedresseurs.frx":25890
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   720
            Left            =   0
            MaskColor       =   &H00FF00FF&
            Picture         =   "FGestionRedresseurs.frx":25A3A
            Style           =   1  'Graphical
            TabIndex        =   34
            ToolTipText     =   " Agrandissement de la fenêtre "
            Top             =   0
            UseMaskColor    =   -1  'True
            Width           =   900
         End
      End
      Begin VB.Timer TimerEtatsRedresseurs 
         Interval        =   500
         Left            =   1380
         Top             =   480
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
   Begin VB.PictureBox PBDeplacementFenetre 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   13545
      Index           =   0
      Left            =   0
      ScaleHeight     =   13545
      ScaleWidth      =   27195
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   375
      Width           =   27195
      Begin VB.PictureBox PBDeplacementFenetre 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   13305
         Index           =   1
         Left            =   0
         ScaleHeight     =   13305
         ScaleWidth      =   28350
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   0
         Width           =   28350
         Begin VB.PictureBox PBCadreLectureValeurs 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0FFC0&
            ForeColor       =   &H80000008&
            Height          =   12135
            Index           =   2
            Left            =   12600
            ScaleHeight     =   12105
            ScaleWidth      =   3945
            TabIndex        =   76
            Top             =   480
            Width           =   3975
            Begin Anodisation.OCXRedresseur OCXRedresseurs 
               Height          =   6480
               Index           =   2
               Left            =   180
               TabIndex        =   81
               Top             =   180
               Width           =   3570
               _ExtentX        =   6297
               _ExtentY        =   11430
               Modele          =   2
            End
            Begin VB.PictureBox PBCadreModificationValeurs 
               Appearance      =   0  'Flat
               BackColor       =   &H00C0E0FF&
               ForeColor       =   &H80000008&
               Height          =   1995
               Index           =   1
               Left            =   180
               ScaleHeight     =   1965
               ScaleWidth      =   3525
               TabIndex        =   90
               Top             =   9900
               Width           =   3555
               Begin VB.CommandButton CBAnnulerTransfertModifications 
                  BackColor       =   &H00C0FFFF&
                  Caption         =   "Annuler"
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
                  Height          =   375
                  Index           =   2
                  Left            =   120
                  Style           =   1  'Graphical
                  TabIndex        =   93
                  ToolTipText     =   " Annule l'entrée des données "
                  Top             =   1500
                  Width           =   1515
               End
               Begin VB.CommandButton CBTransfertModificationsVersAPI 
                  BackColor       =   &H00C0C0FF&
                  Caption         =   "Transférer"
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
                  Height          =   375
                  Index           =   2
                  Left            =   1920
                  Style           =   1  'Graphical
                  TabIndex        =   92
                  ToolTipText     =   " Transfère les valeurs dans l'automate "
                  Top             =   1500
                  Width           =   1515
               End
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
                  TabIndex        =   91
                  Top             =   150
                  Width           =   615
               End
               Begin MSMask.MaskEdBox MEBTempsTotalCycle 
                  Height          =   315
                  Index           =   2
                  Left            =   2400
                  TabIndex        =   94
                  Top             =   180
                  Width           =   915
                  _ExtentX        =   1614
                  _ExtentY        =   556
                  _Version        =   393216
                  ForeColor       =   16711680
                  MaxLength       =   5
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "MS Sans Serif"
                     Size            =   9.75
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
                  TabIndex        =   95
                  Top             =   900
                  Width           =   915
                  _ExtentX        =   1614
                  _ExtentY        =   556
                  _Version        =   393216
                  ForeColor       =   16711680
                  PromptInclude   =   0   'False
                  MaxLength       =   5
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "MS Sans Serif"
                     Size            =   9.75
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Mask            =   "#####"
                  PromptChar      =   "_"
               End
               Begin MSMask.MaskEdBox MEBTension 
                  Height          =   315
                  Index           =   2
                  Left            =   840
                  TabIndex        =   123
                  Top             =   900
                  Width           =   675
                  _ExtentX        =   1191
                  _ExtentY        =   556
                  _Version        =   393216
                  ForeColor       =   16711680
                  MaxLength       =   4
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "MS Sans Serif"
                     Size            =   9.75
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Mask            =   "##.#"
                  PromptChar      =   "_"
               End
               Begin VB.Label LLibelles 
                  Alignment       =   1  'Right Justify
                  Appearance      =   0  'Flat
                  BackColor       =   &H0080C0FF&
                  BackStyle       =   0  'Transparent
                  Caption         =   "I (A)"
                  BeginProperty Font 
                     Name            =   "Arial"
                     Size            =   11.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H80000008&
                  Height          =   240
                  Index           =   15
                  Left            =   1740
                  TabIndex        =   125
                  Top             =   900
                  Width           =   525
               End
               Begin VB.Label LLibelles 
                  Alignment       =   1  'Right Justify
                  Appearance      =   0  'Flat
                  BackColor       =   &H0080C0FF&
                  BackStyle       =   0  'Transparent
                  Caption         =   "U (V)"
                  BeginProperty Font 
                     Name            =   "Arial"
                     Size            =   11.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H80000008&
                  Height          =   240
                  Index           =   11
                  Left            =   120
                  TabIndex        =   124
                  Top             =   900
                  Width           =   585
               End
               Begin VB.Line LDecoration 
                  BorderWidth     =   2
                  Index           =   1
                  X1              =   -60
                  X2              =   4140
                  Y1              =   720
                  Y2              =   720
               End
               Begin VB.Label LLibelles 
                  Alignment       =   2  'Center
                  Appearance      =   0  'Flat
                  BackColor       =   &H0080C0FF&
                  BackStyle       =   0  'Transparent
                  Caption         =   "Ajout au temps de bain (mm:ss)"
                  BeginProperty Font 
                     Name            =   "Arial"
                     Size            =   9.75
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H80000008&
                  Height          =   495
                  Index           =   5
                  Left            =   60
                  TabIndex        =   117
                  Top             =   60
                  Width           =   1635
               End
               Begin VB.Shape SDecoration 
                  BackColor       =   &H000040C0&
                  BackStyle       =   1  'Opaque
                  Height          =   615
                  Index           =   1
                  Left            =   -120
                  Top             =   1380
                  Width           =   4035
               End
            End
            Begin VB.PictureBox PBCadreCommandesReseauRS485 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFC0C0&
               ForeColor       =   &H80000008&
               Height          =   1095
               Index           =   2
               Left            =   180
               ScaleHeight     =   1065
               ScaleWidth      =   3525
               TabIndex        =   77
               Top             =   8400
               Width           =   3555
               Begin VB.CommandButton CBExclusionRedresseur 
                  BackColor       =   &H00FFFFC0&
                  Caption         =   "EXCLUSION du redresseur"
                  BeginProperty Font 
                     Name            =   "Arial"
                     Size            =   9.75
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
                  TabIndex        =   80
                  Top             =   600
                  Width           =   3315
               End
               Begin VB.CommandButton CBMarcheRedresseur 
                  BackColor       =   &H00FFFFC0&
                  Caption         =   "MARCHE"
                  BeginProperty Font 
                     Name            =   "Arial"
                     Size            =   9.75
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
                  TabIndex        =   79
                  Top             =   120
                  Width           =   1575
               End
               Begin VB.CommandButton CBArretRedresseur 
                  BackColor       =   &H00FFFFC0&
                  Caption         =   "ARRET"
                  BeginProperty Font 
                     Name            =   "Arial"
                     Size            =   9.75
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
                  TabIndex        =   78
                  Top             =   120
                  Width           =   1575
               End
            End
            Begin VB.Label LLibelles 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H000040C0&
               BorderStyle     =   1  'Fixed Single
               Caption         =   "MODIFICATION de la GAMME"
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
               Height          =   315
               Index           =   4
               Left            =   180
               TabIndex        =   89
               Top             =   9600
               Width           =   3555
            End
            Begin VB.Label LPhaseEnCours 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               BorderStyle     =   1  'Fixed Single
               BeginProperty Font 
                  Name            =   "Microsoft Sans Serif"
                  Size            =   11.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FF0000&
               Height          =   315
               Index           =   2
               Left            =   180
               TabIndex        =   88
               Top             =   7080
               Width           =   3555
            End
            Begin VB.Label LTempsRestantCycle 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               BorderStyle     =   1  'Fixed Single
               BeginProperty Font 
                  Name            =   "Microsoft Sans Serif"
                  Size            =   11.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FF0000&
               Height          =   315
               Index           =   2
               Left            =   180
               TabIndex        =   87
               Top             =   7680
               Width           =   2115
            End
            Begin VB.Label LTempsTotalCycle 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               BorderStyle     =   1  'Fixed Single
               BeginProperty Font 
                  Name            =   "Microsoft Sans Serif"
                  Size            =   11.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FF0000&
               Height          =   315
               Index           =   2
               Left            =   2280
               TabIndex        =   86
               Top             =   7680
               Width           =   1455
            End
            Begin VB.Label LLibelles 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H0080FF80&
               BorderStyle     =   1  'Fixed Single
               Caption         =   "TOTAL"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   11.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   315
               Index           =   25
               Left            =   2280
               TabIndex        =   85
               Top             =   7380
               Width           =   1455
            End
            Begin VB.Label LLibelles 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H0080FF80&
               BorderStyle     =   1  'Fixed Single
               Caption         =   "TEMPS RESTANT"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   11.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   315
               Index           =   23
               Left            =   180
               TabIndex        =   84
               Top             =   7380
               Width           =   2115
            End
            Begin VB.Label LLibelles 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H0080FF80&
               BorderStyle     =   1  'Fixed Single
               Caption         =   "PHASE"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   11.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   315
               Index           =   22
               Left            =   180
               TabIndex        =   83
               Top             =   6780
               Width           =   3555
            End
            Begin VB.Label LLibelles 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H00FF0000&
               BorderStyle     =   1  'Fixed Single
               Caption         =   "COMMANDES"
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
               Height          =   315
               Index           =   59
               Left            =   180
               TabIndex        =   82
               Top             =   8100
               Width           =   3555
            End
         End
         Begin VB.PictureBox PBCadreLectureValeurs 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0FFC0&
            ForeColor       =   &H80000008&
            Height          =   12135
            Index           =   3
            Left            =   8460
            ScaleHeight     =   12105
            ScaleWidth      =   3945
            TabIndex        =   39
            Top             =   480
            Width           =   3975
            Begin VB.PictureBox PBCadreModificationValeurs 
               Appearance      =   0  'Flat
               BackColor       =   &H00C0E0FF&
               ForeColor       =   &H80000008&
               Height          =   1995
               Index           =   2
               Left            =   180
               ScaleHeight     =   1965
               ScaleWidth      =   3525
               TabIndex        =   97
               Top             =   9900
               Width           =   3555
               Begin VB.CommandButton CBAnnulerTransfertModifications 
                  BackColor       =   &H00C0FFFF&
                  Caption         =   "Annuler"
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
                  Height          =   375
                  Index           =   3
                  Left            =   120
                  Style           =   1  'Graphical
                  TabIndex        =   100
                  ToolTipText     =   " Annule l'entrée des données "
                  Top             =   1500
                  Width           =   1515
               End
               Begin VB.CommandButton CBTransfertModificationsVersAPI 
                  BackColor       =   &H00C0C0FF&
                  Caption         =   "Transférer"
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
                  Height          =   375
                  Index           =   3
                  Left            =   1920
                  Style           =   1  'Graphical
                  TabIndex        =   99
                  ToolTipText     =   " Transfère les valeurs dans l'automate "
                  Top             =   1500
                  Width           =   1515
               End
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
                  TabIndex        =   98
                  Top             =   150
                  Width           =   615
               End
               Begin MSMask.MaskEdBox MEBTempsTotalCycle 
                  Height          =   315
                  Index           =   3
                  Left            =   2400
                  TabIndex        =   101
                  Top             =   180
                  Width           =   915
                  _ExtentX        =   1614
                  _ExtentY        =   556
                  _Version        =   393216
                  ForeColor       =   16711680
                  MaxLength       =   5
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "MS Sans Serif"
                     Size            =   9.75
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
                  TabIndex        =   102
                  Top             =   900
                  Width           =   915
                  _ExtentX        =   1614
                  _ExtentY        =   556
                  _Version        =   393216
                  ForeColor       =   16711680
                  MaxLength       =   5
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "MS Sans Serif"
                     Size            =   9.75
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Mask            =   "#####"
                  PromptChar      =   "_"
               End
               Begin MSMask.MaskEdBox MEBTension 
                  Height          =   315
                  Index           =   3
                  Left            =   840
                  TabIndex        =   126
                  Top             =   900
                  Width           =   675
                  _ExtentX        =   1191
                  _ExtentY        =   556
                  _Version        =   393216
                  ForeColor       =   16711680
                  MaxLength       =   4
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "MS Sans Serif"
                     Size            =   9.75
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Mask            =   "##.#"
                  PromptChar      =   "_"
               End
               Begin VB.Label LLibelles 
                  Alignment       =   1  'Right Justify
                  Appearance      =   0  'Flat
                  BackColor       =   &H0080C0FF&
                  BackStyle       =   0  'Transparent
                  Caption         =   "I (A)"
                  BeginProperty Font 
                     Name            =   "Arial"
                     Size            =   11.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H80000008&
                  Height          =   240
                  Index           =   20
                  Left            =   1740
                  TabIndex        =   128
                  Top             =   900
                  Width           =   525
               End
               Begin VB.Label LLibelles 
                  Alignment       =   1  'Right Justify
                  Appearance      =   0  'Flat
                  BackColor       =   &H0080C0FF&
                  BackStyle       =   0  'Transparent
                  Caption         =   "U (V)"
                  BeginProperty Font 
                     Name            =   "Arial"
                     Size            =   11.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H80000008&
                  Height          =   240
                  Index           =   19
                  Left            =   120
                  TabIndex        =   127
                  Top             =   900
                  Width           =   585
               End
               Begin VB.Line LDecoration 
                  BorderWidth     =   2
                  Index           =   2
                  X1              =   0
                  X2              =   4200
                  Y1              =   720
                  Y2              =   720
               End
               Begin VB.Label LLibelles 
                  Alignment       =   2  'Center
                  Appearance      =   0  'Flat
                  BackColor       =   &H0080C0FF&
                  BackStyle       =   0  'Transparent
                  Caption         =   "Ajout au temps de bain (mm:ss)"
                  BeginProperty Font 
                     Name            =   "Arial"
                     Size            =   9.75
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H80000008&
                  Height          =   495
                  Index           =   8
                  Left            =   60
                  TabIndex        =   118
                  Top             =   60
                  Width           =   1635
               End
               Begin VB.Shape SDecoration 
                  BackColor       =   &H000040C0&
                  BackStyle       =   1  'Opaque
                  Height          =   615
                  Index           =   2
                  Left            =   -240
                  Top             =   1380
                  Width           =   4035
               End
            End
            Begin VB.PictureBox PBCadreCommandesReseauRS485 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFC0C0&
               ForeColor       =   &H80000008&
               Height          =   1095
               Index           =   3
               Left            =   180
               ScaleHeight     =   1065
               ScaleWidth      =   3525
               TabIndex        =   53
               Top             =   8400
               Width           =   3555
               Begin VB.CommandButton CBExclusionRedresseur 
                  BackColor       =   &H00FFFFC0&
                  Caption         =   "EXCLUSION du redresseur"
                  BeginProperty Font 
                     Name            =   "Arial"
                     Size            =   9.75
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
                  TabIndex        =   56
                  Top             =   600
                  Width           =   3315
               End
               Begin VB.CommandButton CBMarcheRedresseur 
                  BackColor       =   &H00FFFFC0&
                  Caption         =   "MARCHE"
                  BeginProperty Font 
                     Name            =   "Arial"
                     Size            =   9.75
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
                  TabIndex        =   55
                  Top             =   120
                  Width           =   1575
               End
               Begin VB.CommandButton CBArretRedresseur 
                  BackColor       =   &H00FFFFC0&
                  Caption         =   "ARRET"
                  BeginProperty Font 
                     Name            =   "Arial"
                     Size            =   9.75
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
                  TabIndex        =   54
                  Top             =   120
                  Width           =   1575
               End
            End
            Begin Anodisation.OCXRedresseur OCXRedresseurs 
               Height          =   6480
               Index           =   3
               Left            =   180
               TabIndex        =   40
               Top             =   180
               Width           =   3570
               _ExtentX        =   6297
               _ExtentY        =   11430
               Modele          =   2
            End
            Begin VB.Label LLibelles 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H000040C0&
               BorderStyle     =   1  'Fixed Single
               Caption         =   "MODIFICATION de la GAMME"
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
               Height          =   315
               Index           =   18
               Left            =   180
               TabIndex        =   96
               Top             =   9600
               Width           =   3555
            End
            Begin VB.Label LLibelles 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H00FF0000&
               BorderStyle     =   1  'Fixed Single
               Caption         =   "COMMANDES"
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
               Height          =   315
               Index           =   60
               Left            =   180
               TabIndex        =   52
               Top             =   8100
               Width           =   3555
            End
            Begin VB.Label LLibelles 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H0080FF80&
               BorderStyle     =   1  'Fixed Single
               Caption         =   "TOTAL"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   11.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   315
               Index           =   30
               Left            =   2280
               TabIndex        =   46
               Top             =   7380
               Width           =   1455
            End
            Begin VB.Label LLibelles 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H0080FF80&
               BorderStyle     =   1  'Fixed Single
               Caption         =   "TEMPS RESTANT"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   11.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   315
               Index           =   31
               Left            =   180
               TabIndex        =   45
               Top             =   7380
               Width           =   2115
            End
            Begin VB.Label LLibelles 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H0080FF80&
               BorderStyle     =   1  'Fixed Single
               Caption         =   "PHASE"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   11.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   315
               Index           =   33
               Left            =   180
               TabIndex        =   44
               Top             =   6780
               Width           =   3555
            End
            Begin VB.Label LTempsTotalCycle 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               BorderStyle     =   1  'Fixed Single
               BeginProperty Font 
                  Name            =   "Microsoft Sans Serif"
                  Size            =   11.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FF0000&
               Height          =   315
               Index           =   3
               Left            =   2280
               TabIndex        =   43
               Top             =   7680
               Width           =   1455
            End
            Begin VB.Label LTempsRestantCycle 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               BorderStyle     =   1  'Fixed Single
               BeginProperty Font 
                  Name            =   "Microsoft Sans Serif"
                  Size            =   11.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FF0000&
               Height          =   315
               Index           =   3
               Left            =   180
               TabIndex        =   42
               Top             =   7680
               Width           =   2115
            End
            Begin VB.Label LPhaseEnCours 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               BorderStyle     =   1  'Fixed Single
               BeginProperty Font 
                  Name            =   "Microsoft Sans Serif"
                  Size            =   11.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FF0000&
               Height          =   315
               Index           =   3
               Left            =   180
               TabIndex        =   41
               Top             =   7080
               Width           =   3555
            End
         End
         Begin VB.PictureBox PBCadreLectureValeurs 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0FFC0&
            ForeColor       =   &H80000008&
            Height          =   12135
            Index           =   0
            Left            =   16740
            ScaleHeight     =   12105
            ScaleWidth      =   3945
            TabIndex        =   22
            Top             =   480
            Width           =   3975
            Begin VB.PictureBox PBCadreModificationValeurs 
               Appearance      =   0  'Flat
               BackColor       =   &H00C0E0FF&
               ForeColor       =   &H80000008&
               Height          =   1995
               Index           =   3
               Left            =   180
               ScaleHeight     =   1965
               ScaleWidth      =   3525
               TabIndex        =   67
               Top             =   9900
               Width           =   3555
               Begin VB.CommandButton CBAnnulerTransfertModifications 
                  BackColor       =   &H00C0FFFF&
                  Caption         =   "Annuler"
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
                  Height          =   375
                  Index           =   1
                  Left            =   120
                  Style           =   1  'Graphical
                  TabIndex        =   70
                  ToolTipText     =   " Annule l'entrée des données "
                  Top             =   1500
                  Width           =   1515
               End
               Begin VB.CommandButton CBTransfertModificationsVersAPI 
                  BackColor       =   &H00C0C0FF&
                  Caption         =   "Transférer"
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
                  Height          =   375
                  Index           =   1
                  Left            =   1920
                  Style           =   1  'Graphical
                  TabIndex        =   69
                  ToolTipText     =   " Transfère les valeurs dans l'automate "
                  Top             =   1500
                  Width           =   1515
               End
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
                  TabIndex        =   68
                  Top             =   150
                  Width           =   615
               End
               Begin MSMask.MaskEdBox MEBTempsTotalCycle 
                  Height          =   315
                  Index           =   1
                  Left            =   2400
                  TabIndex        =   71
                  Top             =   180
                  Width           =   915
                  _ExtentX        =   1614
                  _ExtentY        =   556
                  _Version        =   393216
                  ForeColor       =   16711680
                  MaxLength       =   5
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "MS Sans Serif"
                     Size            =   9.75
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
                  TabIndex        =   72
                  Top             =   900
                  Width           =   915
                  _ExtentX        =   1614
                  _ExtentY        =   556
                  _Version        =   393216
                  ForeColor       =   16711680
                  MaxLength       =   5
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "MS Sans Serif"
                     Size            =   9.75
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Mask            =   "#####"
                  PromptChar      =   "_"
               End
               Begin MSMask.MaskEdBox MEBTension 
                  Height          =   315
                  Index           =   1
                  Left            =   840
                  TabIndex        =   121
                  Top             =   900
                  Width           =   675
                  _ExtentX        =   1191
                  _ExtentY        =   556
                  _Version        =   393216
                  ForeColor       =   16711680
                  MaxLength       =   4
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "MS Sans Serif"
                     Size            =   9.75
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Mask            =   "##.#"
                  PromptChar      =   "_"
               End
               Begin VB.Line LDecoration 
                  BorderWidth     =   2
                  Index           =   0
                  X1              =   -240
                  X2              =   3960
                  Y1              =   720
                  Y2              =   720
               End
               Begin VB.Label LLibelles 
                  Alignment       =   1  'Right Justify
                  Appearance      =   0  'Flat
                  BackColor       =   &H0080C0FF&
                  BackStyle       =   0  'Transparent
                  Caption         =   "U (V)"
                  BeginProperty Font 
                     Name            =   "Arial"
                     Size            =   11.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H80000008&
                  Height          =   240
                  Index           =   14
                  Left            =   120
                  TabIndex        =   122
                  Top             =   900
                  Width           =   585
               End
               Begin VB.Shape SDecoration 
                  BackColor       =   &H000040C0&
                  BackStyle       =   1  'Opaque
                  Height          =   615
                  Index           =   3
                  Left            =   -240
                  Top             =   1380
                  Width           =   4035
               End
               Begin VB.Label LLibelles 
                  Alignment       =   1  'Right Justify
                  Appearance      =   0  'Flat
                  BackColor       =   &H0080C0FF&
                  BackStyle       =   0  'Transparent
                  Caption         =   "I (A)"
                  BeginProperty Font 
                     Name            =   "Arial"
                     Size            =   11.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H80000008&
                  Height          =   240
                  Index           =   16
                  Left            =   1740
                  TabIndex        =   74
                  Top             =   900
                  Width           =   525
               End
               Begin VB.Label LLibelles 
                  Alignment       =   2  'Center
                  Appearance      =   0  'Flat
                  BackColor       =   &H0080C0FF&
                  BackStyle       =   0  'Transparent
                  Caption         =   "Ajout au temps de bain (mm:ss)"
                  BeginProperty Font 
                     Name            =   "Arial"
                     Size            =   9.75
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H80000008&
                  Height          =   495
                  Index           =   13
                  Left            =   60
                  TabIndex        =   73
                  Top             =   60
                  Width           =   1635
               End
            End
            Begin VB.PictureBox PBCadreCommandesReseauRS485 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFC0C0&
               ForeColor       =   &H80000008&
               Height          =   1095
               Index           =   1
               Left            =   180
               ScaleHeight     =   1065
               ScaleWidth      =   3525
               TabIndex        =   48
               Top             =   8400
               Width           =   3555
               Begin VB.CommandButton CBExclusionRedresseur 
                  BackColor       =   &H00FFFFC0&
                  Caption         =   "EXCLUSION du redresseur"
                  BeginProperty Font 
                     Name            =   "Arial"
                     Size            =   9.75
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
                  TabIndex        =   51
                  Top             =   600
                  Width           =   3315
               End
               Begin VB.CommandButton CBMarcheRedresseur 
                  BackColor       =   &H00FFFFC0&
                  Caption         =   "MARCHE"
                  BeginProperty Font 
                     Name            =   "Arial"
                     Size            =   9.75
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
                  TabIndex        =   50
                  Top             =   120
                  Width           =   1575
               End
               Begin VB.CommandButton CBArretRedresseur 
                  BackColor       =   &H00FFFFC0&
                  Caption         =   "ARRET"
                  BeginProperty Font 
                     Name            =   "Arial"
                     Size            =   9.75
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
                  TabIndex        =   49
                  Top             =   120
                  Width           =   1575
               End
            End
            Begin Anodisation.OCXRedresseur OCXRedresseurs 
               Height          =   6480
               Index           =   1
               Left            =   180
               TabIndex        =   30
               Top             =   180
               Width           =   3570
               _ExtentX        =   6297
               _ExtentY        =   11430
               Modele          =   2
            End
            Begin VB.Label LLibelles 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H000040C0&
               BorderStyle     =   1  'Fixed Single
               Caption         =   "MODIFICATION de la GAMME"
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
               Height          =   315
               Index           =   35
               Left            =   180
               TabIndex        =   75
               Top             =   9600
               Width           =   3555
            End
            Begin VB.Label LLibelles 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H00FF0000&
               BorderStyle     =   1  'Fixed Single
               Caption         =   "COMMANDES"
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
               Height          =   315
               Index           =   10
               Left            =   180
               TabIndex        =   47
               Top             =   8100
               Width           =   3555
            End
            Begin VB.Label LLibelles 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H0080FF80&
               BorderStyle     =   1  'Fixed Single
               Caption         =   "TOTAL"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   11.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   315
               Index           =   1
               Left            =   2280
               TabIndex        =   38
               Top             =   7380
               Width           =   1455
            End
            Begin VB.Label LLibelles 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H0080FF80&
               BorderStyle     =   1  'Fixed Single
               Caption         =   "TEMPS RESTANT"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   11.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   315
               Index           =   2
               Left            =   180
               TabIndex        =   27
               Top             =   7380
               Width           =   2115
            End
            Begin VB.Label LLibelles 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H0080FF80&
               BorderStyle     =   1  'Fixed Single
               Caption         =   "PHASE"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   11.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   315
               Index           =   3
               Left            =   180
               TabIndex        =   26
               Top             =   6780
               Width           =   3555
            End
            Begin VB.Label LTempsTotalCycle 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               BorderStyle     =   1  'Fixed Single
               BeginProperty Font 
                  Name            =   "Microsoft Sans Serif"
                  Size            =   11.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FF0000&
               Height          =   315
               Index           =   1
               Left            =   2280
               TabIndex        =   25
               Top             =   7680
               Width           =   1455
            End
            Begin VB.Label LTempsRestantCycle 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               BorderStyle     =   1  'Fixed Single
               BeginProperty Font 
                  Name            =   "Microsoft Sans Serif"
                  Size            =   11.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FF0000&
               Height          =   315
               Index           =   1
               Left            =   180
               TabIndex        =   24
               Top             =   7680
               Width           =   2115
            End
            Begin VB.Label LPhaseEnCours 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               BorderStyle     =   1  'Fixed Single
               BeginProperty Font 
                  Name            =   "Microsoft Sans Serif"
                  Size            =   11.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FF0000&
               Height          =   315
               Index           =   1
               Left            =   180
               TabIndex        =   23
               Top             =   7080
               Width           =   3555
            End
         End
         Begin VB.PictureBox PBCadreLectureValeurs 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0FF&
            ForeColor       =   &H80000008&
            Height          =   12135
            Index           =   5
            Left            =   180
            ScaleHeight     =   12105
            ScaleWidth      =   3945
            TabIndex        =   13
            Top             =   480
            Width           =   3975
            Begin VB.PictureBox PBCadreModificationValeurs 
               Appearance      =   0  'Flat
               BackColor       =   &H00C0E0FF&
               ForeColor       =   &H80000008&
               Height          =   1995
               Index           =   5
               Left            =   180
               ScaleHeight     =   1965
               ScaleWidth      =   3525
               TabIndex        =   111
               Top             =   9900
               Visible         =   0   'False
               Width           =   3555
               Begin VB.CommandButton CBAnnulerTransfertModifications 
                  BackColor       =   &H00C0FFFF&
                  Caption         =   "Annuler"
                  Enabled         =   0   'False
                  BeginProperty Font 
                     Name            =   "Arial"
                     Size            =   9.75
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
                  TabIndex        =   114
                  ToolTipText     =   " Annule l'entrée des données "
                  Top             =   1500
                  Width           =   1515
               End
               Begin VB.CommandButton CBTransfertModificationsVersAPI 
                  BackColor       =   &H00C0C0FF&
                  Caption         =   "Transférer"
                  Enabled         =   0   'False
                  BeginProperty Font 
                     Name            =   "Arial"
                     Size            =   9.75
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
                  TabIndex        =   113
                  ToolTipText     =   " Transfère les valeurs dans l'automate "
                  Top             =   1500
                  Width           =   1515
               End
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
                  TabIndex        =   112
                  Top             =   150
                  Width           =   615
               End
               Begin MSMask.MaskEdBox MEBTempsTotalCycle 
                  Height          =   315
                  Index           =   5
                  Left            =   2400
                  TabIndex        =   115
                  Top             =   180
                  Width           =   915
                  _ExtentX        =   1614
                  _ExtentY        =   556
                  _Version        =   393216
                  ForeColor       =   16711680
                  MaxLength       =   5
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "MS Sans Serif"
                     Size            =   9.75
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
                  TabIndex        =   116
                  Top             =   900
                  Width           =   915
                  _ExtentX        =   1614
                  _ExtentY        =   556
                  _Version        =   393216
                  ForeColor       =   16711680
                  MaxLength       =   5
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "MS Sans Serif"
                     Size            =   9.75
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Mask            =   "#####"
                  PromptChar      =   "_"
               End
               Begin MSMask.MaskEdBox MEBTension 
                  Height          =   315
                  Index           =   5
                  Left            =   840
                  TabIndex        =   132
                  Top             =   900
                  Width           =   675
                  _ExtentX        =   1191
                  _ExtentY        =   556
                  _Version        =   393216
                  ForeColor       =   16711680
                  MaxLength       =   4
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "MS Sans Serif"
                     Size            =   9.75
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Mask            =   "##.#"
                  PromptChar      =   "_"
               End
               Begin VB.Label LLibelles 
                  Alignment       =   1  'Right Justify
                  Appearance      =   0  'Flat
                  BackColor       =   &H0080C0FF&
                  BackStyle       =   0  'Transparent
                  Caption         =   "I (A)"
                  BeginProperty Font 
                     Name            =   "Arial"
                     Size            =   11.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H80000008&
                  Height          =   240
                  Index           =   27
                  Left            =   1740
                  TabIndex        =   134
                  Top             =   900
                  Width           =   525
               End
               Begin VB.Label LLibelles 
                  Alignment       =   1  'Right Justify
                  Appearance      =   0  'Flat
                  BackColor       =   &H0080C0FF&
                  BackStyle       =   0  'Transparent
                  Caption         =   "U (V)"
                  BeginProperty Font 
                     Name            =   "Arial"
                     Size            =   11.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H80000008&
                  Height          =   240
                  Index           =   26
                  Left            =   120
                  TabIndex        =   133
                  Top             =   900
                  Width           =   585
               End
               Begin VB.Line LDecoration 
                  BorderWidth     =   2
                  Index           =   4
                  X1              =   -120
                  X2              =   4080
                  Y1              =   720
                  Y2              =   720
               End
               Begin VB.Label LLibelles 
                  Alignment       =   2  'Center
                  Appearance      =   0  'Flat
                  BackColor       =   &H0080C0FF&
                  BackStyle       =   0  'Transparent
                  Caption         =   "Ajout au temps de bain (mm:ss)"
                  BeginProperty Font 
                     Name            =   "Arial"
                     Size            =   9.75
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H80000008&
                  Height          =   495
                  Index           =   12
                  Left            =   60
                  TabIndex        =   120
                  Top             =   60
                  Width           =   1635
               End
               Begin VB.Shape SDecoration 
                  BackColor       =   &H000040C0&
                  BackStyle       =   1  'Opaque
                  Height          =   615
                  Index           =   5
                  Left            =   -240
                  Top             =   1380
                  Width           =   4035
               End
            End
            Begin VB.PictureBox PBCadreCommandesReseauRS485 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFC0C0&
               ForeColor       =   &H80000008&
               Height          =   1095
               Index           =   5
               Left            =   180
               ScaleHeight     =   1065
               ScaleWidth      =   3525
               TabIndex        =   63
               Top             =   8400
               Visible         =   0   'False
               Width           =   3555
               Begin VB.CommandButton CBExclusionRedresseur 
                  BackColor       =   &H00FFFFC0&
                  Caption         =   "EXCLUSION du redresseur"
                  BeginProperty Font 
                     Name            =   "Arial"
                     Size            =   9.75
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
                  TabIndex        =   66
                  ToolTipText     =   " "
                  Top             =   600
                  Width           =   3315
               End
               Begin VB.CommandButton CBMarcheRedresseur 
                  BackColor       =   &H00FFFFC0&
                  Caption         =   "MARCHE"
                  BeginProperty Font 
                     Name            =   "Arial"
                     Size            =   9.75
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
                  TabIndex        =   65
                  Top             =   120
                  Width           =   1575
               End
               Begin VB.CommandButton CBArretRedresseur 
                  BackColor       =   &H00FFFFC0&
                  Caption         =   "ARRET"
                  BeginProperty Font 
                     Name            =   "Arial"
                     Size            =   9.75
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
                  TabIndex        =   64
                  Top             =   120
                  Width           =   1575
               End
            End
            Begin Anodisation.OCXRedresseur OCXRedresseurs 
               Height          =   6480
               Index           =   5
               Left            =   180
               TabIndex        =   32
               Top             =   180
               Width           =   3570
               _ExtentX        =   6297
               _ExtentY        =   11430
               Modele          =   2
            End
            Begin VB.Label LLibelles 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H000040C0&
               BorderStyle     =   1  'Fixed Single
               Caption         =   "MODIFICATION de la GAMME"
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
               Height          =   315
               Index           =   58
               Left            =   180
               TabIndex        =   110
               Top             =   9600
               Visible         =   0   'False
               Width           =   3555
            End
            Begin VB.Label LLibelles 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H00FF0000&
               BorderStyle     =   1  'Fixed Single
               Caption         =   "COMMANDES"
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
               Height          =   315
               Index           =   62
               Left            =   180
               TabIndex        =   62
               Top             =   8100
               Visible         =   0   'False
               Width           =   3555
            End
            Begin VB.Label LPhaseEnCours 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               BorderStyle     =   1  'Fixed Single
               BeginProperty Font 
                  Name            =   "Microsoft Sans Serif"
                  Size            =   11.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FF0000&
               Height          =   315
               Index           =   5
               Left            =   180
               TabIndex        =   19
               Top             =   7080
               Visible         =   0   'False
               Width           =   3555
            End
            Begin VB.Label LTempsRestantCycle 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               BorderStyle     =   1  'Fixed Single
               BeginProperty Font 
                  Name            =   "Microsoft Sans Serif"
                  Size            =   11.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FF0000&
               Height          =   315
               Index           =   5
               Left            =   180
               TabIndex        =   18
               Top             =   7680
               Visible         =   0   'False
               Width           =   2115
            End
            Begin VB.Label LTempsTotalCycle 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               BorderStyle     =   1  'Fixed Single
               BeginProperty Font 
                  Name            =   "Microsoft Sans Serif"
                  Size            =   11.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FF0000&
               Height          =   315
               Index           =   5
               Left            =   2280
               TabIndex        =   17
               Top             =   7680
               Visible         =   0   'False
               Width           =   1455
            End
            Begin VB.Label LLibelles 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H0080FF80&
               BorderStyle     =   1  'Fixed Single
               Caption         =   "PHASE"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   11.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   315
               Index           =   55
               Left            =   180
               TabIndex        =   16
               Top             =   6780
               Visible         =   0   'False
               Width           =   3555
            End
            Begin VB.Label LLibelles 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H0080FF80&
               BorderStyle     =   1  'Fixed Single
               Caption         =   "TEMPS RESTANT"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   11.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   315
               Index           =   53
               Left            =   180
               TabIndex        =   15
               Top             =   7380
               Visible         =   0   'False
               Width           =   2115
            End
            Begin VB.Label LLibelles 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H0080FF80&
               BorderStyle     =   1  'Fixed Single
               Caption         =   "TOTAL"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   11.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   315
               Index           =   52
               Left            =   2280
               TabIndex        =   14
               Top             =   7380
               Visible         =   0   'False
               Width           =   1455
            End
         End
         Begin VB.PictureBox PBCadreLectureValeurs 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0FFC0&
            ForeColor       =   &H80000008&
            Height          =   12135
            Index           =   4
            Left            =   4320
            ScaleHeight     =   12105
            ScaleWidth      =   3945
            TabIndex        =   6
            Top             =   480
            Width           =   3975
            Begin VB.PictureBox PBCadreModificationValeurs 
               Appearance      =   0  'Flat
               BackColor       =   &H00C0E0FF&
               ForeColor       =   &H80000008&
               Height          =   1995
               Index           =   4
               Left            =   180
               ScaleHeight     =   1965
               ScaleWidth      =   3525
               TabIndex        =   104
               Top             =   9900
               Width           =   3555
               Begin VB.CommandButton CBTransfertModificationsVersAPI 
                  BackColor       =   &H00C0C0FF&
                  Caption         =   "Transférer"
                  Enabled         =   0   'False
                  BeginProperty Font 
                     Name            =   "Arial"
                     Size            =   9.75
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
                  TabIndex        =   107
                  ToolTipText     =   " Transfère les valeurs dans l'automate "
                  Top             =   1500
                  Width           =   1515
               End
               Begin VB.CommandButton CBAnnulerTransfertModifications 
                  BackColor       =   &H00C0FFFF&
                  Caption         =   "Annuler"
                  Enabled         =   0   'False
                  BeginProperty Font 
                     Name            =   "Arial"
                     Size            =   9.75
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
                  TabIndex        =   106
                  ToolTipText     =   " Annule l'entrée des données "
                  Top             =   1500
                  Width           =   1515
               End
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
                  TabIndex        =   105
                  Top             =   150
                  Width           =   615
               End
               Begin MSMask.MaskEdBox MEBTempsTotalCycle 
                  Height          =   315
                  Index           =   4
                  Left            =   2400
                  TabIndex        =   108
                  Top             =   180
                  Width           =   915
                  _ExtentX        =   1614
                  _ExtentY        =   556
                  _Version        =   393216
                  ForeColor       =   16711680
                  MaxLength       =   5
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "MS Sans Serif"
                     Size            =   9.75
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
                  TabIndex        =   109
                  Top             =   900
                  Width           =   915
                  _ExtentX        =   1614
                  _ExtentY        =   556
                  _Version        =   393216
                  ForeColor       =   16711680
                  MaxLength       =   5
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "MS Sans Serif"
                     Size            =   9.75
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Mask            =   "#####"
                  PromptChar      =   "_"
               End
               Begin MSMask.MaskEdBox MEBTension 
                  Height          =   315
                  Index           =   4
                  Left            =   840
                  TabIndex        =   129
                  Top             =   900
                  Width           =   675
                  _ExtentX        =   1191
                  _ExtentY        =   556
                  _Version        =   393216
                  ForeColor       =   16711680
                  MaxLength       =   4
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "MS Sans Serif"
                     Size            =   9.75
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Mask            =   "##.#"
                  PromptChar      =   "_"
               End
               Begin VB.Label LLibelles 
                  Alignment       =   1  'Right Justify
                  Appearance      =   0  'Flat
                  BackColor       =   &H0080C0FF&
                  BackStyle       =   0  'Transparent
                  Caption         =   "I (A)"
                  BeginProperty Font 
                     Name            =   "Arial"
                     Size            =   11.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H80000008&
                  Height          =   240
                  Index           =   24
                  Left            =   1740
                  TabIndex        =   131
                  Top             =   900
                  Width           =   525
               End
               Begin VB.Label LLibelles 
                  Alignment       =   1  'Right Justify
                  Appearance      =   0  'Flat
                  BackColor       =   &H0080C0FF&
                  BackStyle       =   0  'Transparent
                  Caption         =   "U (V)"
                  BeginProperty Font 
                     Name            =   "Arial"
                     Size            =   11.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H80000008&
                  Height          =   240
                  Index           =   21
                  Left            =   120
                  TabIndex        =   130
                  Top             =   900
                  Width           =   585
               End
               Begin VB.Line LDecoration 
                  BorderWidth     =   2
                  Index           =   3
                  X1              =   -180
                  X2              =   4020
                  Y1              =   720
                  Y2              =   720
               End
               Begin VB.Label LLibelles 
                  Alignment       =   2  'Center
                  Appearance      =   0  'Flat
                  BackColor       =   &H0080C0FF&
                  BackStyle       =   0  'Transparent
                  Caption         =   "Ajout au temps de bain (mm:ss)"
                  BeginProperty Font 
                     Name            =   "Arial"
                     Size            =   9.75
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H80000008&
                  Height          =   495
                  Index           =   9
                  Left            =   60
                  TabIndex        =   119
                  Top             =   60
                  Width           =   1635
               End
               Begin VB.Shape SDecoration 
                  BackColor       =   &H000040C0&
                  BackStyle       =   1  'Opaque
                  Height          =   615
                  Index           =   4
                  Left            =   -240
                  Top             =   1380
                  Width           =   4035
               End
            End
            Begin VB.PictureBox PBCadreCommandesReseauRS485 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFC0C0&
               ForeColor       =   &H80000008&
               Height          =   1095
               Index           =   4
               Left            =   180
               ScaleHeight     =   1065
               ScaleWidth      =   3525
               TabIndex        =   58
               Top             =   8400
               Width           =   3555
               Begin VB.CommandButton CBExclusionRedresseur 
                  BackColor       =   &H00FFFFC0&
                  Caption         =   "EXCLUSION du redresseur"
                  BeginProperty Font 
                     Name            =   "Arial"
                     Size            =   9.75
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
                  TabIndex        =   61
                  Top             =   600
                  Width           =   3315
               End
               Begin VB.CommandButton CBMarcheRedresseur 
                  BackColor       =   &H00FFFFC0&
                  Caption         =   "MARCHE"
                  BeginProperty Font 
                     Name            =   "Arial"
                     Size            =   9.75
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
                  TabIndex        =   60
                  Top             =   120
                  Width           =   1575
               End
               Begin VB.CommandButton CBArretRedresseur 
                  BackColor       =   &H00FFFFC0&
                  Caption         =   "ARRET"
                  BeginProperty Font 
                     Name            =   "Arial"
                     Size            =   9.75
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
                  TabIndex        =   59
                  Top             =   120
                  Width           =   1575
               End
            End
            Begin Anodisation.OCXRedresseur OCXRedresseurs 
               Height          =   6480
               Index           =   4
               Left            =   180
               TabIndex        =   31
               Top             =   180
               Width           =   3570
               _ExtentX        =   6297
               _ExtentY        =   11430
               Modele          =   2
            End
            Begin VB.Label LLibelles 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H000040C0&
               BorderStyle     =   1  'Fixed Single
               Caption         =   "MODIFICATION de la GAMME"
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
               Height          =   315
               Index           =   40
               Left            =   180
               TabIndex        =   103
               Top             =   9600
               Width           =   3555
            End
            Begin VB.Label LLibelles 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H00FF0000&
               BorderStyle     =   1  'Fixed Single
               Caption         =   "COMMANDES"
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
               Height          =   315
               Index           =   61
               Left            =   180
               TabIndex        =   57
               Top             =   8100
               Width           =   3555
            End
            Begin VB.Label LPhaseEnCours 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               BorderStyle     =   1  'Fixed Single
               BeginProperty Font 
                  Name            =   "Microsoft Sans Serif"
                  Size            =   11.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FF0000&
               Height          =   315
               Index           =   4
               Left            =   180
               TabIndex        =   12
               Top             =   7080
               Width           =   3555
            End
            Begin VB.Label LTempsRestantCycle 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               BorderStyle     =   1  'Fixed Single
               BeginProperty Font 
                  Name            =   "Microsoft Sans Serif"
                  Size            =   11.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FF0000&
               Height          =   315
               Index           =   4
               Left            =   180
               TabIndex        =   11
               Top             =   7680
               Width           =   2115
            End
            Begin VB.Label LTempsTotalCycle 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               BorderStyle     =   1  'Fixed Single
               BeginProperty Font 
                  Name            =   "Microsoft Sans Serif"
                  Size            =   11.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FF0000&
               Height          =   315
               Index           =   4
               Left            =   2280
               TabIndex        =   10
               Top             =   7680
               Width           =   1455
            End
            Begin VB.Label LLibelles 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H0080FF80&
               BorderStyle     =   1  'Fixed Single
               Caption         =   "PHASE"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   11.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   315
               Index           =   47
               Left            =   180
               TabIndex        =   9
               Top             =   6780
               Width           =   3555
            End
            Begin VB.Label LLibelles 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H0080FF80&
               BorderStyle     =   1  'Fixed Single
               Caption         =   "TEMPS RESTANT"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   11.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   315
               Index           =   45
               Left            =   180
               TabIndex        =   8
               Top             =   7380
               Width           =   2115
            End
            Begin VB.Label LLibelles 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H0080FF80&
               BorderStyle     =   1  'Fixed Single
               Caption         =   "TOTAL"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   11.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   315
               Index           =   44
               Left            =   2280
               TabIndex        =   7
               Top             =   7380
               Width           =   1455
            End
         End
         Begin VB.Label LNumBarres 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   315
            Index           =   24
            Left            =   23160
            TabIndex        =   153
            Top             =   4440
            Width           =   1155
         End
         Begin VB.Label LNumBarres 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   315
            Index           =   21
            Left            =   23160
            TabIndex        =   152
            Top             =   4140
            Width           =   1155
         End
         Begin VB.Label LNumBarres 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   315
            Index           =   20
            Left            =   23160
            TabIndex        =   151
            Top             =   3840
            Width           =   1155
         End
         Begin VB.Label LNumBarres 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   315
            Index           =   19
            Left            =   23160
            TabIndex        =   150
            Top             =   3540
            Width           =   1155
         End
         Begin VB.Label LNumBarres 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   315
            Index           =   18
            Left            =   23160
            TabIndex        =   149
            Top             =   3240
            Width           =   1155
         End
         Begin VB.Label LNumCharges 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   315
            Index           =   24
            Left            =   22020
            TabIndex        =   148
            Top             =   4440
            Width           =   1155
         End
         Begin VB.Label LNumCharges 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   315
            Index           =   21
            Left            =   22020
            TabIndex        =   147
            Top             =   4140
            Width           =   1155
         End
         Begin VB.Label LNumCharges 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   315
            Index           =   20
            Left            =   22020
            TabIndex        =   146
            Top             =   3840
            Width           =   1155
         End
         Begin VB.Label LNumCharges 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   315
            Index           =   19
            Left            =   22020
            TabIndex        =   145
            Top             =   3540
            Width           =   1155
         End
         Begin VB.Label LNumCharges 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   315
            Index           =   18
            Left            =   22020
            TabIndex        =   144
            Top             =   3240
            Width           =   1155
         End
         Begin VB.Label LLibelles 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "C19"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   315
            Index           =   41
            Left            =   20880
            TabIndex        =   143
            Top             =   4440
            Width           =   1155
         End
         Begin VB.Label LLibelles 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "C16"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   315
            Index           =   39
            Left            =   20880
            TabIndex        =   142
            Top             =   4140
            Width           =   1155
         End
         Begin VB.Label LLibelles 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "C15"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   315
            Index           =   38
            Left            =   20880
            TabIndex        =   141
            Top             =   3840
            Width           =   1155
         End
         Begin VB.Label LLibelles 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "C14"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   315
            Index           =   37
            Left            =   20880
            TabIndex        =   140
            Top             =   3540
            Width           =   1155
         End
         Begin VB.Label LLibelles 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "C13"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   315
            Index           =   36
            Left            =   20880
            TabIndex        =   139
            Top             =   3240
            Width           =   1155
         End
         Begin VB.Label LLibelles 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FF0000&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "N° de BARRE"
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
            Height          =   555
            Index           =   34
            Left            =   23160
            TabIndex        =   138
            Top             =   2700
            Width           =   1155
         End
         Begin VB.Label LLibelles 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FF0000&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "N° de CHARGE"
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
            Height          =   555
            Index           =   32
            Left            =   22020
            TabIndex        =   137
            Top             =   2700
            Width           =   1155
         End
         Begin VB.Label LLibelles 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FF0000&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Nom du POSTE"
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
            Height          =   555
            Index           =   29
            Left            =   20880
            TabIndex        =   136
            Top             =   2700
            Width           =   1155
         End
         Begin VB.Label LLibelles 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00004000&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "PHASES D'ANODISATION"
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
            Height          =   315
            Index           =   28
            Left            =   20880
            TabIndex        =   135
            Top             =   180
            Width           =   2925
         End
         Begin VB.Image IPhasesAnodisation 
            Height          =   2010
            Left            =   20880
            Picture         =   "FGestionRedresseurs.frx":25BE4
            Top             =   480
            Width           =   2925
         End
         Begin VB.Label LLibelles 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00004000&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "REDRESSEUR C13"
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
            Height          =   315
            Index           =   17
            Left            =   16740
            TabIndex        =   21
            Top             =   180
            Width           =   3975
         End
         Begin VB.Label LLibelles 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00004000&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "REDRESSEUR SPECTRO."
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
            Height          =   315
            Index           =   57
            Left            =   180
            TabIndex        =   20
            Top             =   180
            Width           =   3975
         End
         Begin VB.Label LLibelles 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00004000&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "REDRESSEUR C14"
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
            Height          =   315
            Index           =   0
            Left            =   12600
            TabIndex        =   5
            Top             =   180
            Width           =   3975
         End
         Begin VB.Label LLibelles 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00004000&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "REDRESSEUR C15"
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
            Height          =   315
            Index           =   6
            Left            =   8460
            TabIndex        =   4
            Top             =   180
            Width           =   3975
         End
         Begin VB.Label LLibelles 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00004000&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "REDRESSEUR C16"
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
            Height          =   315
            Index           =   7
            Left            =   4320
            TabIndex        =   3
            Top             =   180
            Width           =   3975
         End
      End
   End
End
Attribute VB_Name = "FGestionRedresseurs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Rôle                    : Fenêtre gérant les redresseurs
' Nom                    : FGestionRedresseurs.frm
' Date de création : 15/02/2011
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

    '--- divers sur la fenêtre ---
    With Me
        .Caption = UCase(TITRE_FENETRE)
        .WindowState = vbMaximized
    End With
    PBDeplacementFenetre(ZONES_DEPLACEMENT_FENETRE.Z_FILLE).Picture = ImgFondBleu2
    PBBoutons.Picture = ImgFondDesBoutons
    
    '--- renseignements de la fenêtre ---
    LRenseignementsFenetre.Caption = UCase(TITRE_FENETRE)
    
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
    With TimerEtatsRedresseurs
        .Enabled = False
        .Interval = 0
    End With
    
    '--- déchargement de la fenêtre ---
    Me.Visible = False
    Unload Me
    Set OccFGestionRedresseurs = Nothing
    
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
    Const CODE_ARRET_REDRESSEUR As String = "9"
    
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
            NomVariable = Choose(Index, "DemandesDuPCR1", "DemandesDuPCR2", "DemandesDuPCR3", "DemandesDuPCR4")
                    
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
            NomVariable = Choose(Index, "DemandesDuPCR1", "DemandesDuPCR2", "DemandesDuPCR3", "DemandesDuPCR4")
                    
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
    Const CODE_MARCHE_REDRESSEUR As String = "10"
    
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
            NomVariable = Choose(Index, "DemandesDuPCR1", "DemandesDuPCR2", "DemandesDuPCR3", "DemandesDuPCR4")
                    
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
    Dim NumCharge As Integer                                               'numéro de charge
    Dim AjoutTempsTotalCycleSecondes As Integer              'représente l'ajout au temps total en secondes
    Dim NouveauTempsPhase4 As Integer                            'représente le nouveau temps de la phase 4
    
    Dim NouvelleIntensite As Long                                         'représente la nouvelle intensité
    Dim ValeurRetourneeAPI As Long                                      'valeur retournée par une fonction concernant le dialogue avec l'automate
    
    Dim NouvelleTension As Single                                        'représente la nouvelle tension
    
    Dim NouvelleTensionTexte As String                                'représente la nouvelle tension en format texte
    Dim NouvelleIntensiteTexte As String                               'représente la nouvelle intensité en format texte
    
    Dim AjoutTempsTotalCycleTexte As String                       'représente l'ajout de temps total en format texte
    Dim NomGroupe As String                                                 'représente un nom de groupe
    Dim NomElement As String                                                'représente un nom d'élément (variable nommée)
    
    '--- affectation du numéro de charge ---
    NumCharge = TEtatsRedresseurs(Index).NumCharge

    '--- analyse en fonction du numéro de charge ---
    If NumCharge >= CHARGES.C_NUM_MINI And NumCharge <= CHARGES.C_NUM_MAXI Then

        '--- calcul de la valeur du temps pour la dernière couche ---
        With TEtatsCharges(NumCharge)

                
            '--- affectation du nouveau temps en secondes tapé dans le champ d'édition ---
            AjoutTempsTotalCycleTexte = MEBTempsTotalCycle(Index).Text
            AjoutTempsTotalCycleTexte = Replace(AjoutTempsTotalCycleTexte, "_", "0")
            If AjoutTempsTotalCycleTexte = "" Then AjoutTempsTotalCycleTexte = "00:00"
            
            '--- affectation en numérique / affectation du signe ---
            AjoutTempsTotalCycleSecondes = CInt(Left(AjoutTempsTotalCycleTexte, 2)) * 60 + CInt(Right(AjoutTempsTotalCycleTexte, 2))
            If CBSensAjoutTempsDeBain(Index).Caption = TEXTE_EN_MOINS Then
                AjoutTempsTotalCycleSecondes = -AjoutTempsTotalCycleSecondes
            End If
    
            '--- calcul du nouveau temps à faire ---
            NouveauTempsPhase4 = .TDetailsPhasesProduction(PHASES_GAMMES_REDRESSEURS.PH_T4).TempsPhase + AjoutTempsTotalCycleSecondes
            
            '-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
                    
            '--- transfert des nouvelles valeurs dans l'automate ---
            If AppelFenetre(F_MESSAGE, _
                                     TITRE_MESSAGES, _
                                     MESSAGE_4, _
                                     TYPES_MESSAGES.T_ATTENTION, _
                                     TYPES_BOUTONS.T_OUI_NON, _
                                     EMPLACEMENT_FOCUS.E_SUR_NON) = vbYes Then
            
                '--- affectation du groupe ---
                NomGroupe = "CHARGE_" & Right("00" & NumCharge, 2)
        
                '*************************************************************************************************************************************
                
                If AjoutTempsTotalCycleTexte <> "00:00" Then
                
                    '--- transfert dans l'automate du nouveau temps de la phase 4 ---
                    ValeurRetourneeAPI = APIEcritureVariableNommee(NomGroupe, "TpsPhase4", NouveauTempsPhase4)
                    If ValeurRetourneeAPI <> 0 Then
                        Bidon = MessageErreur(TITRE_MESSAGES, MESSAGE_500)
                    End If
                    
                End If
                        
                '*************************************************************************************************************************************
            
                '--- affectation de la nouvelle tension au format texte ---
                NouvelleTensionTexte = MEBTension(Index).Text
                NouvelleTensionTexte = Replace(NouvelleTensionTexte, "_", "")
                
                If NouvelleTensionTexte <> "" Then
                
                    '--- analyse du champ tension ---
                    If IsNumeric(NouvelleTensionTexte) = True Then
                
                        '--- affectation de la tension ---
                        NouvelleTension = CSng(NouvelleTensionTexte)
                
                        If NouvelleTension >= 0 And NouvelleTension <= TEtatsRedresseurs(Index).DefinitionRedresseur.UMaxiRedresseur Then
                    
                            '--- tension de la phase 4 ---
                            NomElement = "UPhase4"
                            ValeurRetourneeAPI = APIEcritureVariableNommee(NomGroupe, NomElement, CInt(NouvelleTension * 10))
                            If ValeurRetourneeAPI <> 0 Then
                                Bidon = MessageErreur(TITRE_MESSAGES, MESSAGE_500)
                            End If

                        Else

                            '--- lancer un message d'erreur ---
                            Bidon = MessageErreur(TITRE_MESSAGES, vbCrLf & vbCrLf & vbCrLf & "cs|La valeur de la tension n'est pas conforme.")

                        End If
                    
                    Else
                        
                        '--- lancer un message d'erreur ---
                        Bidon = MessageErreur(TITRE_MESSAGES, vbCrLf & vbCrLf & vbCrLf & _
                                                                                                "cs|La valeur de la tension doit être supérieure ou" & vbCrLf & _
                                                                                                "cs|égale à 0 et inférieure ou égale à " & TEtatsRedresseurs(Index).DefinitionRedresseur.UMaxiRedresseur & " V")
                
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
                    
                            '--- intensité de la phase 4 ---
                            NomElement = "IPhase4"
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
    PBDeplacementFenetre(ZONES_DEPLACEMENT_FENETRE.Z_MERE).Height = Abs(Me.ScaleHeight - PBRenseignementsFenetre.Height - PBBoutons.Height)
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

Private Sub MEBTension_Change(Index As Integer)
    On Error Resume Next
    If InterdireEvenements = False Then
        CBAnnulerTransfertModifications(Index).Enabled = True
        CBTransfertModificationsVersAPI(Index).Enabled = True
    End If
End Sub

Private Sub MEBTension_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    On Error Resume Next
    FiltreToucheFonction KeyCode, Shift
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
        
    If Index = ZONES_DEPLACEMENT_FENETRE.Z_MERE Then

        If Me.WindowState = vbMaximized Then
            
            '--- agrandir la zone fille ---
            With PBDeplacementFenetre(ZONES_DEPLACEMENT_FENETRE.Z_FILLE)
                
                .Left = PBDeplacementFenetre(ZONES_DEPLACEMENT_FENETRE.Z_MERE).ScaleLeft
                .Top = PBDeplacementFenetre(ZONES_DEPLACEMENT_FENETRE.Z_MERE).ScaleTop
                .Height = PBDeplacementFenetre(ZONES_DEPLACEMENT_FENETRE.Z_MERE).ScaleHeight
                .Width = PBDeplacementFenetre(ZONES_DEPLACEMENT_FENETRE.Z_MERE).ScaleWidth
            
                '--- agrandir en proportion de la zone fille ---
            
            End With
                   
        End If

    Else

    End If
            
    '--- valeur des curseurs ---
    If Me.WindowState <> vbMaximized And Me.WindowState <> vbMinimized Then
        HSDeplacementFenetre.Max = PBDeplacementFenetre(ZONES_DEPLACEMENT_FENETRE.Z_FILLE).Width - _
                                                         PBDeplacementFenetre(ZONES_DEPLACEMENT_FENETRE.Z_MERE).Width
        VSDeplacementFenetre.Max = PBDeplacementFenetre(ZONES_DEPLACEMENT_FENETRE.Z_FILLE).Height - _
                                                         PBDeplacementFenetre(ZONES_DEPLACEMENT_FENETRE.Z_MERE).Height
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
    PBDeplacementFenetre(ZONES_DEPLACEMENT_FENETRE.Z_FILLE).Top = -VSDeplacementFenetre.Value
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
    For a = POSTES.P_C13 To POSTES.P_C19
        
        Select Case a
        
            Case POSTES.P_C13 To POSTES.P_C16, POSTES.P_C19
                '--- postes concernés ---
                With TEtatsPostes(a)
        
                    If .Condamnation = True Then
                                
                        '--- affichage de la croix de condamnation ---
                        Texte = "X"
                        AffichageTexte LNumCharges(a), Texte, COULEURS.BLANC, COULEURS.ROUGE_3
                        AffichageTexte LNumBarres(a), Texte, COULEURS.BLANC, COULEURS.ROUGE_3
                        
                    Else
                        
                        '--- affichage du numéro des charges ---
                        If .NumCharge >= CHARGES.C_NUM_MINI And .NumCharge <= CHARGES.C_NUM_MAXI Then
                        
                            '--- affichage du numéro de charge ---
                            Texte = CStr(.NumCharge)
                            AffichageTexte LNumCharges(a), Texte, COULEURS.JAUNE_2, COULEURS.NOIR
                            
                            '--- affichage du numéro de barre ---
                            Texte = CStr(TEtatsCharges(.NumCharge).NumBarre)
                            AffichageTexte LNumBarres(a), Texte, COULEURS.VERT_2, COULEURS.NOIR
                                    
                        Else
                            
                            '--- vider le champ ---
                            AffichageTexte LNumCharges(a), "", COULEURS.BLANC, COULEURS.NOIR
                            AffichageTexte LNumBarres(a), "", COULEURS.BLANC, COULEURS.NOIR
                        
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

    For a = REDRESSEURS.R_C13 To REDRESSEURS.R_C19
    
        With TEtatsRedresseurs(a)
        
            '***********************************************************************************************************************************************
            '                                                                               SUR LE DESSIN DU REDRESSEUR
            '***********************************************************************************************************************************************
            
            '--- mode du redresseur ---
            Select Case .ModeRedresseur
                Case MODES_REDRESSEUR.MR_MANUEL: OCXRedresseurs(a).Mode = MODE_MANUEL
                Case MODES_REDRESSEUR.MR_AUTOMATIQUE: OCXRedresseurs(a).Mode = MODE_AUTOMATIQUE
                Case Else: OCXRedresseurs(a).Mode = MODE_NON_DEFINI
            End Select
            
            '--- tension ---
            If .EtatRedresseur = ER_ARRET Then
                OCXRedresseurs(a).Tension = 0
            Else
                OCXRedresseurs(a).Tension = .U
            End If
            
            '--- intensité ---
            If .EtatRedresseur = ER_ARRET Then
                OCXRedresseurs(a).Intensite = 0
            Else
                OCXRedresseurs(a).Intensite = .I
            End If
            
            '--- ah ---
            OCXRedresseurs(a).Ah = .Ah
            
            '--- sens ---
            OCXRedresseurs(a).Sens = .SensRedresseur

            '--- temps restant de la phase (99:59 possible) ---
            If .TempsPhaseEnCours > 0 And .TempsEcoulePhaseEnCours > 0 Then
                OCXRedresseurs(a).TempsRestantPhase = CTemps3(Abs(.TempsPhaseEnCours - .TempsEcoulePhaseEnCours))
            Else
                OCXRedresseurs(a).TempsRestantPhase = "-"
            End If
            
            '--- vu-mètre de la phase en cours ---
            Select Case .NumPhaseEnCours
                Case PHASES_GAMMES_REDRESSEURS.PH_T1 To PHASES_GAMMES_REDRESSEURS.PH_T4: OCXRedresseurs(a).Phase = .NumPhaseEnCours
                Case Else: OCXRedresseurs(a).Phase = ETEINT
            End Select
            
            '--- état du redresseur ---
            Select Case .EtatRedresseur
                Case ETATS_REDRESSEUR.ER_ARRET To ETATS_REDRESSEUR.ER_EXCLUSION: OCXRedresseurs(a).Etat = .EtatRedresseur
                Case Else:  OCXRedresseurs(a).Etat = ETAT_NON_DEFINI
            End Select
                
            '***********************************************************************************************************************************************
            '                                                                               SUR LES CHAMPS D'AFFICHAGE
            '***********************************************************************************************************************************************
                
            '--- temps total du cycle ---
            Texte = CTemps(.TempsTotalCycle)
            AffichageTexte LTempsTotalCycle(a), Texte
            
            '--- temps restant du cycle ---
            Texte = CTemps(.TempsRestantCycle)
            AffichageTexte LTempsRestantCycle(a), Texte
                
            '--- phase en cours ---
            Select Case .NumPhaseEnCours
                Case PHASES_GAMMES_REDRESSEURS.PH_T1 To PHASES_GAMMES_REDRESSEURS.PH_T4
                    Texte = "Phase " & .NumPhaseEnCours & " - " & CTemps(.TempsEcoulePhaseEnCours) & " / " & CTemps(.TempsPhaseEnCours)
                Case Else
                    Texte = "-"
            End Select
            AffichageTexte LPhaseEnCours(a), Texte
        
        End With

    Next a

End Sub


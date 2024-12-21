VERSION 5.00
Object = "{0AFE7BE0-11B7-4A3E-978D-D4501E9A57FE}#1.0#0"; "c1sizer.ocx"
Begin VB.Form FProgrammateurCyclique 
   BackColor       =   &H00C0C0C0&
   ClientHeight    =   13005
   ClientLeft      =   2100
   ClientTop       =   705
   ClientWidth     =   13395
   Icon            =   "FProgrammateurCyclique.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   NegotiateMenus  =   0   'False
   ScaleHeight     =   13005
   ScaleWidth      =   13395
   ShowInTaskbar   =   0   'False
   WindowState     =   2  'Maximized
   Begin VB.PictureBox PBBoutons 
      Align           =   2  'Align Bottom
      DrawStyle       =   2  'Dot
      DrawWidth       =   16891
      Height          =   1095
      Left            =   0
      ScaleHeight     =   1035
      ScaleWidth      =   13335
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   11910
      Width           =   13395
      Begin VB.Timer TimerSortieObligatoire 
         Enabled         =   0   'False
         Interval        =   60000
         Left            =   2160
         Top             =   120
      End
      Begin VB.Timer TimerProgCyclique 
         Enabled         =   0   'False
         Interval        =   1000
         Left            =   1620
         Top             =   120
      End
      Begin VB.CommandButton CBQuitter 
         BackColor       =   &H00FFFFFF&
         Cancel          =   -1  'True
         Caption         =   "&Quitter"
         DownPicture     =   "FProgrammateurCyclique.frx":014A
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
         Left            =   15600
         MaskColor       =   &H00FF00FF&
         Picture         =   "FProgrammateurCyclique.frx":084C
         Style           =   1  'Graphical
         TabIndex        =   14
         ToolTipText     =   " Quitter cette fenêtre "
         Top             =   120
         UseMaskColor    =   -1  'True
         Width           =   1515
      End
      Begin VB.CommandButton CBValider 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Valider"
         DownPicture     =   "FProgrammateurCyclique.frx":0F4E
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
         Left            =   13860
         MaskColor       =   &H00FF00FF&
         Picture         =   "FProgrammateurCyclique.frx":1650
         Style           =   1  'Graphical
         TabIndex        =   13
         ToolTipText     =   " Valider l'enregistrement "
         Top             =   120
         UseMaskColor    =   -1  'True
         Width           =   1515
      End
      Begin VB.CommandButton CBAnnuler 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Annuler"
         DownPicture     =   "FProgrammateurCyclique.frx":1D52
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
         Left            =   12120
         MaskColor       =   &H00FF00FF&
         Picture         =   "FProgrammateurCyclique.frx":2454
         Style           =   1  'Graphical
         TabIndex        =   12
         ToolTipText     =   " Annuler les dernières modifications "
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
         TabIndex        =   8
         Top             =   0
         Visible         =   0   'False
         Width           =   1215
         Begin VB.HScrollBar HSDeplacementFenetre 
            Height          =   255
            LargeChange     =   300
            Left            =   0
            SmallChange     =   100
            TabIndex        =   11
            Top             =   720
            Width           =   915
         End
         Begin VB.VScrollBar VSDeplacementFenetre 
            Height          =   975
            LargeChange     =   300
            Left            =   900
            SmallChange     =   100
            TabIndex        =   10
            Top             =   0
            Width           =   255
         End
         Begin VB.CommandButton CBAgrandirFenetre 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Agrandir"
            DownPicture     =   "FProgrammateurCyclique.frx":2B56
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
            Picture         =   "FProgrammateurCyclique.frx":2D00
            Style           =   1  'Graphical
            TabIndex        =   9
            ToolTipText     =   " Agrandissement de la fenêtre "
            Top             =   0
            UseMaskColor    =   -1  'True
            Width           =   900
         End
      End
      Begin VB.Shape SFocus 
         BorderColor     =   &H000000FF&
         BorderWidth     =   5
         Height          =   285
         Left            =   4860
         Top             =   180
         Visible         =   0   'False
         Width           =   780
      End
   End
   Begin VB.PictureBox PBRenseignementsFenetre 
      Align           =   1  'Align Top
      BackColor       =   &H00FF0000&
      Height          =   375
      Left            =   0
      Picture         =   "FProgrammateurCyclique.frx":2EAA
      ScaleHeight     =   315
      ScaleWidth      =   13335
      TabIndex        =   0
      Top             =   0
      Width           =   13395
      Begin VB.Label LRenseignementsFenetre 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00008000&
         BackStyle       =   0  'Transparent
         Caption         =   "JOUR GERE"
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
         TabIndex        =   1
         Top             =   0
         Width           =   11415
         WordWrap        =   -1  'True
      End
   End
   Begin VB.PictureBox PBDeplacementFenetre 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   13395
      Index           =   0
      Left            =   0
      ScaleHeight     =   13395
      ScaleWidth      =   13395
      TabIndex        =   2
      Top             =   375
      Width           =   13395
      Begin VB.PictureBox PBDeplacementFenetre 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   13215
         Index           =   1
         Left            =   0
         ScaleHeight     =   13215
         ScaleWidth      =   28410
         TabIndex        =   3
         Top             =   0
         Width           =   28410
         Begin VB.PictureBox PBModesProgrammation 
            Height          =   12855
            Left            =   21900
            ScaleHeight     =   12795
            ScaleWidth      =   5775
            TabIndex        =   5
            Top             =   0
            Width           =   5835
            Begin VB.PictureBox PBTousModesChauffage 
               Height          =   2655
               Left            =   420
               ScaleHeight     =   2595
               ScaleWidth      =   4155
               TabIndex        =   30
               Top             =   720
               Width           =   4215
               Begin VB.OptionButton OBModesChauffage 
                  BackColor       =   &H00C0E0FF&
                  Caption         =   "PRODUCTION"
                  BeginProperty Font 
                     Name            =   "Arial"
                     Size            =   9.75
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   435
                  Index           =   2
                  Left            =   1140
                  Style           =   1  'Graphical
                  TabIndex        =   36
                  Top             =   1800
                  Width           =   2655
               End
               Begin VB.PictureBox PBModesChauffage 
                  Appearance      =   0  'Flat
                  BackColor       =   &H00FFFFFF&
                  ForeColor       =   &H000080FF&
                  Height          =   435
                  Index           =   2
                  Left            =   360
                  ScaleHeight     =   405
                  ScaleWidth      =   585
                  TabIndex        =   35
                  TabStop         =   0   'False
                  Top             =   1800
                  Width           =   615
               End
               Begin VB.OptionButton OBModesChauffage 
                  BackColor       =   &H00C0E0FF&
                  Caption         =   "VEILLE"
                  BeginProperty Font 
                     Name            =   "Arial"
                     Size            =   9.75
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   435
                  Index           =   1
                  Left            =   1140
                  Style           =   1  'Graphical
                  TabIndex        =   34
                  Top             =   1080
                  Width           =   2655
               End
               Begin VB.PictureBox PBModesChauffage 
                  Appearance      =   0  'Flat
                  BackColor       =   &H00FFFFFF&
                  ForeColor       =   &H00FFFF00&
                  Height          =   435
                  Index           =   1
                  Left            =   360
                  ScaleHeight     =   405
                  ScaleWidth      =   585
                  TabIndex        =   33
                  TabStop         =   0   'False
                  Top             =   1080
                  Width           =   615
               End
               Begin VB.OptionButton OBModesChauffage 
                  BackColor       =   &H00C0E0FF&
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
                  Height          =   435
                  Index           =   0
                  Left            =   1140
                  MaskColor       =   &H00FFFFFF&
                  Style           =   1  'Graphical
                  TabIndex        =   32
                  Top             =   360
                  Width           =   2655
               End
               Begin VB.PictureBox PBModesChauffage 
                  Appearance      =   0  'Flat
                  BackColor       =   &H00FFFFFF&
                  ForeColor       =   &H00FFFFFF&
                  Height          =   435
                  Index           =   0
                  Left            =   360
                  ScaleHeight     =   405
                  ScaleWidth      =   585
                  TabIndex        =   31
                  TabStop         =   0   'False
                  Top             =   360
                  Width           =   615
               End
            End
            Begin VB.PictureBox PBTousCyclesPompe 
               Height          =   1875
               Left            =   420
               ScaleHeight     =   1815
               ScaleWidth      =   4155
               TabIndex        =   25
               Top             =   4140
               Width           =   4215
               Begin VB.OptionButton OBCyclesPompe 
                  BackColor       =   &H00C0E0FF&
                  Caption         =   "Marche"
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   9.75
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   435
                  Index           =   1
                  Left            =   1080
                  Style           =   1  'Graphical
                  TabIndex        =   29
                  Top             =   1020
                  Width           =   2655
               End
               Begin VB.OptionButton OBCyclesPompe 
                  BackColor       =   &H00C0E0FF&
                  Caption         =   "Arrêt"
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   9.75
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   435
                  Index           =   0
                  Left            =   1080
                  Style           =   1  'Graphical
                  TabIndex        =   28
                  Top             =   300
                  Width           =   2655
               End
               Begin VB.PictureBox PBCyclesPompe 
                  Appearance      =   0  'Flat
                  BackColor       =   &H00FFFFFF&
                  ForeColor       =   &H000080FF&
                  Height          =   435
                  Index           =   1
                  Left            =   360
                  ScaleHeight     =   405
                  ScaleWidth      =   525
                  TabIndex        =   27
                  TabStop         =   0   'False
                  Top             =   1020
                  Width           =   555
               End
               Begin VB.PictureBox PBCyclesPompe 
                  Appearance      =   0  'Flat
                  BackColor       =   &H00FFFFFF&
                  ForeColor       =   &H00FFFFFF&
                  Height          =   435
                  Index           =   0
                  Left            =   360
                  ScaleHeight     =   405
                  ScaleWidth      =   525
                  TabIndex        =   26
                  TabStop         =   0   'False
                  Top             =   300
                  Width           =   555
               End
            End
            Begin VB.Label LTitreTousModesChauffage 
               Alignment       =   2  'Center
               BackColor       =   &H00800000&
               BorderStyle     =   1  'Fixed Single
               Caption         =   "MODES D'UN CHAUFFAGE"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   12
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FFFFFF&
               Height          =   375
               Left            =   420
               TabIndex        =   38
               Top             =   360
               Width           =   4215
            End
            Begin VB.Label LTitreTousCyclesPompe 
               Alignment       =   2  'Center
               BackColor       =   &H00800000&
               BorderStyle     =   1  'Fixed Single
               Caption         =   "MODES D'UNE POMPE"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   12
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FFFFFF&
               Height          =   375
               Left            =   420
               TabIndex        =   37
               Top             =   3780
               Width           =   4215
            End
            Begin VB.Label LAvertissement 
               Appearance      =   0  'Flat
               BackColor       =   &H000000FF&
               BorderStyle     =   1  'Fixed Single
               Caption         =   "           ATTENTION                 Changement de journée"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   13.5
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H0000FFFF&
               Height          =   795
               Left            =   600
               TabIndex        =   6
               Top             =   11460
               Visible         =   0   'False
               Width           =   4035
            End
         End
         Begin VB.PictureBox PBProgrammation 
            Height          =   12855
            Left            =   0
            ScaleHeight     =   12795
            ScaleWidth      =   21855
            TabIndex        =   4
            Top             =   0
            Width           =   21915
            Begin VB.PictureBox PBModeGeneralCuves 
               Height          =   2595
               Left            =   300
               ScaleHeight     =   2535
               ScaleWidth      =   5235
               TabIndex        =   80
               Top             =   9780
               Width           =   5295
               Begin VB.CommandButton CBModeGeneralCuves 
                  BackColor       =   &H00C0FFC0&
                  Caption         =   "Passage de toutes les cuves en TRAVAIL"
                  BeginProperty Font 
                     Name            =   "Arial"
                     Size            =   9.75
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   495
                  Index           =   1
                  Left            =   180
                  Style           =   1  'Graphical
                  TabIndex        =   84
                  Top             =   1920
                  Width           =   4875
               End
               Begin VB.CommandButton CBModeGeneralCuves 
                  BackColor       =   &H00C0FFC0&
                  Caption         =   "Passage de toutes les cuves en REPRISE"
                  BeginProperty Font 
                     Name            =   "Arial"
                     Size            =   9.75
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   495
                  Index           =   3
                  Left            =   180
                  Style           =   1  'Graphical
                  TabIndex        =   83
                  Top             =   1320
                  Width           =   4875
               End
               Begin VB.CommandButton CBModeGeneralCuves 
                  BackColor       =   &H00C0FFC0&
                  Caption         =   "Passage de toutes les cuves en VEILLE"
                  BeginProperty Font 
                     Name            =   "Arial"
                     Size            =   9.75
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   495
                  Index           =   2
                  Left            =   180
                  Style           =   1  'Graphical
                  TabIndex        =   82
                  Top             =   720
                  Width           =   4875
               End
               Begin VB.CommandButton CBModeGeneralCuves 
                  BackColor       =   &H00C0FFC0&
                  Caption         =   "Passage de toutes les cuves en ARRET"
                  BeginProperty Font 
                     Name            =   "Arial"
                     Size            =   9.75
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   495
                  Index           =   0
                  Left            =   180
                  Style           =   1  'Graphical
                  TabIndex        =   81
                  Top             =   120
                  Width           =   4875
               End
            End
            Begin VB.PictureBox PBJours 
               Height          =   8415
               Left            =   300
               ScaleHeight     =   8355
               ScaleWidth      =   5235
               TabIndex        =   62
               Top             =   720
               Width           =   5295
               Begin VB.OptionButton OBJours 
                  Appearance      =   0  'Flat
                  BackColor       =   &H00C0E0FF&
                  Caption         =   "-"
                  BeginProperty Font 
                     Name            =   "Arial"
                     Size            =   9.75
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00000000&
                  Height          =   495
                  Index           =   15
                  Left            =   180
                  MaskColor       =   &H00FF00FF&
                  Style           =   1  'Graphical
                  TabIndex        =   78
                  Top             =   7680
                  UseMaskColor    =   -1  'True
                  Width           =   4575
               End
               Begin VB.OptionButton OBJours 
                  Appearance      =   0  'Flat
                  BackColor       =   &H00C0E0FF&
                  Caption         =   "-"
                  BeginProperty Font 
                     Name            =   "Arial"
                     Size            =   9.75
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00000000&
                  Height          =   495
                  Index           =   14
                  Left            =   180
                  MaskColor       =   &H00FF00FF&
                  Style           =   1  'Graphical
                  TabIndex        =   77
                  Top             =   7140
                  UseMaskColor    =   -1  'True
                  Width           =   4575
               End
               Begin VB.OptionButton OBJours 
                  Appearance      =   0  'Flat
                  BackColor       =   &H00C0E0FF&
                  Caption         =   "-"
                  BeginProperty Font 
                     Name            =   "Arial"
                     Size            =   9.75
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00000000&
                  Height          =   495
                  Index           =   13
                  Left            =   180
                  MaskColor       =   &H00FF00FF&
                  Style           =   1  'Graphical
                  TabIndex        =   76
                  Top             =   6600
                  UseMaskColor    =   -1  'True
                  Width           =   4575
               End
               Begin VB.OptionButton OBJours 
                  Appearance      =   0  'Flat
                  BackColor       =   &H00C0E0FF&
                  Caption         =   "-"
                  BeginProperty Font 
                     Name            =   "Arial"
                     Size            =   9.75
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00000000&
                  Height          =   495
                  Index           =   12
                  Left            =   180
                  MaskColor       =   &H00FF00FF&
                  Style           =   1  'Graphical
                  TabIndex        =   75
                  Top             =   6060
                  UseMaskColor    =   -1  'True
                  Width           =   4575
               End
               Begin VB.OptionButton OBJours 
                  Appearance      =   0  'Flat
                  BackColor       =   &H00C0E0FF&
                  Caption         =   "-"
                  BeginProperty Font 
                     Name            =   "Arial"
                     Size            =   9.75
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00000000&
                  Height          =   495
                  Index           =   11
                  Left            =   180
                  MaskColor       =   &H00FF00FF&
                  Style           =   1  'Graphical
                  TabIndex        =   74
                  Top             =   5520
                  UseMaskColor    =   -1  'True
                  Width           =   4575
               End
               Begin VB.OptionButton OBJours 
                  Appearance      =   0  'Flat
                  BackColor       =   &H00C0E0FF&
                  Caption         =   "-"
                  BeginProperty Font 
                     Name            =   "Arial"
                     Size            =   9.75
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00000000&
                  Height          =   495
                  Index           =   10
                  Left            =   180
                  MaskColor       =   &H00FF00FF&
                  Style           =   1  'Graphical
                  TabIndex        =   73
                  Top             =   4980
                  UseMaskColor    =   -1  'True
                  Width           =   4575
               End
               Begin VB.OptionButton OBJours 
                  Appearance      =   0  'Flat
                  BackColor       =   &H00C0E0FF&
                  Caption         =   "-"
                  BeginProperty Font 
                     Name            =   "Arial"
                     Size            =   9.75
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00000000&
                  Height          =   495
                  Index           =   9
                  Left            =   180
                  MaskColor       =   &H00FF00FF&
                  Style           =   1  'Graphical
                  TabIndex        =   72
                  Top             =   4440
                  UseMaskColor    =   -1  'True
                  Width           =   4575
               End
               Begin VB.OptionButton OBJours 
                  Appearance      =   0  'Flat
                  BackColor       =   &H00C0E0FF&
                  Caption         =   "-"
                  BeginProperty Font 
                     Name            =   "Arial"
                     Size            =   9.75
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00000000&
                  Height          =   495
                  Index           =   8
                  Left            =   180
                  MaskColor       =   &H00FF00FF&
                  Style           =   1  'Graphical
                  TabIndex        =   71
                  Top             =   3900
                  UseMaskColor    =   -1  'True
                  Width           =   4575
               End
               Begin VB.OptionButton OBJours 
                  Appearance      =   0  'Flat
                  BackColor       =   &H00C0E0FF&
                  Caption         =   "-"
                  BeginProperty Font 
                     Name            =   "Arial"
                     Size            =   9.75
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00000000&
                  Height          =   495
                  Index           =   7
                  Left            =   180
                  MaskColor       =   &H00FF00FF&
                  Style           =   1  'Graphical
                  TabIndex        =   70
                  Top             =   3360
                  UseMaskColor    =   -1  'True
                  Width           =   4575
               End
               Begin VB.OptionButton OBJours 
                  Appearance      =   0  'Flat
                  BackColor       =   &H00C0E0FF&
                  Caption         =   "-"
                  BeginProperty Font 
                     Name            =   "Arial"
                     Size            =   9.75
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00000000&
                  Height          =   495
                  Index           =   6
                  Left            =   180
                  MaskColor       =   &H00FF00FF&
                  Style           =   1  'Graphical
                  TabIndex        =   69
                  Top             =   2820
                  UseMaskColor    =   -1  'True
                  Width           =   4575
               End
               Begin VB.OptionButton OBJours 
                  Appearance      =   0  'Flat
                  BackColor       =   &H00C0E0FF&
                  Caption         =   "-"
                  BeginProperty Font 
                     Name            =   "Arial"
                     Size            =   9.75
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00000000&
                  Height          =   495
                  Index           =   5
                  Left            =   180
                  MaskColor       =   &H00FF00FF&
                  Style           =   1  'Graphical
                  TabIndex        =   68
                  Top             =   2280
                  UseMaskColor    =   -1  'True
                  Width           =   4575
               End
               Begin VB.OptionButton OBJours 
                  Appearance      =   0  'Flat
                  BackColor       =   &H00C0E0FF&
                  Caption         =   "-"
                  BeginProperty Font 
                     Name            =   "Arial"
                     Size            =   9.75
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00000000&
                  Height          =   495
                  Index           =   4
                  Left            =   180
                  MaskColor       =   &H00FF00FF&
                  Style           =   1  'Graphical
                  TabIndex        =   67
                  Top             =   1740
                  UseMaskColor    =   -1  'True
                  Width           =   4575
               End
               Begin VB.OptionButton OBJours 
                  Appearance      =   0  'Flat
                  BackColor       =   &H00C0E0FF&
                  Caption         =   "-"
                  BeginProperty Font 
                     Name            =   "Arial"
                     Size            =   9.75
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00000000&
                  Height          =   495
                  Index           =   3
                  Left            =   180
                  MaskColor       =   &H00FF00FF&
                  Style           =   1  'Graphical
                  TabIndex        =   66
                  Top             =   1200
                  UseMaskColor    =   -1  'True
                  Width           =   4575
               End
               Begin VB.OptionButton OBJours 
                  Appearance      =   0  'Flat
                  BackColor       =   &H00C0E0FF&
                  Caption         =   "-"
                  BeginProperty Font 
                     Name            =   "Arial"
                     Size            =   9.75
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00000000&
                  Height          =   495
                  Index           =   2
                  Left            =   180
                  MaskColor       =   &H00FF00FF&
                  Style           =   1  'Graphical
                  TabIndex        =   65
                  Top             =   660
                  UseMaskColor    =   -1  'True
                  Width           =   4575
               End
               Begin VB.OptionButton OBJours 
                  Appearance      =   0  'Flat
                  BackColor       =   &H00C0E0FF&
                  Caption         =   "-"
                  BeginProperty Font 
                     Name            =   "Arial"
                     Size            =   9.75
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00000000&
                  Height          =   495
                  Index           =   1
                  Left            =   180
                  MaskColor       =   &H00FF00FF&
                  Style           =   1  'Graphical
                  TabIndex        =   64
                  Top             =   120
                  UseMaskColor    =   -1  'True
                  Width           =   4575
               End
               Begin VB.VScrollBar VSJours 
                  Height          =   5715
                  Left            =   4920
                  Max             =   15
                  Min             =   1
                  TabIndex        =   63
                  Top             =   120
                  Value           =   1
                  Width           =   315
               End
            End
            Begin C1SizerLibCtl.C1Tab CTOnglets 
               Height          =   12015
               Left            =   5940
               TabIndex        =   15
               Top             =   360
               Width           =   15555
               _cx             =   27437
               _cy             =   21193
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Arial"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Enabled         =   -1  'True
               Appearance      =   3
               MousePointer    =   0
               Version         =   801
               BackColor       =   -2147483633
               ForeColor       =   -2147483630
               FrontTabColor   =   65535
               BackTabColor    =   -2147483643
               TabOutlineColor =   0
               FrontTabForeColor=   0
               Caption         =   "Préparation|Anodisation|Coloration / fin de ligne"
               Align           =   0
               CurrTab         =   2
               FirstTab        =   0
               Style           =   1
               Position        =   0
               AutoSwitch      =   -1  'True
               AutoScroll      =   -1  'True
               TabPreview      =   -1  'True
               ShowFocusRect   =   0   'False
               TabsPerPage     =   3
               BorderWidth     =   0
               BoldCurrent     =   -1  'True
               DogEars         =   -1  'True
               MultiRow        =   0   'False
               MultiRowOffset  =   0
               CaptionStyle    =   0
               TabHeight       =   450
               TabCaptionPos   =   4
               TabPicturePos   =   1
               CaptionEmpty    =   ""
               Separators      =   0   'False
               AccessibleName  =   ""
               AccessibleDescription=   ""
               AccessibleValue =   ""
               AccessibleRole  =   37
               Begin VB.PictureBox PBOnglets 
                  Height          =   11475
                  Index           =   0
                  Left            =   -16410
                  ScaleHeight     =   11415
                  ScaleWidth      =   15405
                  TabIndex        =   39
                  Top             =   495
                  Width           =   15465
                  Begin VB.PictureBox PBEchelle24H 
                     Appearance      =   0  'Flat
                     AutoRedraw      =   -1  'True
                     AutoSize        =   -1  'True
                     BackColor       =   &H80000005&
                     BorderStyle     =   0  'None
                     ForeColor       =   &H80000008&
                     Height          =   945
                     Index           =   1
                     Left            =   240
                     Picture         =   "FProgrammateurCyclique.frx":277EC
                     ScaleHeight     =   63
                     ScaleMode       =   3  'Pixel
                     ScaleWidth      =   661
                     TabIndex        =   49
                     TabStop         =   0   'False
                     Top             =   600
                     Width           =   9915
                  End
                  Begin VB.PictureBox PBEchelle24H 
                     Appearance      =   0  'Flat
                     AutoRedraw      =   -1  'True
                     AutoSize        =   -1  'True
                     BackColor       =   &H80000005&
                     BorderStyle     =   0  'None
                     ForeColor       =   &H80000008&
                     Height          =   945
                     Index           =   2
                     Left            =   240
                     Picture         =   "FProgrammateurCyclique.frx":4606E
                     ScaleHeight     =   63
                     ScaleMode       =   3  'Pixel
                     ScaleWidth      =   661
                     TabIndex        =   48
                     TabStop         =   0   'False
                     Top             =   2160
                     Width           =   9915
                  End
                  Begin VB.PictureBox PBEchelle24H 
                     Appearance      =   0  'Flat
                     AutoRedraw      =   -1  'True
                     AutoSize        =   -1  'True
                     BackColor       =   &H80000005&
                     BorderStyle     =   0  'None
                     ForeColor       =   &H80000008&
                     Height          =   945
                     Index           =   3
                     Left            =   240
                     Picture         =   "FProgrammateurCyclique.frx":648F0
                     ScaleHeight     =   63
                     ScaleMode       =   3  'Pixel
                     ScaleWidth      =   661
                     TabIndex        =   47
                     TabStop         =   0   'False
                     Top             =   3720
                     Width           =   9915
                  End
                  Begin VB.PictureBox PBJourneesTypes 
                     BackColor       =   &H00C0C0C0&
                     Height          =   975
                     Index           =   0
                     Left            =   10140
                     ScaleHeight     =   915
                     ScaleWidth      =   4935
                     TabIndex        =   42
                     Top             =   600
                     Width           =   4995
                     Begin VB.OptionButton OBTypesJourneesIdx01 
                        BackColor       =   &H00C0E0FF&
                        Caption         =   "REPRISE"
                        BeginProperty Font 
                           Name            =   "Arial"
                           Size            =   9.75
                           Charset         =   0
                           Weight          =   700
                           Underline       =   0   'False
                           Italic          =   0   'False
                           Strikethrough   =   0   'False
                        EndProperty
                        Height          =   795
                        Index           =   3
                        Left            =   2520
                        MaskColor       =   &H00FFFFFF&
                        Style           =   1  'Graphical
                        TabIndex        =   46
                        Top             =   60
                        Width           =   1095
                     End
                     Begin VB.OptionButton OBTypesJourneesIdx01 
                        BackColor       =   &H00C0E0FF&
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
                        Height          =   795
                        Index           =   0
                        Left            =   120
                        MaskColor       =   &H00FFFFFF&
                        Style           =   1  'Graphical
                        TabIndex        =   45
                        Top             =   60
                        Width           =   1095
                     End
                     Begin VB.OptionButton OBTypesJourneesIdx01 
                        BackColor       =   &H00C0E0FF&
                        Caption         =   "TRAVAIL"
                        BeginProperty Font 
                           Name            =   "Arial"
                           Size            =   9.75
                           Charset         =   0
                           Weight          =   700
                           Underline       =   0   'False
                           Italic          =   0   'False
                           Strikethrough   =   0   'False
                        EndProperty
                        Height          =   795
                        Index           =   1
                        Left            =   3720
                        MaskColor       =   &H00FFFFFF&
                        Style           =   1  'Graphical
                        TabIndex        =   44
                        Top             =   60
                        Width           =   1095
                     End
                     Begin VB.OptionButton OBTypesJourneesIdx01 
                        BackColor       =   &H00C0E0FF&
                        Caption         =   "VEILLE"
                        BeginProperty Font 
                           Name            =   "Arial"
                           Size            =   9.75
                           Charset         =   0
                           Weight          =   700
                           Underline       =   0   'False
                           Italic          =   0   'False
                           Strikethrough   =   0   'False
                        EndProperty
                        Height          =   795
                        Index           =   2
                        Left            =   1320
                        MaskColor       =   &H00FFFFFF&
                        Style           =   1  'Graphical
                        TabIndex        =   43
                        Top             =   60
                        Width           =   1095
                     End
                  End
                  Begin VB.PictureBox PBJourneesTypes 
                     BackColor       =   &H00C0C0C0&
                     Height          =   975
                     Index           =   1
                     Left            =   10140
                     ScaleHeight     =   915
                     ScaleWidth      =   4935
                     TabIndex        =   41
                     Top             =   2160
                     Width           =   4995
                     Begin VB.OptionButton OBTypesJourneesIdx02 
                        BackColor       =   &H00C0E0FF&
                        Caption         =   "REPRISE"
                        BeginProperty Font 
                           Name            =   "Arial"
                           Size            =   9.75
                           Charset         =   0
                           Weight          =   700
                           Underline       =   0   'False
                           Italic          =   0   'False
                           Strikethrough   =   0   'False
                        EndProperty
                        Height          =   795
                        Index           =   3
                        Left            =   2520
                        MaskColor       =   &H00FFFFFF&
                        Style           =   1  'Graphical
                        TabIndex        =   57
                        Top             =   60
                        Width           =   1095
                     End
                     Begin VB.OptionButton OBTypesJourneesIdx02 
                        BackColor       =   &H00C0E0FF&
                        Caption         =   "VEILLE"
                        BeginProperty Font 
                           Name            =   "Arial"
                           Size            =   9.75
                           Charset         =   0
                           Weight          =   700
                           Underline       =   0   'False
                           Italic          =   0   'False
                           Strikethrough   =   0   'False
                        EndProperty
                        Height          =   795
                        Index           =   2
                        Left            =   1320
                        MaskColor       =   &H00FFFFFF&
                        Style           =   1  'Graphical
                        TabIndex        =   56
                        Top             =   60
                        Width           =   1095
                     End
                     Begin VB.OptionButton OBTypesJourneesIdx02 
                        BackColor       =   &H00C0E0FF&
                        Caption         =   "TRAVAIL"
                        BeginProperty Font 
                           Name            =   "Arial"
                           Size            =   9.75
                           Charset         =   0
                           Weight          =   700
                           Underline       =   0   'False
                           Italic          =   0   'False
                           Strikethrough   =   0   'False
                        EndProperty
                        Height          =   795
                        Index           =   1
                        Left            =   3720
                        MaskColor       =   &H00FFFFFF&
                        Style           =   1  'Graphical
                        TabIndex        =   55
                        Top             =   60
                        Width           =   1095
                     End
                     Begin VB.OptionButton OBTypesJourneesIdx02 
                        BackColor       =   &H00C0E0FF&
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
                        Height          =   795
                        Index           =   0
                        Left            =   120
                        MaskColor       =   &H00FFFFFF&
                        Style           =   1  'Graphical
                        TabIndex        =   53
                        Top             =   60
                        Width           =   1095
                     End
                  End
                  Begin VB.PictureBox PBJourneesTypes 
                     BackColor       =   &H00C0C0C0&
                     Height          =   975
                     Index           =   2
                     Left            =   10140
                     ScaleHeight     =   915
                     ScaleWidth      =   4935
                     TabIndex        =   40
                     Top             =   3720
                     Width           =   4995
                     Begin VB.OptionButton OBTypesJourneesIdx03 
                        BackColor       =   &H00C0E0FF&
                        Caption         =   "REPRISE"
                        BeginProperty Font 
                           Name            =   "Arial"
                           Size            =   9.75
                           Charset         =   0
                           Weight          =   700
                           Underline       =   0   'False
                           Italic          =   0   'False
                           Strikethrough   =   0   'False
                        EndProperty
                        Height          =   795
                        Index           =   3
                        Left            =   2520
                        MaskColor       =   &H00FFFFFF&
                        Style           =   1  'Graphical
                        TabIndex        =   60
                        Top             =   60
                        Width           =   1095
                     End
                     Begin VB.OptionButton OBTypesJourneesIdx03 
                        BackColor       =   &H00C0E0FF&
                        Caption         =   "VEILLE"
                        BeginProperty Font 
                           Name            =   "Arial"
                           Size            =   9.75
                           Charset         =   0
                           Weight          =   700
                           Underline       =   0   'False
                           Italic          =   0   'False
                           Strikethrough   =   0   'False
                        EndProperty
                        Height          =   795
                        Index           =   2
                        Left            =   1320
                        MaskColor       =   &H00FFFFFF&
                        Style           =   1  'Graphical
                        TabIndex        =   59
                        Top             =   60
                        Width           =   1095
                     End
                     Begin VB.OptionButton OBTypesJourneesIdx03 
                        BackColor       =   &H00C0E0FF&
                        Caption         =   "TRAVAIL"
                        BeginProperty Font 
                           Name            =   "Arial"
                           Size            =   9.75
                           Charset         =   0
                           Weight          =   700
                           Underline       =   0   'False
                           Italic          =   0   'False
                           Strikethrough   =   0   'False
                        EndProperty
                        Height          =   795
                        Index           =   1
                        Left            =   3720
                        MaskColor       =   &H00FFFFFF&
                        Style           =   1  'Graphical
                        TabIndex        =   58
                        Top             =   60
                        Width           =   1095
                     End
                     Begin VB.OptionButton OBTypesJourneesIdx03 
                        BackColor       =   &H00C0E0FF&
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
                        Height          =   795
                        Index           =   0
                        Left            =   120
                        MaskColor       =   &H00FFFFFF&
                        Style           =   1  'Graphical
                        TabIndex        =   54
                        Top             =   60
                        Width           =   1095
                     End
                  End
                  Begin VB.Label LTitresCuves 
                     Alignment       =   2  'Center
                     BackColor       =   &H00FF0000&
                     BorderStyle     =   1  'Fixed Single
                     Caption         =   "C00"
                     BeginProperty Font 
                        Name            =   "Arial"
                        Size            =   12
                        Charset         =   0
                        Weight          =   700
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     ForeColor       =   &H00FFFFFF&
                     Height          =   375
                     Index           =   1
                     Left            =   240
                     TabIndex        =   52
                     Top             =   240
                     Width           =   14895
                  End
                  Begin VB.Label LTitresCuves 
                     Alignment       =   2  'Center
                     BackColor       =   &H00FF0000&
                     BorderStyle     =   1  'Fixed Single
                     Caption         =   "C02"
                     BeginProperty Font 
                        Name            =   "Arial"
                        Size            =   12
                        Charset         =   0
                        Weight          =   700
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     ForeColor       =   &H00FFFFFF&
                     Height          =   375
                     Index           =   3
                     Left            =   240
                     TabIndex        =   51
                     Top             =   3360
                     Width           =   14895
                  End
                  Begin VB.Label LTitresCuves 
                     Alignment       =   2  'Center
                     BackColor       =   &H00FF0000&
                     BorderStyle     =   1  'Fixed Single
                     Caption         =   "C01"
                     BeginProperty Font 
                        Name            =   "Arial"
                        Size            =   12
                        Charset         =   0
                        Weight          =   700
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     ForeColor       =   &H00FFFFFF&
                     Height          =   375
                     Index           =   2
                     Left            =   240
                     TabIndex        =   50
                     Top             =   1800
                     Width           =   14895
                  End
               End
               Begin VB.PictureBox PBOnglets 
                  Height          =   11475
                  Index           =   3
                  Left            =   16200
                  ScaleHeight     =   11415
                  ScaleWidth      =   15405
                  TabIndex        =   24
                  Top             =   495
                  Width           =   15465
               End
               Begin VB.PictureBox PBOnglets 
                  Height          =   11475
                  Index           =   2
                  Left            =   45
                  ScaleHeight     =   11415
                  ScaleWidth      =   15405
                  TabIndex        =   23
                  Top             =   495
                  Width           =   15465
                  Begin VB.PictureBox PBJourneesTypes 
                     BackColor       =   &H00C0C0C0&
                     Height          =   975
                     Index           =   6
                     Left            =   10140
                     ScaleHeight     =   915
                     ScaleWidth      =   4935
                     TabIndex        =   114
                     Top             =   720
                     Width           =   4995
                     Begin VB.OptionButton OBTypesJourneesIdx07 
                        BackColor       =   &H00C0E0FF&
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
                        Height          =   795
                        Index           =   0
                        Left            =   120
                        MaskColor       =   &H00FFFFFF&
                        Style           =   1  'Graphical
                        TabIndex        =   118
                        Top             =   60
                        Width           =   1095
                     End
                     Begin VB.OptionButton OBTypesJourneesIdx07 
                        BackColor       =   &H00C0E0FF&
                        Caption         =   "TRAVAIL"
                        BeginProperty Font 
                           Name            =   "Arial"
                           Size            =   9.75
                           Charset         =   0
                           Weight          =   700
                           Underline       =   0   'False
                           Italic          =   0   'False
                           Strikethrough   =   0   'False
                        EndProperty
                        Height          =   795
                        Index           =   1
                        Left            =   3720
                        MaskColor       =   &H00FFFFFF&
                        Style           =   1  'Graphical
                        TabIndex        =   117
                        Top             =   60
                        Width           =   1095
                     End
                     Begin VB.OptionButton OBTypesJourneesIdx07 
                        BackColor       =   &H00C0E0FF&
                        Caption         =   "VEILLE"
                        BeginProperty Font 
                           Name            =   "Arial"
                           Size            =   9.75
                           Charset         =   0
                           Weight          =   700
                           Underline       =   0   'False
                           Italic          =   0   'False
                           Strikethrough   =   0   'False
                        EndProperty
                        Height          =   795
                        Index           =   2
                        Left            =   1320
                        MaskColor       =   &H00FFFFFF&
                        Style           =   1  'Graphical
                        TabIndex        =   116
                        Top             =   60
                        Width           =   1095
                     End
                     Begin VB.OptionButton OBTypesJourneesIdx07 
                        BackColor       =   &H00C0E0FF&
                        Caption         =   "REPRISE"
                        BeginProperty Font 
                           Name            =   "Arial"
                           Size            =   9.75
                           Charset         =   0
                           Weight          =   700
                           Underline       =   0   'False
                           Italic          =   0   'False
                           Strikethrough   =   0   'False
                        EndProperty
                        Height          =   795
                        Index           =   3
                        Left            =   2520
                        MaskColor       =   &H00FFFFFF&
                        Style           =   1  'Graphical
                        TabIndex        =   115
                        Top             =   60
                        Width           =   1095
                     End
                  End
                  Begin VB.PictureBox PBEchelle24H 
                     Appearance      =   0  'Flat
                     AutoRedraw      =   -1  'True
                     AutoSize        =   -1  'True
                     BackColor       =   &H80000005&
                     BorderStyle     =   0  'None
                     ForeColor       =   &H80000008&
                     Height          =   945
                     Index           =   7
                     Left            =   240
                     Picture         =   "FProgrammateurCyclique.frx":83172
                     ScaleHeight     =   63
                     ScaleMode       =   3  'Pixel
                     ScaleWidth      =   661
                     TabIndex        =   113
                     TabStop         =   0   'False
                     Top             =   720
                     Width           =   9915
                  End
                  Begin VB.PictureBox PBJourneesTypes 
                     BackColor       =   &H00C0C0C0&
                     Height          =   975
                     Index           =   7
                     Left            =   10140
                     ScaleHeight     =   915
                     ScaleWidth      =   4935
                     TabIndex        =   104
                     Top             =   2280
                     Width           =   4995
                     Begin VB.OptionButton OBTypesJourneesIdx08 
                        BackColor       =   &H00C0E0FF&
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
                        Height          =   795
                        Index           =   0
                        Left            =   120
                        MaskColor       =   &H00FFFFFF&
                        Style           =   1  'Graphical
                        TabIndex        =   108
                        Top             =   60
                        Width           =   1095
                     End
                     Begin VB.OptionButton OBTypesJourneesIdx08 
                        BackColor       =   &H00C0E0FF&
                        Caption         =   "TRAVAIL"
                        BeginProperty Font 
                           Name            =   "Arial"
                           Size            =   9.75
                           Charset         =   0
                           Weight          =   700
                           Underline       =   0   'False
                           Italic          =   0   'False
                           Strikethrough   =   0   'False
                        EndProperty
                        Height          =   795
                        Index           =   1
                        Left            =   3720
                        MaskColor       =   &H00FFFFFF&
                        Style           =   1  'Graphical
                        TabIndex        =   107
                        Top             =   60
                        Width           =   1095
                     End
                     Begin VB.OptionButton OBTypesJourneesIdx08 
                        BackColor       =   &H00C0E0FF&
                        Caption         =   "VEILLE"
                        BeginProperty Font 
                           Name            =   "Arial"
                           Size            =   9.75
                           Charset         =   0
                           Weight          =   700
                           Underline       =   0   'False
                           Italic          =   0   'False
                           Strikethrough   =   0   'False
                        EndProperty
                        Height          =   795
                        Index           =   2
                        Left            =   1320
                        MaskColor       =   &H00FFFFFF&
                        Style           =   1  'Graphical
                        TabIndex        =   106
                        Top             =   60
                        Width           =   1095
                     End
                     Begin VB.OptionButton OBTypesJourneesIdx08 
                        BackColor       =   &H00C0E0FF&
                        Caption         =   "REPRISE"
                        BeginProperty Font 
                           Name            =   "Arial"
                           Size            =   9.75
                           Charset         =   0
                           Weight          =   700
                           Underline       =   0   'False
                           Italic          =   0   'False
                           Strikethrough   =   0   'False
                        EndProperty
                        Height          =   795
                        Index           =   3
                        Left            =   2520
                        MaskColor       =   &H00FFFFFF&
                        Style           =   1  'Graphical
                        TabIndex        =   105
                        Top             =   60
                        Width           =   1095
                     End
                  End
                  Begin VB.PictureBox PBJourneesTypes 
                     BackColor       =   &H00C0C0C0&
                     Height          =   975
                     Index           =   8
                     Left            =   10140
                     ScaleHeight     =   915
                     ScaleWidth      =   4935
                     TabIndex        =   99
                     Top             =   3840
                     Width           =   4995
                     Begin VB.OptionButton OBTypesJourneesIdx09 
                        BackColor       =   &H00C0E0FF&
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
                        Height          =   795
                        Index           =   0
                        Left            =   120
                        MaskColor       =   &H00FFFFFF&
                        Style           =   1  'Graphical
                        TabIndex        =   103
                        Top             =   60
                        Width           =   1095
                     End
                     Begin VB.OptionButton OBTypesJourneesIdx09 
                        BackColor       =   &H00C0E0FF&
                        Caption         =   "TRAVAIL"
                        BeginProperty Font 
                           Name            =   "Arial"
                           Size            =   9.75
                           Charset         =   0
                           Weight          =   700
                           Underline       =   0   'False
                           Italic          =   0   'False
                           Strikethrough   =   0   'False
                        EndProperty
                        Height          =   795
                        Index           =   1
                        Left            =   3720
                        MaskColor       =   &H00FFFFFF&
                        Style           =   1  'Graphical
                        TabIndex        =   102
                        Top             =   60
                        Width           =   1095
                     End
                     Begin VB.OptionButton OBTypesJourneesIdx09 
                        BackColor       =   &H00C0E0FF&
                        Caption         =   "VEILLE"
                        BeginProperty Font 
                           Name            =   "Arial"
                           Size            =   9.75
                           Charset         =   0
                           Weight          =   700
                           Underline       =   0   'False
                           Italic          =   0   'False
                           Strikethrough   =   0   'False
                        EndProperty
                        Height          =   795
                        Index           =   2
                        Left            =   1320
                        MaskColor       =   &H00FFFFFF&
                        Style           =   1  'Graphical
                        TabIndex        =   101
                        Top             =   60
                        Width           =   1095
                     End
                     Begin VB.OptionButton OBTypesJourneesIdx09 
                        BackColor       =   &H00C0E0FF&
                        Caption         =   "REPRISE"
                        BeginProperty Font 
                           Name            =   "Arial"
                           Size            =   9.75
                           Charset         =   0
                           Weight          =   700
                           Underline       =   0   'False
                           Italic          =   0   'False
                           Strikethrough   =   0   'False
                        EndProperty
                        Height          =   795
                        Index           =   3
                        Left            =   2520
                        MaskColor       =   &H00FFFFFF&
                        Style           =   1  'Graphical
                        TabIndex        =   100
                        Top             =   60
                        Width           =   1095
                     End
                  End
                  Begin VB.PictureBox PBJourneesTypes 
                     BackColor       =   &H00C0C0C0&
                     Height          =   975
                     Index           =   9
                     Left            =   10140
                     ScaleHeight     =   915
                     ScaleWidth      =   4935
                     TabIndex        =   94
                     Top             =   5400
                     Width           =   4995
                     Begin VB.OptionButton OBTypesJourneesIdx10 
                        BackColor       =   &H00C0E0FF&
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
                        Height          =   795
                        Index           =   0
                        Left            =   120
                        MaskColor       =   &H00FFFFFF&
                        Style           =   1  'Graphical
                        TabIndex        =   98
                        Top             =   60
                        Width           =   1095
                     End
                     Begin VB.OptionButton OBTypesJourneesIdx10 
                        BackColor       =   &H00C0E0FF&
                        Caption         =   "TRAVAIL"
                        BeginProperty Font 
                           Name            =   "Arial"
                           Size            =   9.75
                           Charset         =   0
                           Weight          =   700
                           Underline       =   0   'False
                           Italic          =   0   'False
                           Strikethrough   =   0   'False
                        EndProperty
                        Height          =   795
                        Index           =   1
                        Left            =   3720
                        MaskColor       =   &H00FFFFFF&
                        Style           =   1  'Graphical
                        TabIndex        =   97
                        Top             =   60
                        Width           =   1095
                     End
                     Begin VB.OptionButton OBTypesJourneesIdx10 
                        BackColor       =   &H00C0E0FF&
                        Caption         =   "VEILLE"
                        BeginProperty Font 
                           Name            =   "Arial"
                           Size            =   9.75
                           Charset         =   0
                           Weight          =   700
                           Underline       =   0   'False
                           Italic          =   0   'False
                           Strikethrough   =   0   'False
                        EndProperty
                        Height          =   795
                        Index           =   2
                        Left            =   1320
                        MaskColor       =   &H00FFFFFF&
                        Style           =   1  'Graphical
                        TabIndex        =   96
                        Top             =   60
                        Width           =   1095
                     End
                     Begin VB.OptionButton OBTypesJourneesIdx10 
                        BackColor       =   &H00C0E0FF&
                        Caption         =   "REPRISE"
                        BeginProperty Font 
                           Name            =   "Arial"
                           Size            =   9.75
                           Charset         =   0
                           Weight          =   700
                           Underline       =   0   'False
                           Italic          =   0   'False
                           Strikethrough   =   0   'False
                        EndProperty
                        Height          =   795
                        Index           =   3
                        Left            =   2520
                        MaskColor       =   &H00FFFFFF&
                        Style           =   1  'Graphical
                        TabIndex        =   95
                        Top             =   60
                        Width           =   1095
                     End
                  End
                  Begin VB.PictureBox PBJourneesTypes 
                     BackColor       =   &H00C0C0C0&
                     Height          =   975
                     Index           =   10
                     Left            =   10140
                     ScaleHeight     =   915
                     ScaleWidth      =   4935
                     TabIndex        =   89
                     Top             =   6960
                     Width           =   4995
                     Begin VB.OptionButton OBTypesJourneesIdx11 
                        BackColor       =   &H00C0E0FF&
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
                        Height          =   795
                        Index           =   0
                        Left            =   120
                        MaskColor       =   &H00FFFFFF&
                        Style           =   1  'Graphical
                        TabIndex        =   93
                        Top             =   60
                        Width           =   1095
                     End
                     Begin VB.OptionButton OBTypesJourneesIdx11 
                        BackColor       =   &H00C0E0FF&
                        Caption         =   "TRAVAIL"
                        BeginProperty Font 
                           Name            =   "Arial"
                           Size            =   9.75
                           Charset         =   0
                           Weight          =   700
                           Underline       =   0   'False
                           Italic          =   0   'False
                           Strikethrough   =   0   'False
                        EndProperty
                        Height          =   795
                        Index           =   1
                        Left            =   3720
                        MaskColor       =   &H00FFFFFF&
                        Style           =   1  'Graphical
                        TabIndex        =   92
                        Top             =   60
                        Width           =   1095
                     End
                     Begin VB.OptionButton OBTypesJourneesIdx11 
                        BackColor       =   &H00C0E0FF&
                        Caption         =   "VEILLE"
                        BeginProperty Font 
                           Name            =   "Arial"
                           Size            =   9.75
                           Charset         =   0
                           Weight          =   700
                           Underline       =   0   'False
                           Italic          =   0   'False
                           Strikethrough   =   0   'False
                        EndProperty
                        Height          =   795
                        Index           =   2
                        Left            =   1320
                        MaskColor       =   &H00FFFFFF&
                        Style           =   1  'Graphical
                        TabIndex        =   91
                        Top             =   60
                        Width           =   1095
                     End
                     Begin VB.OptionButton OBTypesJourneesIdx11 
                        BackColor       =   &H00C0E0FF&
                        Caption         =   "REPRISE"
                        BeginProperty Font 
                           Name            =   "Arial"
                           Size            =   9.75
                           Charset         =   0
                           Weight          =   700
                           Underline       =   0   'False
                           Italic          =   0   'False
                           Strikethrough   =   0   'False
                        EndProperty
                        Height          =   795
                        Index           =   3
                        Left            =   2520
                        MaskColor       =   &H00FFFFFF&
                        Style           =   1  'Graphical
                        TabIndex        =   90
                        Top             =   60
                        Width           =   1095
                     End
                  End
                  Begin VB.PictureBox PBEchelle24H 
                     Appearance      =   0  'Flat
                     AutoRedraw      =   -1  'True
                     AutoSize        =   -1  'True
                     BackColor       =   &H80000005&
                     BorderStyle     =   0  'None
                     ForeColor       =   &H80000008&
                     Height          =   945
                     Index           =   9
                     Left            =   240
                     Picture         =   "FProgrammateurCyclique.frx":A19F4
                     ScaleHeight     =   63
                     ScaleMode       =   3  'Pixel
                     ScaleWidth      =   661
                     TabIndex        =   88
                     TabStop         =   0   'False
                     Top             =   3840
                     Width           =   9915
                  End
                  Begin VB.PictureBox PBEchelle24H 
                     Appearance      =   0  'Flat
                     AutoRedraw      =   -1  'True
                     AutoSize        =   -1  'True
                     BackColor       =   &H80000005&
                     BorderStyle     =   0  'None
                     ForeColor       =   &H80000008&
                     Height          =   945
                     Index           =   10
                     Left            =   240
                     Picture         =   "FProgrammateurCyclique.frx":C0276
                     ScaleHeight     =   63
                     ScaleMode       =   3  'Pixel
                     ScaleWidth      =   661
                     TabIndex        =   87
                     TabStop         =   0   'False
                     Top             =   5400
                     Width           =   9915
                  End
                  Begin VB.PictureBox PBEchelle24H 
                     Appearance      =   0  'Flat
                     AutoRedraw      =   -1  'True
                     AutoSize        =   -1  'True
                     BackColor       =   &H80000005&
                     BorderStyle     =   0  'None
                     ForeColor       =   &H80000008&
                     Height          =   945
                     Index           =   11
                     Left            =   240
                     Picture         =   "FProgrammateurCyclique.frx":DEAF8
                     ScaleHeight     =   63
                     ScaleMode       =   3  'Pixel
                     ScaleWidth      =   661
                     TabIndex        =   86
                     TabStop         =   0   'False
                     Top             =   6960
                     Width           =   9915
                  End
                  Begin VB.PictureBox PBEchelle24H 
                     Appearance      =   0  'Flat
                     AutoRedraw      =   -1  'True
                     AutoSize        =   -1  'True
                     BackColor       =   &H80000005&
                     BorderStyle     =   0  'None
                     ForeColor       =   &H80000008&
                     Height          =   945
                     Index           =   8
                     Left            =   240
                     Picture         =   "FProgrammateurCyclique.frx":FD37A
                     ScaleHeight     =   63
                     ScaleMode       =   3  'Pixel
                     ScaleWidth      =   661
                     TabIndex        =   85
                     TabStop         =   0   'False
                     Top             =   2280
                     Width           =   9915
                  End
                  Begin VB.Label LTitresCuves 
                     Alignment       =   2  'Center
                     BackColor       =   &H00FF0000&
                     BorderStyle     =   1  'Fixed Single
                     Caption         =   "C07"
                     BeginProperty Font 
                        Name            =   "Arial"
                        Size            =   12
                        Charset         =   0
                        Weight          =   700
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     ForeColor       =   &H00FFFFFF&
                     Height          =   375
                     Index           =   7
                     Left            =   240
                     TabIndex        =   119
                     Top             =   360
                     Width           =   14895
                  End
                  Begin VB.Label LTitresCuves 
                     Alignment       =   2  'Center
                     BackColor       =   &H00FF0000&
                     BorderStyle     =   1  'Fixed Single
                     Caption         =   "C16"
                     BeginProperty Font 
                        Name            =   "Arial"
                        Size            =   12
                        Charset         =   0
                        Weight          =   700
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     ForeColor       =   &H00FFFFFF&
                     Height          =   375
                     Index           =   11
                     Left            =   240
                     TabIndex        =   112
                     Top             =   6600
                     Width           =   14895
                  End
                  Begin VB.Label LTitresCuves 
                     Alignment       =   2  'Center
                     BackColor       =   &H00FF0000&
                     BorderStyle     =   1  'Fixed Single
                     Caption         =   "C15"
                     BeginProperty Font 
                        Name            =   "Arial"
                        Size            =   12
                        Charset         =   0
                        Weight          =   700
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     ForeColor       =   &H00FFFFFF&
                     Height          =   375
                     Index           =   10
                     Left            =   240
                     TabIndex        =   111
                     Top             =   5040
                     Width           =   14895
                  End
                  Begin VB.Label LTitresCuves 
                     Alignment       =   2  'Center
                     BackColor       =   &H00FF0000&
                     BorderStyle     =   1  'Fixed Single
                     Caption         =   "C14"
                     BeginProperty Font 
                        Name            =   "Arial"
                        Size            =   12
                        Charset         =   0
                        Weight          =   700
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     ForeColor       =   &H00FFFFFF&
                     Height          =   375
                     Index           =   9
                     Left            =   240
                     TabIndex        =   110
                     Top             =   3480
                     Width           =   14895
                  End
                  Begin VB.Label LTitresCuves 
                     Alignment       =   2  'Center
                     BackColor       =   &H00FF0000&
                     BorderStyle     =   1  'Fixed Single
                     Caption         =   "C13"
                     BeginProperty Font 
                        Name            =   "Arial"
                        Size            =   12
                        Charset         =   0
                        Weight          =   700
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     ForeColor       =   &H00FFFFFF&
                     Height          =   375
                     Index           =   8
                     Left            =   240
                     TabIndex        =   109
                     Top             =   1920
                     Width           =   14895
                  End
               End
               Begin VB.PictureBox PBOnglets 
                  Height          =   11475
                  Index           =   8
                  Left            =   18000
                  ScaleHeight     =   11415
                  ScaleWidth      =   15405
                  TabIndex        =   22
                  Top             =   495
                  Width           =   15465
               End
               Begin VB.PictureBox PBOnglets 
                  Height          =   11475
                  Index           =   7
                  Left            =   17700
                  ScaleHeight     =   11415
                  ScaleWidth      =   15405
                  TabIndex        =   21
                  Top             =   495
                  Width           =   15465
               End
               Begin VB.PictureBox PBOnglets 
                  Height          =   11475
                  Index           =   6
                  Left            =   17400
                  ScaleHeight     =   11415
                  ScaleWidth      =   15405
                  TabIndex        =   20
                  Top             =   495
                  Width           =   15465
               End
               Begin VB.PictureBox PBOnglets 
                  Height          =   11475
                  Index           =   5
                  Left            =   17100
                  ScaleHeight     =   11415
                  ScaleWidth      =   15405
                  TabIndex        =   19
                  Top             =   495
                  Width           =   15465
               End
               Begin VB.PictureBox PBOnglets 
                  Height          =   11475
                  Index           =   4
                  Left            =   16500
                  ScaleHeight     =   11415
                  ScaleWidth      =   15405
                  TabIndex        =   18
                  Top             =   495
                  Width           =   15465
               End
               Begin VB.PictureBox PBOnglets 
                  Height          =   11475
                  Index           =   1
                  Left            =   -16110
                  ScaleHeight     =   11415
                  ScaleWidth      =   15405
                  TabIndex        =   17
                  Top             =   495
                  Width           =   15465
                  Begin VB.PictureBox PBJourneesTypes 
                     BackColor       =   &H00C0C0C0&
                     Height          =   975
                     Index           =   3
                     Left            =   10140
                     ScaleHeight     =   915
                     ScaleWidth      =   4935
                     TabIndex        =   133
                     Top             =   840
                     Width           =   4995
                     Begin VB.OptionButton OBTypesJourneesIdx04 
                        BackColor       =   &H00C0E0FF&
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
                        Height          =   795
                        Index           =   0
                        Left            =   120
                        MaskColor       =   &H00FFFFFF&
                        Style           =   1  'Graphical
                        TabIndex        =   137
                        Top             =   60
                        Width           =   1095
                     End
                     Begin VB.OptionButton OBTypesJourneesIdx04 
                        BackColor       =   &H00C0E0FF&
                        Caption         =   "TRAVAIL"
                        BeginProperty Font 
                           Name            =   "Arial"
                           Size            =   9.75
                           Charset         =   0
                           Weight          =   700
                           Underline       =   0   'False
                           Italic          =   0   'False
                           Strikethrough   =   0   'False
                        EndProperty
                        Height          =   795
                        Index           =   1
                        Left            =   3720
                        MaskColor       =   &H00FFFFFF&
                        Style           =   1  'Graphical
                        TabIndex        =   136
                        Top             =   60
                        Width           =   1095
                     End
                     Begin VB.OptionButton OBTypesJourneesIdx04 
                        BackColor       =   &H00C0E0FF&
                        Caption         =   "VEILLE"
                        BeginProperty Font 
                           Name            =   "Arial"
                           Size            =   9.75
                           Charset         =   0
                           Weight          =   700
                           Underline       =   0   'False
                           Italic          =   0   'False
                           Strikethrough   =   0   'False
                        EndProperty
                        Height          =   795
                        Index           =   2
                        Left            =   1320
                        MaskColor       =   &H00FFFFFF&
                        Style           =   1  'Graphical
                        TabIndex        =   135
                        Top             =   60
                        Width           =   1095
                     End
                     Begin VB.OptionButton OBTypesJourneesIdx04 
                        BackColor       =   &H00C0E0FF&
                        Caption         =   "REPRISE"
                        BeginProperty Font 
                           Name            =   "Arial"
                           Size            =   9.75
                           Charset         =   0
                           Weight          =   700
                           Underline       =   0   'False
                           Italic          =   0   'False
                           Strikethrough   =   0   'False
                        EndProperty
                        Height          =   795
                        Index           =   3
                        Left            =   2520
                        MaskColor       =   &H00FFFFFF&
                        Style           =   1  'Graphical
                        TabIndex        =   134
                        Top             =   60
                        Width           =   1095
                     End
                  End
                  Begin VB.PictureBox PBEchelle24H 
                     Appearance      =   0  'Flat
                     AutoRedraw      =   -1  'True
                     AutoSize        =   -1  'True
                     BackColor       =   &H80000005&
                     BorderStyle     =   0  'None
                     ForeColor       =   &H80000008&
                     Height          =   945
                     Index           =   4
                     Left            =   240
                     Picture         =   "FProgrammateurCyclique.frx":11BBFC
                     ScaleHeight     =   63
                     ScaleMode       =   3  'Pixel
                     ScaleWidth      =   661
                     TabIndex        =   132
                     TabStop         =   0   'False
                     Top             =   840
                     Width           =   9915
                  End
                  Begin VB.PictureBox PBJourneesTypes 
                     BackColor       =   &H00C0C0C0&
                     Height          =   975
                     Index           =   5
                     Left            =   10140
                     ScaleHeight     =   915
                     ScaleWidth      =   4935
                     TabIndex        =   127
                     Top             =   3960
                     Width           =   4995
                     Begin VB.OptionButton OBTypesJourneesIdx06 
                        BackColor       =   &H00C0E0FF&
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
                        Height          =   795
                        Index           =   0
                        Left            =   120
                        MaskColor       =   &H00FFFFFF&
                        Style           =   1  'Graphical
                        TabIndex        =   131
                        Top             =   60
                        Width           =   1095
                     End
                     Begin VB.OptionButton OBTypesJourneesIdx06 
                        BackColor       =   &H00C0E0FF&
                        Caption         =   "TRAVAIL"
                        BeginProperty Font 
                           Name            =   "Arial"
                           Size            =   9.75
                           Charset         =   0
                           Weight          =   700
                           Underline       =   0   'False
                           Italic          =   0   'False
                           Strikethrough   =   0   'False
                        EndProperty
                        Height          =   795
                        Index           =   1
                        Left            =   3720
                        MaskColor       =   &H00FFFFFF&
                        Style           =   1  'Graphical
                        TabIndex        =   130
                        Top             =   60
                        Width           =   1095
                     End
                     Begin VB.OptionButton OBTypesJourneesIdx06 
                        BackColor       =   &H00C0E0FF&
                        Caption         =   "VEILLE"
                        BeginProperty Font 
                           Name            =   "Arial"
                           Size            =   9.75
                           Charset         =   0
                           Weight          =   700
                           Underline       =   0   'False
                           Italic          =   0   'False
                           Strikethrough   =   0   'False
                        EndProperty
                        Height          =   795
                        Index           =   2
                        Left            =   1320
                        MaskColor       =   &H00FFFFFF&
                        Style           =   1  'Graphical
                        TabIndex        =   129
                        Top             =   60
                        Width           =   1095
                     End
                     Begin VB.OptionButton OBTypesJourneesIdx06 
                        BackColor       =   &H00C0E0FF&
                        Caption         =   "REPRISE"
                        BeginProperty Font 
                           Name            =   "Arial"
                           Size            =   9.75
                           Charset         =   0
                           Weight          =   700
                           Underline       =   0   'False
                           Italic          =   0   'False
                           Strikethrough   =   0   'False
                        EndProperty
                        Height          =   795
                        Index           =   3
                        Left            =   2520
                        MaskColor       =   &H00FFFFFF&
                        Style           =   1  'Graphical
                        TabIndex        =   128
                        Top             =   60
                        Width           =   1095
                     End
                  End
                  Begin VB.PictureBox PBEchelle24H 
                     Appearance      =   0  'Flat
                     AutoRedraw      =   -1  'True
                     AutoSize        =   -1  'True
                     BackColor       =   &H80000005&
                     BorderStyle     =   0  'None
                     ForeColor       =   &H80000008&
                     Height          =   945
                     Index           =   6
                     Left            =   240
                     Picture         =   "FProgrammateurCyclique.frx":13A47E
                     ScaleHeight     =   63
                     ScaleMode       =   3  'Pixel
                     ScaleWidth      =   661
                     TabIndex        =   126
                     TabStop         =   0   'False
                     Top             =   3960
                     Width           =   9915
                  End
                  Begin VB.PictureBox PBJourneesTypes 
                     BackColor       =   &H00C0C0C0&
                     Height          =   975
                     Index           =   4
                     Left            =   10140
                     ScaleHeight     =   915
                     ScaleWidth      =   4935
                     TabIndex        =   121
                     Top             =   2400
                     Width           =   4995
                     Begin VB.OptionButton OBTypesJourneesIdx05 
                        BackColor       =   &H00C0E0FF&
                        Caption         =   "TRAVAIL"
                        BeginProperty Font 
                           Name            =   "Arial"
                           Size            =   9.75
                           Charset         =   0
                           Weight          =   700
                           Underline       =   0   'False
                           Italic          =   0   'False
                           Strikethrough   =   0   'False
                        EndProperty
                        Height          =   795
                        Index           =   1
                        Left            =   3720
                        MaskColor       =   &H00FFFFFF&
                        Style           =   1  'Graphical
                        TabIndex        =   125
                        Top             =   60
                        Width           =   1095
                     End
                     Begin VB.OptionButton OBTypesJourneesIdx05 
                        BackColor       =   &H00C0E0FF&
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
                        Height          =   795
                        Index           =   0
                        Left            =   120
                        MaskColor       =   &H00FFFFFF&
                        Style           =   1  'Graphical
                        TabIndex        =   124
                        Top             =   60
                        Width           =   1095
                     End
                     Begin VB.OptionButton OBTypesJourneesIdx05 
                        BackColor       =   &H00C0E0FF&
                        Caption         =   "VEILLE"
                        BeginProperty Font 
                           Name            =   "Arial"
                           Size            =   9.75
                           Charset         =   0
                           Weight          =   700
                           Underline       =   0   'False
                           Italic          =   0   'False
                           Strikethrough   =   0   'False
                        EndProperty
                        Height          =   795
                        Index           =   2
                        Left            =   1320
                        MaskColor       =   &H00FFFFFF&
                        Style           =   1  'Graphical
                        TabIndex        =   123
                        Top             =   60
                        Width           =   1095
                     End
                     Begin VB.OptionButton OBTypesJourneesIdx05 
                        BackColor       =   &H00C0E0FF&
                        Caption         =   "REPRISE"
                        BeginProperty Font 
                           Name            =   "Arial"
                           Size            =   9.75
                           Charset         =   0
                           Weight          =   700
                           Underline       =   0   'False
                           Italic          =   0   'False
                           Strikethrough   =   0   'False
                        EndProperty
                        Height          =   795
                        Index           =   3
                        Left            =   2520
                        MaskColor       =   &H00FFFFFF&
                        Style           =   1  'Graphical
                        TabIndex        =   122
                        Top             =   60
                        Width           =   1095
                     End
                  End
                  Begin VB.PictureBox PBEchelle24H 
                     Appearance      =   0  'Flat
                     AutoRedraw      =   -1  'True
                     AutoSize        =   -1  'True
                     BackColor       =   &H80000005&
                     BorderStyle     =   0  'None
                     ForeColor       =   &H80000008&
                     Height          =   945
                     Index           =   5
                     Left            =   240
                     Picture         =   "FProgrammateurCyclique.frx":158D00
                     ScaleHeight     =   63
                     ScaleMode       =   3  'Pixel
                     ScaleWidth      =   661
                     TabIndex        =   120
                     TabStop         =   0   'False
                     Top             =   2400
                     Width           =   9915
                  End
                  Begin VB.Label LTitresCuves 
                     Alignment       =   2  'Center
                     BackColor       =   &H00FF0000&
                     BorderStyle     =   1  'Fixed Single
                     Caption         =   "C03"
                     BeginProperty Font 
                        Name            =   "Arial"
                        Size            =   12
                        Charset         =   0
                        Weight          =   700
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     ForeColor       =   &H00FFFFFF&
                     Height          =   375
                     Index           =   4
                     Left            =   240
                     TabIndex        =   140
                     Top             =   480
                     Width           =   14895
                  End
                  Begin VB.Label LTitresCuves 
                     Alignment       =   2  'Center
                     BackColor       =   &H00FF0000&
                     BorderStyle     =   1  'Fixed Single
                     Caption         =   "C05"
                     BeginProperty Font 
                        Name            =   "Arial"
                        Size            =   12
                        Charset         =   0
                        Weight          =   700
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     ForeColor       =   &H00FFFFFF&
                     Height          =   375
                     Index           =   5
                     Left            =   240
                     TabIndex        =   139
                     Top             =   2040
                     Width           =   14895
                  End
                  Begin VB.Label LTitresCuves 
                     Alignment       =   2  'Center
                     BackColor       =   &H00FF0000&
                     BorderStyle     =   1  'Fixed Single
                     Caption         =   "C06"
                     BeginProperty Font 
                        Name            =   "Arial"
                        Size            =   12
                        Charset         =   0
                        Weight          =   700
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     ForeColor       =   &H00FFFFFF&
                     Height          =   375
                     Index           =   6
                     Left            =   240
                     TabIndex        =   138
                     Top             =   3600
                     Width           =   14895
                  End
               End
               Begin VB.PictureBox PBOnglets 
                  Height          =   11475
                  Index           =   9
                  Left            =   16800
                  ScaleHeight     =   11415
                  ScaleWidth      =   15405
                  TabIndex        =   16
                  Top             =   495
                  Width           =   15465
               End
            End
            Begin VB.Label LLibelles 
               Alignment       =   2  'Center
               BackColor       =   &H00800000&
               BorderStyle     =   1  'Fixed Single
               Caption         =   "MODE GENERAL POUR TOUTES LES CUVES"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   12
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FFFFFF&
               Height          =   375
               Index           =   0
               Left            =   300
               TabIndex        =   79
               Top             =   9420
               Width           =   5295
            End
            Begin VB.Label LLibelles 
               Alignment       =   2  'Center
               BackColor       =   &H00800000&
               BorderStyle     =   1  'Fixed Single
               Caption         =   "CHOIX DE LA JOURNEE"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   12
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FFFFFF&
               Height          =   375
               Index           =   1
               Left            =   300
               TabIndex        =   61
               Top             =   360
               Width           =   5295
            End
         End
      End
   End
End
Attribute VB_Name = "FProgrammateurCyclique"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Rôle                    : Fenêtre gérant le programmateur cyclique
' Nom                    : FProgrammateurCyclique.frm
' Date de création : 06/03/2001
' Détails                :
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

'--- déclarations obligatoires ---
Option Explicit

'--- options générales ---
Option Base 1
DefVar A-Z
    
'--- constantes privées ---
Private Const LARGEUR_ECHELLE As Integer = 19                               'largeur de base de l'échelle BMP
Private Const COULEUR_SELECTION As Long = COULEURS.CYAN_6  'couleur avant traitement graphique

Private Const X_MINI_POMPE As Single = 2
Private Const X_MAXI_POMPE As Single = 626                                     'encadrement de la zone pompe
Private Const Y_MINI_POMPE As Single = 2                                         '(X mini, Y mini, X maxi, Y Maxi) de la zone
Private Const Y_MAXI_POMPE As Single = 16

Private Const X_MINI_CHAUFFAGE As Single = 2
Private Const X_MAXI_CHAUFFAGE As Single = 626                             'encadrement de la zone chauffage
Private Const Y_MINI_CHAUFFAGE As Single = 46                                '(X mini, Y mini, X maxi, Y Maxi) de la zone
Private Const Y_MAXI_CHAUFFAGE As Single = 60

Private Const LARGEUR_SEGMENT As Single = 13                              'largeur d'un segment correspondant à 1/2 heure

Private Const TITRE_FENETRE As String = "PROGRAMMATEUR CYCLIQUE"
Private Const TITRE_MESSAGES As String = INDICATIF_PROGRAMME & TITRE_FENETRE

'--- énumérations privées ---

'--- onglets ---
Private Enum ONGLETS_PROGRAMMATEUR_CYCLIQUE
    O_PREPARATION = 0
    O_ANODISATION = 1
    O_COLORATION_FIN_DE_LIGNE = 2
End Enum

'--- variables privées ---
Private PremiereActivation As Boolean
Private InterdireEvenements As Boolean              'pour interdire certains évènements
Private CouleurPompe As Long                              'couleur du segment pompe
Private CouleurChauffage As Long                         'couleur du segment chauffage
Private XDepart As Single                                       'point de départ d'une programmation
Private XArrivee As Single                                      'point d'arrivée d'une programmation
Private MemX As Single                                          'mémoire des X
Private MemY As Single                                          'mémoire des Y
Private ModificationEnCours As Boolean                'TRUE=modification d'un programme en cours
Private ProgrammationEnCours As Boolean           'TRUE=programmation en cours
Private LaDateATraiter As String * 8                        'date à traiter du programmateur cyclique

'--- tableaux privées ---
Private TempsSortieObligatoire As Integer              'temps avant sortie obligatoire du programmateur cyclique

Private TCopieProgCyclique(1 To NBR_JOURS_PROG_CYCLIQUE, CUVES_REGULATION.C_C00 To DERNIERE_CUV_REGULATION) As VarCycle24HProgCyclique
Private TTypesJourneesEnCours(CUVES_REGULATION.C_C00 To DERNIERE_CUV_REGULATION) As JOURNEES_TYPES


'--- variables publiques ---
Public NumFenetre As Long                             'numéro de la fenêtre lorsqu'elle devient active

Private Sub CBAgrandirFENETRE_Click()
    On Error Resume Next
    Me.WindowState = vbMaximized
End Sub

Private Sub CBAnnuler_Click()
    
    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next

    '--- restitue les dernières manipulations et valeurs sur la fenêtre ---
    LectureValeursFenetre
    
    '--- RAZ de la variable de comptage avant sortie obligatoire ---
    TempsSortieObligatoire = 0

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

Private Sub CBModeGeneralCuves_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next

    '--- changement de la couleur de fond ---
    CBModeGeneralCuves(Index).BackColor = COULEURS.ROUGE_1

    '--- changement du mode de la totalité des cuves ---
    OBTypesJourneesIdx01(Index).value = True
    OBTypesJourneesIdx02(Index).value = True
    OBTypesJourneesIdx03(Index).value = True
    OBTypesJourneesIdx04(Index).value = True
    OBTypesJourneesIdx05(Index).value = True
    OBTypesJourneesIdx06(Index).value = True
    OBTypesJourneesIdx07(Index).value = True
    OBTypesJourneesIdx08(Index).value = True
    OBTypesJourneesIdx09(Index).value = True
    OBTypesJourneesIdx10(Index).value = True
    OBTypesJourneesIdx11(Index).value = True
    'OBTypesJourneesIdx12(Index).Value = True
    'OBTypesJourneesIdx13(Index).Value = True
    'OBTypesJourneesIdx14(Index).Value = True
    'OBTypesJourneesIdx15(Index).Value = True
    'OBTypesJourneesIdx16(Index).Value = True
    'OBTypesJourneesIdx17(Index).Value = True
    'OBTypesJourneesIdx18(Index).Value = True
    'OBTypesJourneesIdx22(Index).Value = True

End Sub

Private Sub CBModeGeneralCuves_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next

    '--- changement de la couleur de fond ---
    CBModeGeneralCuves(Index).BackColor = COULEURS.VERT_0

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
    
    '--- enregistrement des valeurs ---
    EnregistreValeursfenetre

    '--- ne plus permettre la validation ---
    PermettreValidation False
    VSJours.SetFocus
    
    '--- RAZ de la variable de comptage avant sortie obligatoire ---
    TempsSortieObligatoire = 0
    
    '--- curseur de la souris ---
    SourisEnAttente False

    '--- affectation ---
    CBQuitter.Enabled = True

End Sub

'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Rôle      : Fixe les couleurs de traçage
' Entrées :
' Retours :
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Private Sub CouleursDeTraçage()
    
    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next

    '--- modes de la pompe ---
    PBCyclesPompe(CYCLES_POMPES.CP_ARRET).BackColor = COULEURS_ECHELLES_GRAPHIQUES.C_ARRET_POMPE
    PBCyclesPompe(CYCLES_POMPES.CP_MARCHE).BackColor = COULEURS_ECHELLES_GRAPHIQUES.C_MARCHE_POMPE

    '--- modes du chauffage ---
    PBModesChauffage(MODES_PRODUCTION.M_ARRET).BackColor = COULEURS_ECHELLES_GRAPHIQUES.C_MODE_ARRET
    PBModesChauffage(MODES_PRODUCTION.M_VEILLE).BackColor = COULEURS_ECHELLES_GRAPHIQUES.C_MODE_VEILLE
    PBModesChauffage(MODES_PRODUCTION.M_PRODUCTION).BackColor = COULEURS_ECHELLES_GRAPHIQUES.C_MODE_PRODUCTION

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
        PremiereActivation = True
        Me.Refresh
    End If

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    
    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next

    '--- action en fonction des touches ---
    Select Case KeyCode
        
        Case vbKeyF1 To vbKeyF11
            '--- touches de fonctions ---
            OccFSynoptique.SetFocus
            Call OccFSynoptique.GestionTouches(KeyCode, Shift)
        
        Case vbKeyF12
            '--- acquittement des alarmes ---
            Call OccFSynoptique.GestionTouches(KeyCode, Shift)
        
        Case vbKeyEscape
            '--- touche échap ---
            CBQuitter_Click
        
        Case vbKeyPageUp
            '--- saut de page arrière ---
            If VSJours.value > VSJours.min Then
                VSJours.value = Pred(VSJours.value)
            End If
            KeyCode = 0
        
        Case vbKeyPageDown
            '--- saut de page avant ---
            If VSJours.value < VSJours.Max Then
                VSJours.value = Succ(VSJours.value)
            End If
            KeyCode = 0
        
        Case Else

    End Select

End Sub

'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Rôle      : Détermine la zone d'appareillage concerné
' Entrées :
' Retours :
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Private Sub ChoixAppareillage()
    
    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next
    
    '--- affichage et effacement des zones concernées ---
    With VManipsProgCyclique

        '--- pompe ---
        PBCyclesPompe(CYCLES_POMPES.CP_MARCHE).Visible = Not (.AppareillageConcerne)
        PBCyclesPompe(CYCLES_POMPES.CP_ARRET).Visible = Not (.AppareillageConcerne)
        OBCyclesPompe(CYCLES_POMPES.CP_MARCHE).Visible = Not (.AppareillageConcerne)
        OBCyclesPompe(CYCLES_POMPES.CP_ARRET).Visible = Not (.AppareillageConcerne)

        '--- chauffage ---
        PBModesChauffage(MODES_PRODUCTION.M_ARRET).Visible = .AppareillageConcerne
        PBModesChauffage(MODES_PRODUCTION.M_VEILLE).Visible = .AppareillageConcerne
        PBModesChauffage(MODES_PRODUCTION.M_PRODUCTION).Visible = .AppareillageConcerne
        OBModesChauffage(MODES_PRODUCTION.M_ARRET).Visible = .AppareillageConcerne
        OBModesChauffage(MODES_PRODUCTION.M_VEILLE).Visible = .AppareillageConcerne
        OBModesChauffage(MODES_PRODUCTION.M_PRODUCTION).Visible = .AppareillageConcerne

    End With

End Sub

'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Rôle      : Visualisation des différents états du programmateur cyclique
' Entrées :
' Retours :
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Private Sub EtatsProgCyclique()
    
    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next

    '--- analyse du changement de jour ---
    'If TypePC = TYPES_PC.PC_SUR_LIGNE Then
        With LAvertissement
            If MemDateProgCyclique <> DateMaintenant Then
                .Visible = True
                .Refresh
                Call Sleep(3000)
                DechargeFenetre
            Else
                .Visible = False
            End If
        End With
    'End If

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

    '--- fixe les couleurs de traçage ---
    CouleursDeTraçage

    '--- zone à griser en fonction de la cuve ---
    For a = PBEchelle24H.LBound To PBEchelle24H.UBound
        Select Case a
        
            Case CUVES_REGULATION.C_DEC, CUVES_REGULATION.C_C31, CUVES_REGULATION.C_C32 'CUVES_REGULATION.C_SAT,
                '--- cuves avec pompe ---
                Set PBEchelle24H(a).Picture = TImgEchelles24H(ECHELLES_24H.E_POMPE_CHAUFFAGE)
    
            'Case CUVES_REGULATION.C_C37
                '--- cas de l'étuve ---
                'Set PBEchelle24H(a).Picture = TImgEchelles24H(ECHELLES_24H.E_VENTILATION_CHAUFFAGE)
            
            Case Else
                '--- cuves sans pompe ---
                Set PBEchelle24H(a).Picture = TImgEchelles24H(ECHELLES_24H.E_CHAUFFAGE)
    
        End Select
    
    Next a
    
    '--- sélecteurs ---
    With VManipsProgCyclique
        OBCyclesPompe(.CyclesPompe).value = True
        OBModesChauffage(.ModesChauffage).value = True
    End With

    '--- choix de l'appareillage ---
    ChoixAppareillage

    '--- copie du programmateur cyclique ---
    CopieProgCyclique

    '--- affichage du jour géré ---
    AffichageJourGere

    '--- lecture des types de journées ---
    LectureTypesDeJournees

    '--- lecture du cycle par cuve ---
    LectureCycleParCuve

    '--- affectation ---
    VSJours.Enabled = True
    For a = OBJours.LBound To OBJours.UBound
        OBJours(a).Enabled = True
    Next a
    CBValider.Enabled = False
    CBAnnuler.Enabled = False
    
    '--- autoriser les évènements ---
    InterdireEvenements = False
    
End Sub

'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Rôle      : Lecture du cycle pour chaque cuve
' Entrées :
' Retours :
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Private Sub LectureCycleParCuve(Optional NumCuve As Integer = 0)
    
    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next

    '--- déclaration ---
    Dim a As Integer, _
            b As Integer
    Dim HeureDebutTop As Integer, _
            MinuteDebutTop As Integer, _
            HeureFinTop As Integer, _
            MinuteFinTop As Integer
    Dim CouleurRemplissagePompe As Long, _
            CouleurRemplissageChauffage As Long
    Dim XDebut As Single, _
            XFin As Single

    '--- analyse de toutes les échelles graphiques ---
    For a = CUVES_REGULATION.C_C00 To DERNIERE_CUV_REGULATION

        If a = NumCuve Or NumCuve = 0 Then

            '--- analyse de la présence de la pompe en fonction de la cuve ---
            Select Case a
    
                Case CUVES_REGULATION.C_DEC, CUVES_REGULATION.C_C31, CUVES_REGULATION.C_C32  'CUVES_REGULATION.C_SAT,
                    '--- cuves avec pompe ---
                    Set PBEchelle24H(a).Picture = TImgEchelles24H(ECHELLES_24H.E_POMPE_CHAUFFAGE)
    
                'Case CUVES_REGULATION.C_C37
                    '--- cas de l'étuve --
                    'Set PBEchelle24H(a).Picture = TImgEchelles24H(ECHELLES_24H.E_VENTILATION_CHAUFFAGE)
            
                Case Else
                    '--- cuves sans pompe ---
                    Set PBEchelle24H(a).Picture = TImgEchelles24H(ECHELLES_24H.E_CHAUFFAGE)
    
            End Select
            
            With TCopieProgCyclique(VSJours.value, a)
                
                Select Case a
    
                    Case CUVES_REGULATION.C_DEC, CUVES_REGULATION.C_C31, CUVES_REGULATION.C_C32 'CUVES_REGULATION.C_SAT,
                        '--- cuves avec pompe ---
                        For b = 1 To NBR_TOPS_POSSIBLES

                            '--- contrôle ---
                            If .TTopsDebutPompe(b) = String(14, 0) Or _
                                Trim(.TTopsDebutPompe(b)) = "" Then
                                    Exit For
                            End If

                            '--- heure, minute de début et de fin pour la pompe ---
                            HeureDebutTop = Val(Mid(.TTopsDebutPompe(b), 9, 2))
                            MinuteDebutTop = Val(Mid(.TTopsDebutPompe(b), 11, 2))
                            HeureFinTop = Val(Mid(.TTopsFinPompe(b), 9, 2))
                            MinuteFinTop = Val(Mid(.TTopsFinPompe(b), 11, 2))

                            '--- calcul du point de début et de fin de tracer pour la pompe ---
                            XDebut = X_MINI_POMPE + HeureDebutTop * (2 * LARGEUR_SEGMENT)
                            XDebut = XDebut - LARGEUR_SEGMENT * (MinuteDebutTop = 30)
                            XFin = X_MINI_POMPE + HeureFinTop * (2 * LARGEUR_SEGMENT)
                            XFin = XFin - LARGEUR_SEGMENT * (MinuteFinTop = 29) - (2 * LARGEUR_SEGMENT) * (MinuteFinTop = 59)

                            '--- couleur de remplissage pour la pompe ---
                            Select Case .TCyclesPompe(b)
                                Case CYCLES_POMPES.CP_ARRET: CouleurRemplissagePompe = COULEURS_ECHELLES_GRAPHIQUES.C_ARRET_POMPE
                                Case CYCLES_POMPES.CP_MARCHE: CouleurRemplissagePompe = COULEURS_ECHELLES_GRAPHIQUES.C_MARCHE_POMPE
                                Case Else: CouleurRemplissagePompe = COULEURS_ECHELLES_GRAPHIQUES.C_ARRET_POMPE
                            End Select

                            '--- tracer de la pompe ---
                            If XFin > X_MAXI_POMPE Then XFin = X_MAXI_POMPE
                            PBEchelle24H(a).Line (XDebut, Y_MINI_POMPE)-(XFin, Y_MAXI_POMPE), CouleurRemplissagePompe, BF

                        Next b
                
                    Case Else

                End Select

                For b = 1 To NBR_TOPS_POSSIBLES

                    '--- contrôle ---
                    If .TTopsDebutChauffage(b) = String(14, 0) Or _
                        Trim(.TTopsDebutChauffage(b)) = "" Then
                            Exit For
                    End If

                    '--- heure, minute de début et de fin pour le chauffage ---
                    HeureDebutTop = Val(Mid(.TTopsDebutChauffage(b), 9, 2))
                    MinuteDebutTop = Val(Mid(.TTopsDebutChauffage(b), 11, 2))
                    HeureFinTop = Val(Mid(.TTopsFinChauffage(b), 9, 2))
                    MinuteFinTop = Val(Mid(.TTopsFinChauffage(b), 11, 2))

                    '--- calcul du point de début et de fin de tracer pour le chauffage ---
                    XDebut = X_MINI_CHAUFFAGE + HeureDebutTop * (2 * LARGEUR_SEGMENT)
                    XDebut = XDebut - LARGEUR_SEGMENT * (MinuteDebutTop = 30)
                    XFin = X_MINI_CHAUFFAGE + HeureFinTop * (2 * LARGEUR_SEGMENT)
                    XFin = XFin - LARGEUR_SEGMENT * (MinuteFinTop = 29) - (2 * LARGEUR_SEGMENT) * (MinuteFinTop = 59)

                    '--- couleur de remplissage pour le chauffage ---
                    Select Case .TModesChauffage(b)
                        Case MODES_PRODUCTION.M_ARRET: CouleurRemplissageChauffage = COULEURS_ECHELLES_GRAPHIQUES.C_MODE_ARRET
                        Case MODES_PRODUCTION.M_VEILLE: CouleurRemplissageChauffage = COULEURS_ECHELLES_GRAPHIQUES.C_MODE_VEILLE
                        Case MODES_PRODUCTION.M_PRODUCTION: CouleurRemplissageChauffage = COULEURS_ECHELLES_GRAPHIQUES.C_MODE_PRODUCTION
                        Case Else: CouleurRemplissageChauffage = COULEURS_ECHELLES_GRAPHIQUES.C_MODE_ARRET
                    End Select

                    '--- tracer du chauffage ---
                    If XFin > X_MAXI_CHAUFFAGE Then XFin = X_MAXI_CHAUFFAGE
                    PBEchelle24H(a).Line (XDebut, Y_MINI_CHAUFFAGE)-(XFin, Y_MAXI_CHAUFFAGE), CouleurRemplissageChauffage, BF
                
                Next b

            End With

        End If

    Next a
            
End Sub

'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Rôle      : Enregistre les dernières manipulations et valeurs de la fenêtre
' Entrées :
' Retours :
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Private Sub EnregistreValeursfenetre()
    On Error Resume Next
    If VSJours.Enabled = False Then EnregistreCycleParCuve
    RestaureProgCyclique
    SauveProgCyclique
    
End Sub

'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Rôle      : Enregistre le cycle pour chaque cuve
' Entrées :
' Retours :
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Private Sub EnregistreCycleParCuve(Optional NumCuve As Integer = 0)
    
    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next

    '--- déclaration ---
    Dim a As Integer, _
            b As Integer, _
            c As Integer
    Dim PointeurPompe As Integer, _
            PointeurChauffage As Integer
    Dim CouleurPointPompe As Long, _
            CouleurPointChauffage As Long, _
            MemCouleurPointPompe As Long, _
            MemCouleurPointChauffage As Long
    Dim xPoint As Single
    Dim TopDebutCycle As String * 14, _
           TopFinCycle As String * 14

    '--- analyse de toutes les échelles graphiques ---
    For a = CUVES_REGULATION.C_C00 To DERNIERE_CUV_REGULATION

        If a = NumCuve Or NumCuve = 0 Then

            '--- affectation ---
            PointeurPompe = 0
            MemCouleurPointPompe = -1                         'forcer un mode inexistant pour la comparaison
            PointeurChauffage = 0
            MemCouleurPointChauffage = -1                       'forcer un mode inexistant pour la comparaison

            With TCopieProgCyclique(VSJours.value, a)

                '--- analyse des graphiques ---
                For b = 0 To 23                                                     'cycle des heures
                    For c = 0 To 30 Step 30                                    'cycle des demi-heures

                        '--- affectation ---
                        TopDebutCycle = LaDateATraiter & Right("0" & CStr(b), 2) & Right("0" & CStr(c), 2) & "00"
                        xPoint = 3 + (2 * LARGEUR_SEGMENT) * b + LARGEUR_SEGMENT * Abs(c = 30)
                        
                        CouleurPointPompe = PBEchelle24H(a).Point(xPoint, Y_MINI_POMPE)
                        CouleurPointChauffage = PBEchelle24H(a).Point(xPoint, Y_MINI_CHAUFFAGE)
                        
                        '--- vérification avec la couleur pour la pompe ---
                        'SZP   a = CUVES_REGULATION.C_SAT Or
                        If CouleurPointPompe <> MemCouleurPointPompe And _
                            (a = CUVES_REGULATION.C_DEC Or _
                            a = CUVES_REGULATION.C_C31 Or _
                            a = CUVES_REGULATION.C_C32) Then
                            
                            '--- heures pour la pompe ---
                            If PointeurPompe > 0 Then
                                .TTopsFinPompe(PointeurPompe) = TopFinCycle
                            End If
                            Inc PointeurPompe
                            .TTopsDebutPompe(PointeurPompe) = TopDebutCycle

                            '--- mode  pour la pompe ---
                            Select Case CouleurPointPompe
                                Case COULEURS_ECHELLES_GRAPHIQUES.C_ARRET_POMPE: .TCyclesPompe(PointeurPompe) = CYCLES_POMPES.CP_ARRET
                                Case COULEURS_ECHELLES_GRAPHIQUES.C_MARCHE_POMPE: .TCyclesPompe(PointeurPompe) = CYCLES_POMPES.CP_MARCHE
                                Case Else
                                    Bidon = MessageErreur(TITRE_MESSAGES, MESSAGE_400)
                                    Exit Sub
                            End Select

                            '--- mise en mémoire de la couleur ---
                            MemCouleurPointPompe = CouleurPointPompe

                        End If

                        '--- vérification avec la couleur pour le chauffage ---
                        If CouleurPointChauffage <> MemCouleurPointChauffage Then

                            '--- heures pour le chauffage ---
                            If PointeurChauffage > 0 Then
                                .TTopsFinChauffage(PointeurChauffage) = TopFinCycle
                            End If
                            Inc PointeurChauffage
                            .TTopsDebutChauffage(PointeurChauffage) = TopDebutCycle

                            '--- mode  pour le chauffage ---
                            Select Case CouleurPointChauffage
                                Case COULEURS_ECHELLES_GRAPHIQUES.C_MODE_ARRET: .TModesChauffage(PointeurChauffage) = MODES_PRODUCTION.M_ARRET
                                Case COULEURS_ECHELLES_GRAPHIQUES.C_MODE_VEILLE: .TModesChauffage(PointeurChauffage) = MODES_PRODUCTION.M_VEILLE
                                Case COULEURS_ECHELLES_GRAPHIQUES.C_MODE_PRODUCTION: .TModesChauffage(PointeurChauffage) = MODES_PRODUCTION.M_PRODUCTION
                                Case Else: Bidon = MessageErreur(TITRE_MESSAGES, MESSAGE_400)
                            End Select

                            '--- mise en mémoire de la couleur ---
                            MemCouleurPointChauffage = CouleurPointChauffage

                        End If

                        '--- calcul de l'heure de fin de cycle ---
                        If c = 0 Then
                            TopFinCycle = LaDateATraiter & Right("0" & CStr(b), 2) & "2959"
                        Else
                            TopFinCycle = LaDateATraiter & Right("0" & CStr(b), 2) & "5959"
                        End If

                    Next c
                Next b

                '--- forçage de la dernière heure ---
                If PointeurPompe > 0 And PBTousCyclesPompe.Visible = True Then
                    .TTopsFinPompe(PointeurPompe) = LaDateATraiter & "235959"
                End If
                If PointeurChauffage > 0 Then
                    .TTopsFinChauffage(PointeurChauffage) = LaDateATraiter & "235959"
                End If

                '--- vidage des champs inutilisés de la pompe ---
                If PBTousCyclesPompe.Visible = True Then
                    For b = Succ(PointeurPompe) To NBR_TOPS_POSSIBLES
                        .TTopsDebutPompe(b) = ""
                        .TTopsFinPompe(b) = ""
                        .TCyclesPompe(b) = CYCLES_POMPES.CP_ARRET
                    Next b
                End If

                '--- vidage des champs inutilisés du chauffage ---
                For b = Succ(PointeurChauffage) To NBR_TOPS_POSSIBLES
                    .TTopsDebutChauffage(b) = ""
                    .TTopsFinChauffage(b) = ""
                    .TModesChauffage(b) = MODES_PRODUCTION.M_ARRET
                Next b

            End With

        End If

    Next a

End Sub

'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Rôle      : Affiche le jour géré en haut de l'écran
' Entrées :
' Retours :
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Private Sub AffichageJourGere()
    
    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next

    '--- déclaration ---
    Dim DateATraiter As String

    '--- date en cours ---
    DateATraiter = DateAdd("d", Now, Pred(Me.VSJours.value))
    LRenseignementsFenetre.Caption = StrConv(Format(CDate(DateATraiter), "Long Date"), vbProperCase)
    LaDateATraiter = Format(DateATraiter, "yyyymmdd")
    
End Sub

Private Sub Form_Resize()
    
    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next
    
    '--- déclaration ---
    
    '--- zone mére et fille du déplacement de la fenetre ---
    PBDeplacementFenetre(ZONES_DEPLACEMENT_FENETRE.Z_MERE).Height = Abs(Me.ScaleHeight - PBRenseignementsFenetre.Height)
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
    PBDeplacementFenetre(ZONES_DEPLACEMENT_FENETRE.Z_FILLE).Left = -HSDeplacementFenetre.value
End Sub

Private Sub LRenseignementsFenetre_DblClick()
    On Error Resume Next
    If Me.WindowState = vbMaximized Then
        Me.WindowState = vbNormal
    Else
        Me.WindowState = vbMaximized
    End If
End Sub

Private Sub LTitreTousModesChauffage_Click()
    On Error Resume Next
    VManipsProgCyclique.AppareillageConcerne = True
    ChoixAppareillage
End Sub

Private Sub LTitreTousCyclesPompe_Click()
    On Error Resume Next
    VManipsProgCyclique.AppareillageConcerne = False
    ChoixAppareillage
End Sub

Private Sub OBJours_Click(Index As Integer)
    
    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next

    '--- déclaration ---
    Dim OCCBoutonOption As OptionButton

    For Each OCCBoutonOption In OBJours
        With OCCBoutonOption
            If .value = True Then
                .BackColor = COULEURS.ROUGE_3
                .ForeColor = COULEURS.NOIR
            Else
                .BackColor = COULEURS.ORANGE_1
                .ForeColor = COULEURS.NOIR
            End If
        End With
    Next
    
    '--- changement de l'ascenseur (si click direct dans un des boutons) ---
    If InterdireEvenements = False Then
        VSJours.value = Index
    End If

End Sub

Private Sub OBCyclesPompe_Click(Index As Integer)
    
    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next
    
    '--- déclaration ---
    Dim OCCBoutonOption As OptionButton

    '--- changement de couleur de l'outil sélectionné ---
    For Each OCCBoutonOption In OBCyclesPompe
        With OCCBoutonOption
            If .value = True Then
                .BackColor = COULEURS.ROUGE_3
                .ForeColor = COULEURS.NOIR
            Else
                .BackColor = COULEURS.ORANGE_1
                .ForeColor = COULEURS.NOIR
            End If
        End With
    Next
    
    '--- affectation ---
    VManipsProgCyclique.CyclesPompe = Index
    CouleurPompe = PBCyclesPompe(Index).BackColor

End Sub

Private Sub OBModesChauffage_Click(Index As Integer)
     
     '--- aiguillage en cas d'erreurs ---
    On Error Resume Next
    
    '--- déclaration ---
    Dim OCCBoutonOption As OptionButton

    '--- changement de couleur de l'outil sélectionné ---
    For Each OCCBoutonOption In OBModesChauffage
        With OCCBoutonOption
            If .value = True Then
                .BackColor = COULEURS.ROUGE_3
                .ForeColor = COULEURS.NOIR
            Else
                .BackColor = COULEURS.ORANGE_1
                .ForeColor = COULEURS.NOIR
            End If
        End With
    Next
   
    '--- affectation ---
    VManipsProgCyclique.ModesChauffage = Index
    CouleurChauffage = PBModesChauffage(Index).BackColor

End Sub

Private Sub OBTypesJourneesIdx01_Click(Index As Integer)
    On Error Resume Next
    ChangementTypesJournees CUVES_REGULATION.C_C00, Index, OBTypesJourneesIdx01
End Sub

Private Sub OBTypesJourneesIdx02_Click(Index As Integer)
    On Error Resume Next
    ChangementTypesJournees CUVES_REGULATION.C_DEC, Index, OBTypesJourneesIdx02
End Sub

Private Sub OBTypesJourneesIdx03_Click(Index As Integer)
    On Error Resume Next
    ChangementTypesJournees CUVES_REGULATION.C_C07, Index, OBTypesJourneesIdx03
End Sub

Private Sub OBTypesJourneesIdx04_Click(Index As Integer)
    On Error Resume Next
    ChangementTypesJournees CUVES_REGULATION.C_C13, Index, OBTypesJourneesIdx04
End Sub

Private Sub OBTypesJourneesIdx05_Click(Index As Integer)
    On Error Resume Next
    ChangementTypesJournees CUVES_REGULATION.C_C14, Index, OBTypesJourneesIdx05
End Sub

Private Sub OBTypesJourneesIdx06_Click(Index As Integer)
    On Error Resume Next
    ChangementTypesJournees CUVES_REGULATION.C_C15, Index, OBTypesJourneesIdx06
End Sub

Private Sub OBTypesJourneesIdx07_Click(Index As Integer)
    On Error Resume Next
    ChangementTypesJournees CUVES_REGULATION.C_C22, Index, OBTypesJourneesIdx07
End Sub

Private Sub OBTypesJourneesIdx08_Click(Index As Integer)
    On Error Resume Next
    ChangementTypesJournees CUVES_REGULATION.C_C27, Index, OBTypesJourneesIdx08
End Sub

Private Sub OBTypesJourneesIdx09_Click(Index As Integer)
    On Error Resume Next
    ChangementTypesJournees CUVES_REGULATION.C_C28, Index, OBTypesJourneesIdx09
End Sub

Private Sub OBTypesJourneesIdx10_Click(Index As Integer)
    On Error Resume Next
    ChangementTypesJournees CUVES_REGULATION.C_C31, Index, OBTypesJourneesIdx10
End Sub

Private Sub OBTypesJourneesIdx11_Click(Index As Integer)
    On Error Resume Next
    ChangementTypesJournees CUVES_REGULATION.C_C32, Index, OBTypesJourneesIdx11
End Sub



Private Sub PBBoutons_Resize()
    
    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next
    
    '--- calculs de l'emplacement des boutons ---
    CBQuitter.Left = PBBoutons.ScaleWidth - MARGES.M_BORD_DROIT - CBQuitter.Width
    CBValider.Left = CBQuitter.Left - MARGES.M_ENTRE_BOUTONS - CBValider.Width
    CBAnnuler.Left = CBValider.Left - MARGES.M_ENTRE_BOUTONS - CBAnnuler.Width
    
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
            
            End With
                   
        End If
        
    Else
        
        '--- la zone fille a bougé ---
        With PBDeplacementFenetre(ZONES_DEPLACEMENT_FENETRE.Z_FILLE)
        
            '--- graphismes de programmation ---
            PBProgrammation.Left = 0
            PBProgrammation.Top = 0
            PBProgrammation.Height = .ScaleHeight - PBBoutons.Height

            '--- les modes de programmation ---
            PBModesProgrammation.Left = PBProgrammation.Left + PBProgrammation.Width
            PBModesProgrammation.Top = PBProgrammation.Top
            PBModesProgrammation.Width = Abs(.ScaleWidth - PBProgrammation.Width)
            PBModesProgrammation.Height = .ScaleHeight - PBBoutons.Height

        End With
    
    End If
            
    '--- valeur des curseurs ---
    If Me.WindowState <> vbMaximized And Me.WindowState <> vbMinimized Then
        HSDeplacementFenetre.Max = PBDeplacementFenetre(ZONES_DEPLACEMENT_FENETRE.Z_FILLE).Width - _
                                                         PBDeplacementFenetre(ZONES_DEPLACEMENT_FENETRE.Z_MERE).Width
        VSDeplacementFenetre.Max = PBDeplacementFenetre(ZONES_DEPLACEMENT_FENETRE.Z_FILLE).Height - _
                                                        PBDeplacementFenetre(ZONES_DEPLACEMENT_FENETRE.Z_MERE).Height
    End If

End Sub

Private Sub PBEchelle24H_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next
    
    '--- déclaration ---
    Dim a As Integer                'pour les boucles FOR...NEXT
    Dim SegmentDepart As Integer, _
            ZoneTravail As Integer

    If Button = vbLeftButton Then

        '--- recherche de la zone de travail ---
        ZoneTravail = 0
        If TTypesJourneesEnCours(Index) <> JOURNEES_TYPES.J_ARRET Then
            If VManipsProgCyclique.AppareillageConcerne = False Then
                If X >= X_MINI_POMPE And X <= X_MAXI_POMPE And Y >= Y_MINI_POMPE And Y <= Y_MAXI_POMPE Then
                    Select Case Index
                        Case CUVES_REGULATION.C_DEC, CUVES_REGULATION.C_C31, CUVES_REGULATION.C_C32: ZoneTravail = 1                             'cuves avec pompe SZP  CUVES_REGULATION.C_SAT,
                        Case Else: ZoneTravail = 0
                    End Select
                End If
            Else
                If X >= X_MINI_CHAUFFAGE And X <= X_MAXI_CHAUFFAGE And Y >= Y_MINI_CHAUFFAGE And Y <= Y_MAXI_CHAUFFAGE Then
                    ZoneTravail = 2
                End If
            End If
        End If

        '--- analyse ---
        If ZoneTravail > 0 Then

            '--- neutralisation de l'ascenseur des jours pour obliger à valider, cela permet de savoir _
                  si il y a eu une modification d'un des graphiques ---
            VSJours.Enabled = False
            For a = OBJours.LBound To OBJours.UBound
                OBJours(a).Enabled = False
            Next a
            CBValider.Enabled = True
            CBAnnuler.Enabled = True
            
            '--- calcul du point de départ ---
            SegmentDepart = Int(X / LARGEUR_SEGMENT)
            XDepart = (SegmentDepart * LARGEUR_SEGMENT) + 2

            '--- tracer de la zone de départ ---
            Select Case ZoneTravail

                Case 1
                    '--- zone Pompe ---
                    If SegmentDepart < 48 Then
                        PBEchelle24H(Index).Line (XDepart, Y_MINI_POMPE)-(XDepart + LARGEUR_SEGMENT, Y_MAXI_POMPE), CouleurPompe, BF
                    Else
                        PBEchelle24H(Index).Line (XDepart, Y_MINI_POMPE)-(XDepart - LARGEUR_SEGMENT, Y_MAXI_POMPE), CouleurPompe, BF
                    End If

                Case 2
                    '--- zone chauffage ---
                    If SegmentDepart < 48 Then
                        PBEchelle24H(Index).Line (XDepart, Y_MINI_CHAUFFAGE)-(XDepart + LARGEUR_SEGMENT, Y_MAXI_CHAUFFAGE), CouleurChauffage, BF
                    Else
                        PBEchelle24H(Index).Line (XDepart, Y_MINI_CHAUFFAGE)-(XDepart - LARGEUR_SEGMENT, Y_MAXI_CHAUFFAGE), CouleurChauffage, BF
                    End If

                Case Else
            End Select

            '--- affectation ---
            CBValider.Enabled = True
            CBAnnuler.Enabled = True
        
        End If

    End If

End Sub

Private Sub PBEchelle24H_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
      
    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next
  
    '--- déclaration ---
    Dim ZoneTravail As Integer

    If Button = vbLeftButton Then

        '--- recherche de la zone de travail ---
        ZoneTravail = 0
        If TTypesJourneesEnCours(Index) <> JOURNEES_TYPES.J_ARRET Then
            If VManipsProgCyclique.AppareillageConcerne = False Then
                If X >= X_MINI_POMPE And X <= X_MAXI_POMPE And Y >= Y_MINI_POMPE And Y <= Y_MAXI_POMPE Then
                    Select Case Index
                        Case CUVES_REGULATION.C_DEC, CUVES_REGULATION.C_C31, CUVES_REGULATION.C_C32: ZoneTravail = 1                             'cuves avec pompe CUVES_REGULATION.C_SAT,
                        Case Else: ZoneTravail = 0
                    End Select
                End If
            Else
                If X >= X_MINI_CHAUFFAGE And X <= X_MAXI_CHAUFFAGE And Y >= Y_MINI_CHAUFFAGE And Y <= Y_MAXI_CHAUFFAGE Then
                    ZoneTravail = 2
                End If
            End If
        End If

        If (X <> MemX Or Y <> MemY) And ZoneTravail > 0 Then

            '--- calcul du point d'arrivée ---
            XArrivee = (Int(X / LARGEUR_SEGMENT) * LARGEUR_SEGMENT) + 2

            '--- tracer de la zone de départ ---
            Select Case ZoneTravail

                Case 1
                    '--- zone Pompe ---
                    PBEchelle24H(Index).Line (XDepart, Y_MINI_POMPE)-(XArrivee, Y_MAXI_POMPE), CouleurPompe, BF

                Case 2
                    '--- zone chauffage ---
                    PBEchelle24H(Index).Line (XDepart, Y_MINI_CHAUFFAGE)-(XArrivee, Y_MAXI_CHAUFFAGE), CouleurChauffage, BF

                Case Else
            End Select

            '--- mémorisation ---
            MemX = XArrivee
            MemY = Y

        End If

    End If
            
End Sub

Private Sub PBEchelle24H_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next
    
    '--- affectation ---
    MemX = 0
    MemY = 0

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

Private Sub PBCyclesPompe_Click(Index As Integer)
    On Error Resume Next
    OBCyclesPompe(Index).value = True
End Sub

Private Sub PBTousModesChauffage_Click()
    On Error Resume Next
    VManipsProgCyclique.AppareillageConcerne = True
    ChoixAppareillage
End Sub

Private Sub PBTousCyclesPompe_Click()
    On Error Resume Next
    VManipsProgCyclique.AppareillageConcerne = False
    ChoixAppareillage
End Sub

Private Sub TimerProgCyclique_Timer()
    
    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next

    '--- appel de la routine ---
    TimerProgCyclique.Enabled = False
    EtatsProgCyclique
    TimerProgCyclique.Enabled = True
    
    '--- bip de passage dans la routine UNIQUEMENT POUR LES TESTS ---
    If PROGRAMME_AVEC_AUTOMATE = False Then Beep

End Sub

Private Sub TimerSortieObligatoire_Timer()
    
    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next

    '--- incrémentation du compteur ---
    Inc TempsSortieObligatoire
    
    '--- sortie obligatoire au bout d'une heure ---
    If TempsSortieObligatoire >= 60 Then
        DechargeFenetre
    End If

End Sub

Private Sub VSDeplacementFENETRE_Change()
    On Error Resume Next
    PBDeplacementFenetre(ZONES_DEPLACEMENT_FENETRE.Z_FILLE).Top = -VSDeplacementFenetre.value
End Sub

Private Sub VSJours_Change()
    
    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next
    
    '--- interdire certains évènements ---
    InterdireEvenements = True
    
    '--- réaffichage ---
    OBJours(VSJours.value).value = True                 'sélection de la journée
    AffichageJourGere
    LectureCycleParCuve
    LectureTypesDeJournees

    '--- autoriser les évènements ---
    InterdireEvenements = False

End Sub

'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Rôle      : Lecture des types de journées
' Entrées :
' Retours :
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Private Sub LectureTypesDeJournees()
    
    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next

    '--- déclaration ---
    Dim a As Integer

    '--- affichage ---
    For a = CUVES_REGULATION.C_C00 To DERNIERE_CUV_REGULATION
        With TCopieProgCyclique(VSJours.value, a)
            Select Case a
                Case CUVES_REGULATION.C_C00: OBTypesJourneesIdx01(.TypeDeJournee).value = True
                'Case CUVES_REGULATION.C_SAT: OBTypesJourneesIdx02(.TypeDeJournee).Value = True
                Case CUVES_REGULATION.C_DEC: OBTypesJourneesIdx02(.TypeDeJournee).value = True
                'Case CUVES_REGULATION.C_C03: OBTypesJourneesIdx04(.TypeDeJournee).Value = True
                'Case CUVES_REGULATION.C_C05: OBTypesJourneesIdx05(.TypeDeJournee).Value = True
                'Case CUVES_REGULATION.C_C06: OBTypesJourneesIdx06(.TypeDeJournee).Value = True
                Case CUVES_REGULATION.C_C07: OBTypesJourneesIdx03(.TypeDeJournee).value = True
                Case CUVES_REGULATION.C_C13: OBTypesJourneesIdx04(.TypeDeJournee).value = True
                Case CUVES_REGULATION.C_C14: OBTypesJourneesIdx05(.TypeDeJournee).value = True
                Case CUVES_REGULATION.C_C15: OBTypesJourneesIdx06(.TypeDeJournee).value = True
                'Case CUVES_REGULATION.C_C16: OBTypesJourneesIdx11(.TypeDeJournee).Value = True
                'Case CUVES_REGULATION.C_C19: OBTypesJourneesIdx12(.TypeDeJournee).Value = True
                Case CUVES_REGULATION.C_C22: OBTypesJourneesIdx07(.TypeDeJournee).value = True
                Case CUVES_REGULATION.C_C27: OBTypesJourneesIdx08(.TypeDeJournee).value = True
                Case CUVES_REGULATION.C_C28: OBTypesJourneesIdx09(.TypeDeJournee).value = True
                Case CUVES_REGULATION.C_C31: OBTypesJourneesIdx10(.TypeDeJournee).value = True
                Case CUVES_REGULATION.C_C32: OBTypesJourneesIdx11(.TypeDeJournee).value = True
                'Case CUVES_REGULATION.C_C33: OBTypesJourneesIdx18(.TypeDeJournee).Value = True
                'Case CUVES_REGULATION.C_MAX: OBTypesJourneesIdx19(.TypeDeJournee).Value = True
                'Case CUVES_REGULATION.C_C35: OBTypesJourneesIdx20(.TypeDeJournee).Value = True
                'Case CUVES_REGULATION.C_C36: OBTypesJourneesIdx21(.TypeDeJournee).Value = True
                'Case CUVES_REGULATION.C_C37: OBTypesJourneesIdx22(.TypeDeJournee).Value = True
                'Case CUVES_REGULATION.C_MAX: OBTypesJourneesIdx23(.TypeDeJournee).Value = True
                
                Case Else
            End Select
        End With
    Next a
    
End Sub

'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Rôle      : Copie du programmateur cyclique
' Entrées :
' Retours :
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Private Sub CopieProgCyclique()
    
    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next

    '--- déclaration ---
    Dim a As Integer, _
           b As Integer, _
           c As Integer

    '--- copie du tableau ---
    For a = 1 To NBR_JOURS_PROG_CYCLIQUE
        For b = CUVES_REGULATION.C_C00 To DERNIERE_CUV_REGULATION
            With TCopieProgCyclique(a, b)
                 .TypeDeJournee = TProgCyclique(a, b).TypeDeJournee
                For c = 1 To NBR_TOPS_POSSIBLES
                    .TTopsDebutPompe(c) = TProgCyclique(a, b).TTopsDebutPompe(c)
                    .TTopsFinPompe(c) = TProgCyclique(a, b).TTopsFinPompe(c)
                    .TCyclesPompe(c) = TProgCyclique(a, b).TCyclesPompe(c)
                    .TTopsDebutChauffage(c) = TProgCyclique(a, b).TTopsDebutChauffage(c)
                    .TTopsFinChauffage(c) = TProgCyclique(a, b).TTopsFinChauffage(c)
                    .TModesChauffage(c) = TProgCyclique(a, b).TModesChauffage(c)
                Next c
            End With
        Next b
    Next a

End Sub

'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Rôle      : Restaure le programmateur cyclique
' Entrées :
' Retours :
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Private Sub RestaureProgCyclique()
    
    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next
    
    '--- déclaration ---
    Dim a As Integer, _
           b As Integer, _
           c As Integer

    '--- copie du tableau ---
    For a = 1 To NBR_JOURS_PROG_CYCLIQUE
        For b = CUVES_REGULATION.C_C00 To DERNIERE_CUV_REGULATION
            With TProgCyclique(a, b)
                .TypeDeJournee = TCopieProgCyclique(a, b).TypeDeJournee
                For c = 1 To NBR_TOPS_POSSIBLES
                    .TTopsDebutPompe(c) = TCopieProgCyclique(a, b).TTopsDebutPompe(c)
                    .TTopsFinPompe(c) = TCopieProgCyclique(a, b).TTopsFinPompe(c)
                    .TCyclesPompe(c) = TCopieProgCyclique(a, b).TCyclesPompe(c)
                    .TTopsDebutChauffage(c) = TCopieProgCyclique(a, b).TTopsDebutChauffage(c)
                    .TTopsFinChauffage(c) = TCopieProgCyclique(a, b).TTopsFinChauffage(c)
                    .TModesChauffage(c) = TCopieProgCyclique(a, b).TModesChauffage(c)
                Next c
            End With
        Next b
    Next a

End Sub

'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Rôle      : Permet la validation d'une journée
' Entrées : EtatSouhaite -> FALSE = interdire la validation, TRUE= permettre la validation
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Private Sub PermettreValidation(ByVal EtatSouhaite As Boolean)
    
    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next
    
    '--- déclaration ---
    Dim a As Integer                    'pour les boucles FOR...NEXT

    '--- affichage ou non de certains objets ---
    VSJours.Enabled = Not (EtatSouhaite)
    For a = OBJours.LBound To OBJours.UBound
        OBJours(a).Enabled = Not (EtatSouhaite)
    Next a
    CBValider.Enabled = EtatSouhaite
    CBAnnuler.Enabled = EtatSouhaite

End Sub

'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Rôle      : Initialise la fenêtre (chargement ou en vue de la rendre visible)
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Sub InitialisationFenetre()
    
    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next

    '--- déclaration ---
    Dim a As Integer
    Dim DateATraiter As String                  'représente une des dates à traiter
    
    '--- affectation ---

    '--- divers sur la fenêtre ---
    With Me
        .Caption = TITRE_FENETRE
        .WindowState = vbMaximized
    End With
    
    '--- couleurs des fonds ---
    PBModesProgrammation.Picture = ImgFondVert1
    PBProgrammation.Picture = ImgFondVert1
    
    PBJours.Picture = ImgFondGris1
    PBModeGeneralCuves.Picture = ImgFondGris1
    For a = PBOnglets.LBound To PBOnglets.UBound
        PBOnglets(a).Picture = ImgFondGris1
    Next a
    PBTousModesChauffage.Picture = ImgFondGris1
    PBTousCyclesPompe.Picture = ImgFondGris1
    
    PBBoutons.Picture = ImgFondDesBoutons
    
    '--- calculs de l'emplacement de la barre de défilement ---
    With PBJours
        VSJours.Top = 0
        VSJours.Height = .ScaleHeight
    End With
    
    '--- affichage des journées ---
    For a = OBJours.LBound To OBJours.UBound
        With OBJours(a)
            DateATraiter = DateAdd("d", Now, Pred(a))
            .Caption = StrConv(Format(CDate(DateATraiter), "Long Date"), vbProperCase)
        End With
    Next a
    
    '--- affichage des noms des cuves ---
    For a = LTitresCuves.LBound To LTitresCuves.UBound
        With LTitresCuves(a)
            .Caption = "Cuve " & TEtatsCuves(a).DefinitionCuve.NomCuve & " - " & TEtatsCuves(a).DefinitionCuve.LibelleCuve
        End With
    Next a
    
    '--- onglet par défaut ---
    CTOnglets.CurrTab = ONGLETS_PROGRAMMATEUR_CYCLIQUE.O_PREPARATION
    
    '--- fixer la valeur maxi de l'ascenseur ---
    With VSJours
        .min = 1
        .Max = NBR_JOURS_PROG_CYCLIQUE
        .value = VSJours.value   'uniquement pour déclencher l'évènement
    End With
    
    '--- affectation ---
    FProgCycliqueChargee = True
    
    '--- valeurs de la fenêtre ---
    LectureValeursFenetre
    
    '--- lancement des timers ---
    TimerProgCyclique.Enabled = True
    TimerSortieObligatoire.Enabled = True
    
End Sub

'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Rôle      : Routine permettant le changement des types de journées
' Entrées :                    NumCuve -> Numéro de cuve
'                              JourneeType -> Journée type choisie fonction de l'énumération JOURNEES_TYPES
'                ObjOBTypesJournees -> objet OptionButton des types de journées concernées
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Private Sub ChangementTypesJournees(ByVal NumCuve As Integer, _
                                                                  ByVal JourneeType As JOURNEES_TYPES, _
                                                                  ByRef ObjOBTypesJournees As Object)

    '--- aiguillage en cas d'erreur ---
    On Error Resume Next

    '--- déclaration ---
    Dim a As Integer                                                                'pour les boucles FOR...NEXT
    Dim OCCBoutonOption As OptionButton
    
    '--- changement de couleur de l'outil sélectionné ---
    For Each OCCBoutonOption In ObjOBTypesJournees
        With OCCBoutonOption
            If .value = True Then
                .BackColor = COULEURS.ROUGE_3
                .ForeColor = COULEURS.NOIR
            Else
                .BackColor = COULEURS.ORANGE_1
                .ForeColor = COULEURS.NOIR
            End If
        End With
    Next
    
    '--- mémorisation du type de journée par rapport à la cuve ---
    TTypesJourneesEnCours(NumCuve) = JourneeType
    
    '--- sortie directe si interdiction d'évènements ---
    If InterdireEvenements = True Then Exit Sub
    
    '--- permettre la validation ---
    PermettreValidation True

    '--- transfert dans le tableau ---
    With TCopieProgCyclique(VSJours.value, NumCuve)

        '--- type de journée ---
        .TypeDeJournee = JourneeType

        '--- transfert des nouvelles valeurs ---
        For a = 1 To NBR_TOPS_POSSIBLES

            '--- pompe ---
            .TTopsDebutPompe(a) = TJourneesTypes(NumCuve, JourneeType).TTopsDebutPompe(a)
            If Left(.TTopsDebutPompe(a), 1) = "X" Then
                .TTopsDebutPompe(a) = LaDateATraiter + Mid(.TTopsDebutPompe(a), 9)
            End If
            .TTopsFinPompe(a) = TJourneesTypes(NumCuve, JourneeType).TTopsFinPompe(a)
            If Left(.TTopsFinPompe(a), 1) = "X" Then
                .TTopsFinPompe(a) = LaDateATraiter + Mid(.TTopsFinPompe(a), 9)
            End If
            .TCyclesPompe(a) = TJourneesTypes(NumCuve, JourneeType).TCyclesPompe(a)

            '--- chauffage ---
            .TTopsDebutChauffage(a) = TJourneesTypes(NumCuve, JourneeType).TTopsDebutChauffage(a)
            If Left(.TTopsDebutChauffage(a), 1) = "X" Then
                .TTopsDebutChauffage(a) = LaDateATraiter + Mid(.TTopsDebutChauffage(a), 9)
            End If
            .TTopsFinChauffage(a) = TJourneesTypes(NumCuve, JourneeType).TTopsFinChauffage(a)
            If Left(.TTopsFinChauffage(a), 1) = "X" Then
                .TTopsFinChauffage(a) = LaDateATraiter + Mid(.TTopsFinChauffage(a), 9)
            End If
            .TModesChauffage(a) = TJourneesTypes(NumCuve, JourneeType).TModesChauffage(a)

        Next a

    End With

    '--- réaffichage ---
    LectureCycleParCuve NumCuve
    
    '--- gestion des boutons ---
    CBValider.Enabled = True
    CBAnnuler.Enabled = True

End Sub

'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Rôle      : Décharge la fenêtre
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Private Sub DechargeFenetre()
    
    '--- aiguillage en cas d'erreur ---
    On Error Resume Next
    
    '--- affectation ---
    PremiereActivation = False
    FProgCycliqueChargee = False
    
    '--- curseur souris par défaut ---
    SourisEnAttente False

    '--- neutralisation des timers ---
    With TimerProgCyclique
        .Enabled = False
        .Interval = 0
    End With
    With TimerSortieObligatoire
        .Enabled = False
        .Interval = 0
    End With
    
    '--- déchargement de la fenêtre ---
    Me.Visible = False
    Unload Me
    Set OccFProgrammateurCyclique = Nothing
    
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


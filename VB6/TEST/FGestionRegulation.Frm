VERSION 5.00
Begin VB.Form FGestionRegulation 
   BackColor       =   &H00C0C0C0&
   ClientHeight    =   13050
   ClientLeft      =   150
   ClientTop       =   2130
   ClientWidth     =   13440
   BeginProperty Font 
      Name            =   "Marlett"
      Size            =   7.5
      Charset         =   2
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FGestionRegulation.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   13050
   ScaleWidth      =   13440
   ShowInTaskbar   =   0   'False
   WindowState     =   2  'Maximized
   Begin VB.PictureBox PBBoutons 
      Align           =   2  'Align Bottom
      DrawStyle       =   2  'Dot
      DrawWidth       =   16891
      Height          =   1095
      Left            =   0
      ScaleHeight     =   1035
      ScaleWidth      =   13380
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   11955
      Width           =   13440
      Begin VB.PictureBox PBOutilsDeplacementFenetre 
         BackColor       =   &H00E0E0E0&
         Height          =   1035
         Left            =   0
         ScaleHeight     =   975
         ScaleWidth      =   1155
         TabIndex        =   16
         Top             =   0
         Visible         =   0   'False
         Width           =   1215
         Begin VB.CommandButton CBAgrandirFenetre 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Agrandir"
            DownPicture     =   "FGestionRegulation.frx":014A
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
            Picture         =   "FGestionRegulation.frx":02F4
            Style           =   1  'Graphical
            TabIndex        =   19
            ToolTipText     =   " Agrandissement de la fenêtre "
            Top             =   0
            UseMaskColor    =   -1  'True
            Width           =   900
         End
         Begin VB.VScrollBar VSDeplacementFenetre 
            Height          =   975
            LargeChange     =   300
            Left            =   900
            SmallChange     =   100
            TabIndex        =   18
            Top             =   0
            Width           =   255
         End
         Begin VB.HScrollBar HSDeplacementFenetre 
            Height          =   255
            LargeChange     =   300
            Left            =   0
            SmallChange     =   100
            TabIndex        =   17
            Top             =   720
            Width           =   915
         End
      End
      Begin VB.CommandButton CBAnnuler 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Annuler"
         DownPicture     =   "FGestionRegulation.frx":049E
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
         Picture         =   "FGestionRegulation.frx":0BA0
         Style           =   1  'Graphical
         TabIndex        =   15
         ToolTipText     =   " Annuler les dernières modifications "
         Top             =   120
         UseMaskColor    =   -1  'True
         Width           =   1515
      End
      Begin VB.CommandButton CBValider 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Valider"
         DownPicture     =   "FGestionRegulation.frx":12A2
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
         Picture         =   "FGestionRegulation.frx":19A4
         Style           =   1  'Graphical
         TabIndex        =   14
         ToolTipText     =   " Valider l'enregistrement "
         Top             =   120
         UseMaskColor    =   -1  'True
         Width           =   1515
      End
      Begin VB.CommandButton CBQuitter 
         BackColor       =   &H00FFFFFF&
         Cancel          =   -1  'True
         Caption         =   "&Quitter"
         DownPicture     =   "FGestionRegulation.frx":20A6
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
         Picture         =   "FGestionRegulation.frx":27A8
         Style           =   1  'Graphical
         TabIndex        =   13
         ToolTipText     =   " Quitter cette fenêtre "
         Top             =   120
         UseMaskColor    =   -1  'True
         Width           =   1515
      End
      Begin VB.Timer TimerEtatsRegulation 
         Enabled         =   0   'False
         Interval        =   500
         Left            =   1620
         Top             =   180
      End
      Begin VB.Shape SFocus 
         BorderColor     =   &H000000FF&
         BorderWidth     =   5
         Height          =   285
         Left            =   2220
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
      Picture         =   "FGestionRegulation.frx":2EAA
      ScaleHeight     =   315
      ScaleWidth      =   13380
      TabIndex        =   6
      Top             =   0
      Width           =   13440
      Begin VB.Label LRenseignementsFenetre 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00008000&
         BackStyle       =   0  'Transparent
         Caption         =   "CUVE GEREE"
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
         TabIndex        =   7
         Top             =   60
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
      Height          =   13575
      Index           =   0
      Left            =   0
      ScaleHeight     =   13575
      ScaleWidth      =   13440
      TabIndex        =   8
      Top             =   375
      Width           =   13440
      Begin VB.PictureBox PBDeplacementFenetre 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   13575
         Index           =   1
         Left            =   0
         ScaleHeight     =   13575
         ScaleWidth      =   28545
         TabIndex        =   9
         Top             =   -240
         Width           =   28545
         Begin VB.PictureBox PBModesProgrammation 
            Height          =   12855
            Left            =   20580
            ScaleHeight     =   12795
            ScaleWidth      =   5295
            TabIndex        =   11
            Top             =   240
            Width           =   5355
            Begin VB.PictureBox PBTousCyclesPompe 
               Height          =   1875
               Left            =   300
               ScaleHeight     =   1815
               ScaleWidth      =   4155
               TabIndex        =   53
               Top             =   4020
               Width           =   4215
               Begin VB.PictureBox PBCyclesPompe 
                  Appearance      =   0  'Flat
                  BackColor       =   &H00FFFFFF&
                  ForeColor       =   &H00FFFFFF&
                  Height          =   435
                  Index           =   0
                  Left            =   360
                  ScaleHeight     =   405
                  ScaleWidth      =   525
                  TabIndex        =   58
                  TabStop         =   0   'False
                  Top             =   300
                  Width           =   555
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
                  TabIndex        =   57
                  TabStop         =   0   'False
                  Top             =   1020
                  Width           =   555
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
                  TabIndex        =   56
                  Top             =   300
                  Width           =   2655
               End
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
                  TabIndex        =   55
                  Top             =   1020
                  Width           =   2655
               End
            End
            Begin VB.PictureBox PBTousModesChauffage 
               Height          =   2655
               Left            =   360
               ScaleHeight     =   2595
               ScaleWidth      =   4155
               TabIndex        =   45
               Top             =   840
               Width           =   4215
               Begin VB.PictureBox PBModesChauffage 
                  Appearance      =   0  'Flat
                  BackColor       =   &H00FFFFFF&
                  ForeColor       =   &H00FFFFFF&
                  Height          =   435
                  Index           =   0
                  Left            =   360
                  ScaleHeight     =   405
                  ScaleWidth      =   585
                  TabIndex        =   52
                  TabStop         =   0   'False
                  Top             =   360
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
                  TabIndex        =   51
                  Top             =   360
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
                  TabIndex        =   50
                  TabStop         =   0   'False
                  Top             =   1080
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
                  TabIndex        =   49
                  Top             =   1080
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
                  TabIndex        =   48
                  TabStop         =   0   'False
                  Top             =   1800
                  Width           =   615
               End
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
                  TabIndex        =   47
                  Top             =   1800
                  Width           =   2655
               End
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
               Left            =   300
               TabIndex        =   54
               Top             =   3660
               Width           =   4215
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
               Left            =   300
               TabIndex        =   46
               Top             =   300
               Width           =   4215
            End
         End
         Begin VB.PictureBox PBProgrammation 
            Height          =   12855
            Left            =   0
            ScaleHeight     =   12795
            ScaleWidth      =   20535
            TabIndex        =   10
            Top             =   240
            Width           =   20595
            Begin VB.PictureBox PBSelectionNomsCuves 
               Height          =   11775
               Left            =   300
               ScaleHeight     =   11715
               ScaleWidth      =   9135
               TabIndex        =   59
               Top             =   660
               Width           =   9195
               Begin VB.OptionButton OBNomsCuves 
                  BackColor       =   &H00C00000&
                  Caption         =   "Index 18"
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
                  Height          =   495
                  Index           =   18
                  Left            =   360
                  Style           =   1  'Graphical
                  TabIndex        =   80
                  Top             =   10440
                  Visible         =   0   'False
                  Width           =   8115
               End
               Begin VB.OptionButton OBNomsCuves 
                  BackColor       =   &H00C00000&
                  Caption         =   "Index 11"
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
                  Height          =   495
                  Index           =   11
                  Left            =   360
                  Style           =   1  'Graphical
                  TabIndex        =   79
                  Top             =   6180
                  Width           =   8115
               End
               Begin VB.OptionButton OBNomsCuves 
                  BackColor       =   &H00C00000&
                  Caption         =   "Index 12"
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
                  Height          =   495
                  Index           =   12
                  Left            =   360
                  Style           =   1  'Graphical
                  TabIndex        =   78
                  Top             =   6720
                  Visible         =   0   'False
                  Width           =   8115
               End
               Begin VB.OptionButton OBNomsCuves 
                  BackColor       =   &H00C00000&
                  Caption         =   "Index 13"
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
                  Height          =   495
                  Index           =   13
                  Left            =   360
                  Style           =   1  'Graphical
                  TabIndex        =   77
                  Top             =   7260
                  Visible         =   0   'False
                  Width           =   8115
               End
               Begin VB.OptionButton OBNomsCuves 
                  BackColor       =   &H00C00000&
                  Caption         =   "Index 14"
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
                  Height          =   495
                  Index           =   14
                  Left            =   360
                  Style           =   1  'Graphical
                  TabIndex        =   76
                  Top             =   7800
                  Visible         =   0   'False
                  Width           =   8115
               End
               Begin VB.OptionButton OBNomsCuves 
                  BackColor       =   &H00C00000&
                  Caption         =   "Index 15"
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
                  Height          =   495
                  Index           =   15
                  Left            =   360
                  Style           =   1  'Graphical
                  TabIndex        =   75
                  Top             =   8340
                  Visible         =   0   'False
                  Width           =   8115
               End
               Begin VB.OptionButton OBNomsCuves 
                  BackColor       =   &H00C00000&
                  Caption         =   "Index 16"
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
                  Height          =   495
                  Index           =   16
                  Left            =   360
                  Style           =   1  'Graphical
                  TabIndex        =   74
                  Top             =   9360
                  Visible         =   0   'False
                  Width           =   8115
               End
               Begin VB.OptionButton OBNomsCuves 
                  BackColor       =   &H00C00000&
                  Caption         =   "Index 17"
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
                  Height          =   495
                  Index           =   17
                  Left            =   360
                  Style           =   1  'Graphical
                  TabIndex        =   73
                  Top             =   9900
                  Visible         =   0   'False
                  Width           =   8115
               End
               Begin VB.OptionButton OBNomsCuves 
                  BackColor       =   &H00FFFFFF&
                  Caption         =   "Index 4"
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
                  Left            =   360
                  Style           =   1  'Graphical
                  TabIndex        =   71
                  Top             =   2040
                  Width           =   8115
               End
               Begin VB.OptionButton OBNomsCuves 
                  BackColor       =   &H00FFFFFF&
                  Caption         =   "Index 5"
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
                  Left            =   360
                  Style           =   1  'Graphical
                  TabIndex        =   70
                  Top             =   2580
                  Width           =   8115
               End
               Begin VB.OptionButton OBNomsCuves 
                  BackColor       =   &H00FFFFFF&
                  Caption         =   "Index 6"
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
                  Left            =   360
                  Style           =   1  'Graphical
                  TabIndex        =   69
                  Top             =   3120
                  Width           =   8115
               End
               Begin VB.OptionButton OBNomsCuves 
                  BackColor       =   &H00008000&
                  Caption         =   "Index 7"
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
                  Height          =   495
                  Index           =   7
                  Left            =   360
                  Style           =   1  'Graphical
                  TabIndex        =   68
                  Top             =   3720
                  Width           =   8115
               End
               Begin VB.OptionButton OBNomsCuves 
                  BackColor       =   &H00008000&
                  Caption         =   "Index 8"
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
                  Height          =   500
                  Index           =   8
                  Left            =   360
                  Style           =   1  'Graphical
                  TabIndex        =   67
                  Top             =   4220
                  Width           =   8115
               End
               Begin VB.OptionButton OBNomsCuves 
                  BackColor       =   &H00008000&
                  Caption         =   "Index 9"
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
                  Height          =   495
                  Index           =   9
                  Left            =   360
                  Style           =   1  'Graphical
                  TabIndex        =   66
                  Top             =   4800
                  Width           =   8115
               End
               Begin VB.OptionButton OBNomsCuves 
                  BackColor       =   &H00008000&
                  Caption         =   "Index 10"
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
                  Height          =   495
                  Index           =   10
                  Left            =   360
                  Style           =   1  'Graphical
                  TabIndex        =   65
                  Top             =   5400
                  Width           =   8115
               End
               Begin VB.OptionButton OBNomsCuves 
                  Appearance      =   0  'Flat
                  BackColor       =   &H00FFFFFF&
                  Caption         =   "Index 1"
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
                  Left            =   360
                  MaskColor       =   &H00FF00FF&
                  Style           =   1  'Graphical
                  TabIndex        =   64
                  Top             =   360
                  UseMaskColor    =   -1  'True
                  Width           =   8115
               End
               Begin VB.OptionButton OBNomsCuves 
                  Appearance      =   0  'Flat
                  BackColor       =   &H00FFFFFF&
                  Caption         =   "Index 2"
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
                  Left            =   360
                  Style           =   1  'Graphical
                  TabIndex        =   63
                  Top             =   840
                  Width           =   8115
               End
               Begin VB.OptionButton OBNomsCuves 
                  BackColor       =   &H00FFFFFF&
                  Caption         =   "Index 3"
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
                  Left            =   360
                  Style           =   1  'Graphical
                  TabIndex        =   62
                  Top             =   1440
                  Width           =   8115
               End
               Begin VB.OptionButton OBNomsCuves 
                  BackColor       =   &H00FFFFFF&
                  Caption         =   "Index 19"
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
                  Index           =   19
                  Left            =   360
                  Style           =   1  'Graphical
                  TabIndex        =   61
                  Top             =   10920
                  Visible         =   0   'False
                  Width           =   8115
               End
               Begin VB.VScrollBar VSChangementCuve 
                  Height          =   7455
                  Left            =   8820
                  Max             =   18
                  Min             =   1
                  TabIndex        =   60
                  Top             =   0
                  Value           =   1
                  Width           =   315
               End
            End
            Begin VB.PictureBox PBRegulationAlarme 
               Height          =   6255
               Left            =   9780
               ScaleHeight     =   6195
               ScaleWidth      =   10395
               TabIndex        =   29
               Top             =   6180
               Width           =   10455
               Begin VB.TextBox TBRegulation 
                  Alignment       =   1  'Right Justify
                  Appearance      =   0  'Flat
                  BackColor       =   &H00C0FFFF&
                  BorderStyle     =   0  'None
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
                  Left            =   5160
                  MaxLength       =   4
                  MultiLine       =   -1  'True
                  TabIndex        =   0
                  Top             =   660
                  Width           =   1575
               End
               Begin VB.TextBox TBRegulation 
                  Alignment       =   1  'Right Justify
                  Appearance      =   0  'Flat
                  BackColor       =   &H00C0FFFF&
                  BorderStyle     =   0  'None
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
                  Index           =   1
                  Left            =   5160
                  MaxLength       =   4
                  MultiLine       =   -1  'True
                  TabIndex        =   1
                  Top             =   1020
                  Width           =   1575
               End
               Begin VB.TextBox TBRegulation 
                  Alignment       =   1  'Right Justify
                  Appearance      =   0  'Flat
                  BackColor       =   &H00C0FFFF&
                  BorderStyle     =   0  'None
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
                  Index           =   2
                  Left            =   5160
                  MaxLength       =   4
                  MultiLine       =   -1  'True
                  TabIndex        =   2
                  Top             =   1980
                  Width           =   1575
               End
               Begin VB.TextBox TBRegulation 
                  Alignment       =   1  'Right Justify
                  Appearance      =   0  'Flat
                  BackColor       =   &H00C0FFFF&
                  BorderStyle     =   0  'None
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
                  Index           =   3
                  Left            =   5160
                  MaxLength       =   4
                  MultiLine       =   -1  'True
                  TabIndex        =   3
                  Top             =   2340
                  Width           =   1575
               End
               Begin VB.TextBox TBRegulation 
                  Alignment       =   1  'Right Justify
                  Appearance      =   0  'Flat
                  BackColor       =   &H00C0FFFF&
                  BorderStyle     =   0  'None
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
                  Index           =   4
                  Left            =   5160
                  MaxLength       =   4
                  MultiLine       =   -1  'True
                  TabIndex        =   4
                  Top             =   3300
                  Width           =   1575
               End
               Begin VB.TextBox TBRegulation 
                  Alignment       =   1  'Right Justify
                  Appearance      =   0  'Flat
                  BackColor       =   &H00C0FFFF&
                  BorderStyle     =   0  'None
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
                  Index           =   5
                  Left            =   5160
                  MaxLength       =   4
                  MultiLine       =   -1  'True
                  TabIndex        =   5
                  Top             =   3660
                  Width           =   1575
               End
               Begin VB.Label LTitresZones 
                  Alignment       =   2  'Center
                  Appearance      =   0  'Flat
                  BackColor       =   &H00FF0000&
                  BorderStyle     =   1  'Fixed Single
                  Caption         =   "Alarme du chauffage"
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   9.75
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00FFFFFF&
                  Height          =   315
                  Index           =   1
                  Left            =   240
                  TabIndex        =   44
                  Top             =   2940
                  Width           =   7035
               End
               Begin VB.Label LTitresZones 
                  Alignment       =   2  'Center
                  Appearance      =   0  'Flat
                  BackColor       =   &H00FF0000&
                  BorderStyle     =   1  'Fixed Single
                  Caption         =   "Régulation du chauffage"
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   9.75
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00FFFFFF&
                  Height          =   315
                  Index           =   0
                  Left            =   240
                  TabIndex        =   43
                  Top             =   1620
                  Width           =   7035
               End
               Begin VB.Label LTitresZones 
                  Alignment       =   2  'Center
                  Appearance      =   0  'Flat
                  BackColor       =   &H00FF0000&
                  BorderStyle     =   1  'Fixed Single
                  Caption         =   "Températures"
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   9.75
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00FFFFFF&
                  Height          =   315
                  Index           =   100
                  Left            =   240
                  TabIndex        =   42
                  Top             =   300
                  Width           =   7035
               End
               Begin VB.Label LLibelle 
                  Appearance      =   0  'Flat
                  BackColor       =   &H00FFFFFF&
                  Caption         =   " Température de VEILLE"
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   9.75
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00000000&
                  Height          =   255
                  Index           =   2
                  Left            =   300
                  TabIndex        =   41
                  Top             =   660
                  Width           =   3795
               End
               Begin VB.Label LLibelle 
                  Appearance      =   0  'Flat
                  BackColor       =   &H00FFFFFF&
                  Caption         =   " Température de PRODUCTION"
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   9.75
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00000000&
                  Height          =   255
                  Index           =   3
                  Left            =   300
                  TabIndex        =   40
                  Top             =   1020
                  Width           =   3795
               End
               Begin VB.Label LLibelle 
                  Appearance      =   0  'Flat
                  BackColor       =   &H00FFFFFF&
                  Caption         =   " Ecart supérieur de régulation"
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   9.75
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00000000&
                  Height          =   255
                  Index           =   6
                  Left            =   300
                  TabIndex        =   39
                  Top             =   2340
                  Width           =   3795
               End
               Begin VB.Label LLibelle 
                  Appearance      =   0  'Flat
                  BackColor       =   &H00FFFFFF&
                  Caption         =   " Ecart supérieur d'alarme"
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   9.75
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00000000&
                  Height          =   255
                  Index           =   9
                  Left            =   300
                  TabIndex        =   38
                  Top             =   3660
                  Width           =   3795
               End
               Begin VB.Label LLibelle 
                  Appearance      =   0  'Flat
                  BackColor       =   &H00FFFFFF&
                  Caption         =   " Ecart inférieur de régulation"
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   9.75
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00000000&
                  Height          =   255
                  Index           =   5
                  Left            =   300
                  TabIndex        =   37
                  Top             =   1980
                  Width           =   3795
               End
               Begin VB.Label LLibelle 
                  Appearance      =   0  'Flat
                  BackColor       =   &H00FFFFFF&
                  Caption         =   " Ecart inférieur d'alarme"
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   9.75
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00000000&
                  Height          =   255
                  Index           =   8
                  Left            =   300
                  TabIndex        =   36
                  Top             =   3300
                  Width           =   3795
               End
               Begin VB.Label LUniteMesure 
                  Alignment       =   2  'Center
                  Appearance      =   0  'Flat
                  BackColor       =   &H00C0FFFF&
                  Caption         =   "°C"
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   9.75
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H80000008&
                  Height          =   255
                  Index           =   0
                  Left            =   6840
                  TabIndex        =   35
                  Top             =   660
                  Width           =   315
               End
               Begin VB.Label LUniteMesure 
                  Alignment       =   2  'Center
                  Appearance      =   0  'Flat
                  BackColor       =   &H00C0FFFF&
                  Caption         =   "°C"
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   9.75
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H80000008&
                  Height          =   255
                  Index           =   1
                  Left            =   6840
                  TabIndex        =   34
                  Top             =   1020
                  Width           =   315
               End
               Begin VB.Label LUniteMesure 
                  Alignment       =   2  'Center
                  Appearance      =   0  'Flat
                  BackColor       =   &H00C0FFFF&
                  Caption         =   "°C"
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   9.75
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H80000008&
                  Height          =   255
                  Index           =   2
                  Left            =   6840
                  TabIndex        =   33
                  Top             =   1980
                  Width           =   315
               End
               Begin VB.Label LUniteMesure 
                  Alignment       =   2  'Center
                  Appearance      =   0  'Flat
                  BackColor       =   &H00C0FFFF&
                  Caption         =   "°C"
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   9.75
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H80000008&
                  Height          =   255
                  Index           =   3
                  Left            =   6840
                  TabIndex        =   32
                  Top             =   2340
                  Width           =   315
               End
               Begin VB.Label LUniteMesure 
                  Alignment       =   2  'Center
                  Appearance      =   0  'Flat
                  BackColor       =   &H00C0FFFF&
                  Caption         =   "°C"
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   9.75
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H80000008&
                  Height          =   255
                  Index           =   4
                  Left            =   6840
                  TabIndex        =   31
                  Top             =   3300
                  Width           =   315
               End
               Begin VB.Label LUniteMesure 
                  Alignment       =   2  'Center
                  Appearance      =   0  'Flat
                  BackColor       =   &H00C0FFFF&
                  Caption         =   "°C"
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   9.75
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H80000008&
                  Height          =   255
                  Index           =   5
                  Left            =   6840
                  TabIndex        =   30
                  Top             =   3660
                  Width           =   315
               End
               Begin VB.Shape SCadresLibellesEtats 
                  BackColor       =   &H00FFFFFF&
                  BackStyle       =   1  'Opaque
                  Height          =   375
                  Index           =   1
                  Left            =   240
                  Top             =   600
                  Width           =   4815
               End
               Begin VB.Shape SCadresLibellesEtats 
                  BackColor       =   &H00FFFFFF&
                  BackStyle       =   1  'Opaque
                  Height          =   375
                  Index           =   0
                  Left            =   240
                  Top             =   960
                  Width           =   4815
               End
               Begin VB.Shape SCadresLibellesEtats 
                  BackColor       =   &H00FFFFFF&
                  BackStyle       =   1  'Opaque
                  Height          =   375
                  Index           =   2
                  Left            =   240
                  Top             =   1920
                  Width           =   4815
               End
               Begin VB.Shape SCadresLibellesEtats 
                  BackColor       =   &H00FFFFFF&
                  BackStyle       =   1  'Opaque
                  Height          =   375
                  Index           =   3
                  Left            =   240
                  Top             =   2280
                  Width           =   4815
               End
               Begin VB.Shape SCadresLibellesEtats 
                  BackColor       =   &H00FFFFFF&
                  BackStyle       =   1  'Opaque
                  Height          =   375
                  Index           =   4
                  Left            =   240
                  Top             =   3240
                  Width           =   4815
               End
               Begin VB.Shape SCadresLibellesEtats 
                  BackColor       =   &H00FFFFFF&
                  BackStyle       =   1  'Opaque
                  Height          =   375
                  Index           =   5
                  Left            =   240
                  Top             =   3600
                  Width           =   4815
               End
               Begin VB.Shape SDecoration 
                  BackColor       =   &H00C0FFFF&
                  BackStyle       =   1  'Opaque
                  Height          =   375
                  Index           =   0
                  Left            =   5040
                  Top             =   600
                  Width           =   2235
               End
               Begin VB.Shape SDecoration 
                  BackColor       =   &H00C0FFFF&
                  BackStyle       =   1  'Opaque
                  Height          =   375
                  Index           =   1
                  Left            =   5040
                  Top             =   960
                  Width           =   2235
               End
               Begin VB.Shape SDecoration 
                  BackColor       =   &H00C0FFFF&
                  BackStyle       =   1  'Opaque
                  Height          =   375
                  Index           =   2
                  Left            =   5040
                  Top             =   1920
                  Width           =   2235
               End
               Begin VB.Shape SDecoration 
                  BackColor       =   &H00C0FFFF&
                  BackStyle       =   1  'Opaque
                  Height          =   375
                  Index           =   3
                  Left            =   5040
                  Top             =   2280
                  Width           =   2235
               End
               Begin VB.Shape SDecoration 
                  BackColor       =   &H00C0FFFF&
                  BackStyle       =   1  'Opaque
                  Height          =   375
                  Index           =   4
                  Left            =   5040
                  Top             =   3240
                  Width           =   2235
               End
               Begin VB.Shape SDecoration 
                  BackColor       =   &H00C0FFFF&
                  BackStyle       =   1  'Opaque
                  Height          =   375
                  Index           =   5
                  Left            =   5040
                  Top             =   3600
                  Width           =   2235
               End
            End
            Begin VB.PictureBox PBJourneesTypes 
               Height          =   4875
               Left            =   9780
               ScaleHeight     =   4815
               ScaleWidth      =   10395
               TabIndex        =   20
               Top             =   660
               Width           =   10455
               Begin VB.PictureBox PBEchelle24H 
                  Appearance      =   0  'Flat
                  AutoSize        =   -1  'True
                  BackColor       =   &H00FFFFFF&
                  BorderStyle     =   0  'None
                  ForeColor       =   &H00FFFFFF&
                  Height          =   945
                  Index           =   3
                  Left            =   240
                  Picture         =   "FGestionRegulation.frx":277EC
                  ScaleHeight     =   63
                  ScaleMode       =   3  'Pixel
                  ScaleWidth      =   661
                  TabIndex        =   23
                  TabStop         =   0   'False
                  Top             =   3600
                  Width           =   9915
               End
               Begin VB.PictureBox PBEchelle24H 
                  Appearance      =   0  'Flat
                  AutoSize        =   -1  'True
                  BackColor       =   &H00FFFFFF&
                  BorderStyle     =   0  'None
                  ForeColor       =   &H00FFFFFF&
                  Height          =   945
                  Index           =   2
                  Left            =   240
                  Picture         =   "FGestionRegulation.frx":4606E
                  ScaleHeight     =   63
                  ScaleMode       =   3  'Pixel
                  ScaleWidth      =   661
                  TabIndex        =   22
                  TabStop         =   0   'False
                  Top             =   2100
                  Width           =   9915
               End
               Begin VB.PictureBox PBEchelle24H 
                  Appearance      =   0  'Flat
                  AutoSize        =   -1  'True
                  BackColor       =   &H00FFFFFF&
                  BorderStyle     =   0  'None
                  ForeColor       =   &H00FFFFFF&
                  Height          =   945
                  Index           =   1
                  Left            =   240
                  Picture         =   "FGestionRegulation.frx":648F0
                  ScaleHeight     =   63
                  ScaleMode       =   3  'Pixel
                  ScaleWidth      =   661
                  TabIndex        =   21
                  TabStop         =   0   'False
                  Top             =   600
                  Width           =   9915
               End
               Begin VB.Label LTitresZones 
                  Alignment       =   2  'Center
                  Appearance      =   0  'Flat
                  BackColor       =   &H00FF0000&
                  BorderStyle     =   1  'Fixed Single
                  Caption         =   "Journée type de TRAVAIL"
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   9.75
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00FFFFFF&
                  Height          =   315
                  Index           =   4
                  Left            =   240
                  TabIndex        =   26
                  Top             =   300
                  Width           =   9915
               End
               Begin VB.Label LTitresZones 
                  Alignment       =   2  'Center
                  Appearance      =   0  'Flat
                  BackColor       =   &H00FF0000&
                  BorderStyle     =   1  'Fixed Single
                  Caption         =   "Journée type de VEILLE"
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   9.75
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00FFFFFF&
                  Height          =   315
                  Index           =   3
                  Left            =   240
                  TabIndex        =   25
                  Top             =   1800
                  Width           =   9915
               End
               Begin VB.Label LTitresZones 
                  Alignment       =   2  'Center
                  Appearance      =   0  'Flat
                  BackColor       =   &H00FF0000&
                  BorderStyle     =   1  'Fixed Single
                  Caption         =   "Journée type de REPRISE"
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   9.75
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00FFFFFF&
                  Height          =   315
                  Index           =   2
                  Left            =   240
                  TabIndex        =   24
                  Top             =   3300
                  Width           =   9915
               End
            End
            Begin VB.Label LLibelles 
               Alignment       =   2  'Center
               BackColor       =   &H00800000&
               BorderStyle     =   1  'Fixed Single
               Caption         =   "CHOIX DE LA CUVE"
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
               TabIndex        =   72
               Top             =   300
               Width           =   9195
            End
            Begin VB.Label LLibelles 
               Alignment       =   2  'Center
               BackColor       =   &H00800000&
               BorderStyle     =   1  'Fixed Single
               Caption         =   "REGULATION ET ALARME"
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
               Left            =   9780
               TabIndex        =   28
               Top             =   5820
               Width           =   10455
            End
            Begin VB.Label LLibelles 
               Alignment       =   2  'Center
               BackColor       =   &H00800000&
               BorderStyle     =   1  'Fixed Single
               Caption         =   "JOURNEES TYPES"
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
               Left            =   9780
               TabIndex        =   27
               Top             =   300
               Width           =   10455
            End
         End
      End
   End
End
Attribute VB_Name = "FGestionRegulation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Rôle                    : Fenêtre gérant la gestion de la régulation d'un cuve
' Nom                    : FGestionRegulation.frm
' Date de création : 11/05/2011
' Détails                :
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

'--- déclarations obligatoires ---
Option Explicit

'--- options générales ---
Option Base 1
DefVar A-Z

'--- constantes privées ---
Private Const X_MINI_POMPE As Single = 2
Private Const X_MAXI_POMPE As Single = 626                            'encadrement de la zone pompe
Private Const Y_MINI_POMPE As Single = 2                                 '(X mini, Y mini, X maxi, Y Maxi) de la zone
Private Const Y_MAXI_POMPE As Single = 16

Private Const X_MINI_CHAUFFAGE As Single = 2
Private Const X_MAXI_CHAUFFAGE As Single = 626                      'encadrement de la zone chauffage
Private Const Y_MINI_CHAUFFAGE As Single = 46                        '(X mini, Y mini, X maxi, Y Maxi) de la zone
Private Const Y_MAXI_CHAUFFAGE As Single = 60

Private Const LARGEUR_SEGMENT As Single = 13                       'largeur d'un segment correspondant à 1/2 heure


Private Const TITRE_FENETRE As String = "GESTION DE LA REGULATION"
Private Const TITRE_MESSAGES As String = TITRE_FENETRE

'--- énumérations privées ---
Private Enum INDEX_CHAMPS
    IDX_TEMP_VEILLE = 0
    IDX_TEMP_PRODUCTION = 1
    IDX_ECART_INFERIEUR_REGUL = 2
    IDX_ECART_SUPERIEUR_REGUL = 3
    IDX_ECART_INFERIEUR_ALARME = 4
    IDX_ECART_SUPERIEUR_ALARME = 5
End Enum

'--- variables privées ---
Private PremiereActivation As Boolean
Private InterdireEvenements As Boolean      'pour interdire certains évènements
Private CouleurPompe As Long                      'couleur du segment Pompe
Private CouleurChauffage As Long                 'couleur du segment chauffage
Private XDepart As Single                              'point de départ d'une programmation
Private XArrivee As Single                             'point d'arrivée d'une programmation
Private MemX As Single                                 'mémoire des X
Private MemY As Single                                 'mémoire des Y

'--- variables publiques ---
Public NumFenetre As Long                           'numéro de la fenêtre lorsqu'elle devient active
Public CuveEnCours As Integer                      'cuve en cours

'--- tableaux privés ---
Private TAdrAPI_TBRegulation(INDEX_CHAMPS.IDX_TEMP_VEILLE To INDEX_CHAMPS.IDX_ECART_SUPERIEUR_ALARME) As String

Private Sub CBAgrandirFENETRE_Click()
    On Error Resume Next
    Me.WindowState = vbMaximized
End Sub

Private Sub CBAnnuler_Click()
    
    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next

    '--- restitue les dernières manipulations et valeurs sur la fenêtre ---
    LectureValeursFenetre

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
    
    '--- enregistre les valeurs dans l'automate ---
    If EnregistreVersAutomate = True Then
        
        '--- enregistrement des valeurs de la fenêtre ---
        EnregistreValeursfenetre
    
        '--- sauvegardes ---
        SauveJourneesTypes
        SauveRegulation
    
        '--- sauvegarde de la configuration ---
        'SauveConfiguration
    
        '--- affectation ---
        BloqueChangementCuve False
        CBValider.Enabled = False
        CBAnnuler.Enabled = False
    
    Else
        
        '--- lancer un message d'erreur ---
        Bidon = MessageErreur(TITRE_MESSAGES, MESSAGE_500)
        
        '--- chargement des journées types et de la régulation ---
        ChargeJourneesTypes
        ChargeRegulation
        
        '--- restitue les dernières manipulations et valeurs sur la fenêtre ---
        LectureValeursFenetre
    
    End If
    
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
        PremiereActivation = True
        Me.Refresh
    End If

End Sub

'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Rôle      : Bloque ou non le changement de cuve
' Entrées : Blocage -> FALSE = Changement de cuve autorisé
'                                    TRUE = Changement de cuve interdit
' Retours :
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Private Sub BloqueChangementCuve(ByVal Blocage As Boolean)
    
    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next
    
    '--- déclaration ---
    Dim a As Integer
    
    '--- blocage des boutons de sélection ---
    'ATTENTION partir du plus éléver vers le plus bas pour éviter les problèmes de rebonds d'évènements
    For a = OBNomsCuves.UBound To OBNomsCuves.LBound Step -1
        OBNomsCuves(a).Enabled = Not (Blocage)
    Next a
    
    '--- blocage de la barre de scrolling ---
    VSChangementCuve.Enabled = Not (Blocage)

End Sub

'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Rôle      : Enregistre les valeurs dans l'automate
' Entrées :
' Retours :
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Private Function EnregistreVersAutomate() As Boolean
    
    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next

    '--- déclaration ---
    Dim ValeurRetourneeAPI As Long                  'valeur retournée par une fonction concernant le dialogue avec l'automate
    
    '--- affectation par défaut ---
    EnregistreVersAutomate = True
    
    '--- blocage du timer ---
    TimerEtatsRegulation.Enabled = False
    
    Dim NumCuve As Integer
    NumCuve = TEtatsCuves(CuveEnCours).IndexAutomate
    
    '--- calculs des adresses ---
    TAdrAPI_TBRegulation(INDEX_CHAMPS.IDX_TEMP_VEILLE) = "TempVeillePCCuve" & Right("00" & NumCuve, 2)
    'Groupe TEMP_VEILLE_CUVES
    
    TAdrAPI_TBRegulation(INDEX_CHAMPS.IDX_TEMP_PRODUCTION) = "TempProductionPCCuve" & Right("00" & NumCuve, 2)
    
    'Call Log("APIEcritureVariableNommee('TEMP_PRODUCTION_CUVES',TempProductionPCCuve" & Right("00" & NumCuve, 2) & "," & CStr(Fix(CSng(TBRegulation(INDEX_CHAMPS.IDX_TEMP_PRODUCTION).Text) * 10)) & ")")
    'Groupe TEMP_PRODUCTION_CUVES
    
    TAdrAPI_TBRegulation(INDEX_CHAMPS.IDX_ECART_INFERIEUR_REGUL) = "EcartInfRegulPCCuve" & Right("00" & NumCuve, 2)
    'Groupe ECARTS_INF_REGUL_CUVES
    
    TAdrAPI_TBRegulation(INDEX_CHAMPS.IDX_ECART_SUPERIEUR_REGUL) = "EcartSupRegulPCCuve" & Right("00" & NumCuve, 2)
    'Groupe ECARTS_SUP_REGUL_CUVES
    
    TAdrAPI_TBRegulation(INDEX_CHAMPS.IDX_ECART_INFERIEUR_ALARME) = "EcartInfAlarmePCCuve" & Right("00" & NumCuve, 2)
    'Groupe ECARTS_INF_ALARME_CUVES
    
    TAdrAPI_TBRegulation(INDEX_CHAMPS.IDX_ECART_SUPERIEUR_ALARME) = "EcartSupAlarmePCCuve" & Right("00" & NumCuve, 2)
    'Groupe ECARTS_SUP_ALARME_CUVES

    If PROGRAMME_AVEC_AUTOMATE = True Then
        
        '--- températures de veille ---
        ValeurRetourneeAPI = APIEcritureVariableNommee("TEMP_VEILLE_CUVES", _
                                                                                           TAdrAPI_TBRegulation(INDEX_CHAMPS.IDX_TEMP_VEILLE), _
                                                                                           CStr(Fix(CSng(TBRegulation(INDEX_CHAMPS.IDX_TEMP_VEILLE).Text) * 10)))
        If ValeurRetourneeAPI <> 0 Then
            EnregistreVersAutomate = False
        End If
        
        '--- températures de production ---
        ValeurRetourneeAPI = APIEcritureVariableNommee("TEMP_PRODUCTION_CUVES", _
                                                                                           TAdrAPI_TBRegulation(INDEX_CHAMPS.IDX_TEMP_PRODUCTION), _
                                                                                           CStr(Fix(CSng(TBRegulation(INDEX_CHAMPS.IDX_TEMP_PRODUCTION).Text) * 10)))
        If ValeurRetourneeAPI <> 0 Then
            EnregistreVersAutomate = False
        End If
            
        '--- écart infèrieur de régulation ---
        ValeurRetourneeAPI = APIEcritureVariableNommee("ECARTS_INF_REGUL_CUVES", _
                                                                                           TAdrAPI_TBRegulation(INDEX_CHAMPS.IDX_ECART_INFERIEUR_REGUL), _
                                                                                           CStr(Fix(CSng(TBRegulation(INDEX_CHAMPS.IDX_ECART_INFERIEUR_REGUL).Text) * 10)))
        If ValeurRetourneeAPI <> 0 Then
            EnregistreVersAutomate = False
        End If
        
        '--- écart supèrieur de régulation ---
        ValeurRetourneeAPI = APIEcritureVariableNommee("ECARTS_SUP_REGUL_CUVES", _
                                                                                           TAdrAPI_TBRegulation(INDEX_CHAMPS.IDX_ECART_SUPERIEUR_REGUL), _
                                                                                           CStr(Fix(CSng(TBRegulation(INDEX_CHAMPS.IDX_ECART_SUPERIEUR_REGUL).Text) * 10)))
        If ValeurRetourneeAPI <> 0 Then
            EnregistreVersAutomate = False
        End If

        '--- écart infèrieur d'alarme ---
        ValeurRetourneeAPI = APIEcritureVariableNommee("ECARTS_INF_ALARME_CUVES", _
                                                                                           TAdrAPI_TBRegulation(INDEX_CHAMPS.IDX_ECART_INFERIEUR_ALARME), _
                                                                                           CStr(Fix(CSng(TBRegulation(INDEX_CHAMPS.IDX_ECART_INFERIEUR_ALARME).Text) * 10)))
        If ValeurRetourneeAPI <> 0 Then
            EnregistreVersAutomate = False
        End If
        
        '--- écart supèrieur d'alarme ---
        ValeurRetourneeAPI = APIEcritureVariableNommee("ECARTS_SUP_ALARME_CUVES", _
                                                                                           TAdrAPI_TBRegulation(INDEX_CHAMPS.IDX_ECART_SUPERIEUR_ALARME), _
                                                                                           CStr(Fix(CSng(TBRegulation(INDEX_CHAMPS.IDX_ECART_SUPERIEUR_ALARME).Text) * 10)))
        If ValeurRetourneeAPI <> 0 Then
            EnregistreVersAutomate = False
        End If
                                                       
    End If
    
    '--- déblocage du timer ---
    TimerEtatsRegulation.Enabled = True
    
End Function

'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Rôle      : Visualisation des différents états de la régulation
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Private Sub EtatsRegulation()
    
    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next

    '--- contrôle automatique des changements ---
    With TEtatsCuves(CuveEnCours)
        If .API_Changements = True Then

            '--- lecture de la régulation ---
            InterdireEvenements = True
            LectureRegulation
            InterdireEvenements = False

            '--- affectation ---
            .API_Changements = False

        End If
    End With
        
End Sub

'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Rôle      : Enregistre les valeurs des journées types
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Private Sub EnregistreJourneesTypes()
    
    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next

    '--- déclaration ---
    Dim a As Integer, _
            b As Integer, _
            c As Integer, _
            PointeurPompe As Integer, _
            PointeurChauffage As Integer
    Dim CouleurPointPompe As Long, _
           CouleurPointChauffage As Long, _
           MemCouleurPointPompe As Long, _
           MemCouleurPointChauffage As Long
    Dim xPoint As Single
    Dim TopDebutCycle As String * 14, _
            TopFinCycle As String * 14

    '--- analyse de toutes les échelles graphiques ---
    For a = JOURNEES_TYPES.J_TRAVAIL To JOURNEES_TYPES.J_REPRISE
        
        '--- affectation ---
        PointeurPompe = 0
        MemCouleurPointPompe = -1                            'forcer un mode inexistant pour la comparaison
        PointeurChauffage = 0
        MemCouleurPointChauffage = -1                       'forcer un mode inexistant pour la comparaison
        
        With TJourneesTypes(CuveEnCours, a)
            
            For b = 0 To 23                                               'cycle des heures
                For c = 0 To 30 Step 30                              'cycle des demi-heures
                
                    '--- affectation ---
                    TopDebutCycle = "XXXXXXXX" & Right("0" & CStr(b), 2) & Right("0" & CStr(c), 2) & "00"
                    xPoint = 3 + (2 * LARGEUR_SEGMENT) * b + LARGEUR_SEGMENT * Abs(c = 30)
                    
                    CouleurPointPompe = PBEchelle24H(a).Point(xPoint, Y_MINI_POMPE)
                    CouleurPointChauffage = PBEchelle24H(a).Point(xPoint, Y_MINI_CHAUFFAGE)
                    
                    '--- vérification avec la couleur pour la pompe---
                    If CouleurPointPompe <> MemCouleurPointPompe And _
                       PBTousCyclesPompe.Visible = True Then
                    
                        '--- heures pour la pompe---
                        If PointeurPompe > 0 Then
                            .TTopsFinPompe(PointeurPompe) = TopFinCycle
                        End If
                        Inc PointeurPompe
                        .TTopsDebutPompe(PointeurPompe) = TopDebutCycle
                                                    
                        '--- mode pour la pompe---
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
                        TopFinCycle = "XXXXXXXX" & Right("0" & CStr(b), 2) & "2959"
                    Else
                        TopFinCycle = "XXXXXXXX" & Right("0" & CStr(b), 2) & "5959"
                    End If
                    
                Next c
            Next b
        
            '--- forçage de la dernière heure ---
            If PointeurPompe > 0 And PBTousCyclesPompe.Visible = True Then
                .TTopsFinPompe(PointeurPompe) = "XXXXXXXX235959"
            End If
            If PointeurChauffage > 0 Then
                .TTopsFinChauffage(PointeurChauffage) = "XXXXXXXX235959"
            End If
        
            '--- vidage des champs inutilisés de la pompe---
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
    Next a

End Sub

'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Rôle      : Lecture des valeurs des journées types
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Private Sub LectureJourneesTypes()
    
    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next

    '--- déclaration ---
    Dim a As Integer, _
            b As Integer, _
            HeureDebutTop As Integer, _
            MinuteDebutTop As Integer, _
            HeureFinTop As Integer, _
            MinuteFinTop As Integer
    Dim CouleurRemplissagePompe As Long, _
           CouleurRemplissageChauffage As Long
    Dim XDebut As Single, XFin As Single

    '--- analyse de toutes les échelles graphiques ---
    For a = JOURNEES_TYPES.J_TRAVAIL To JOURNEES_TYPES.J_REPRISE
        
        With TJourneesTypes(CuveEnCours, a)
            
            Select Case CuveEnCours

                Case CUVES_REGULATION.C_DEC, CUVES_REGULATION.C_C31, CUVES_REGULATION.C_C32 ' CUVES_REGULATION.C_SAT,
                    '--- cuves avec pompe ---
                    For b = 1 To NBR_TOPS_POSSIBLES
                        
                        '--- contrôle ---
                        If .TTopsDebutPompe(b) = String(14, 0) Or _
                            Trim(.TTopsDebutPompe(b)) = "" Then
                                Exit For
                        End If
            
                        '--- heure, minute de début et de fin pour la pompe---
                        HeureDebutTop = Val(Mid(.TTopsDebutPompe(b), 9, 2))
                        MinuteDebutTop = Val(Mid(.TTopsDebutPompe(b), 11, 2))
                        HeureFinTop = Val(Mid(.TTopsFinPompe(b), 9, 2))
                        MinuteFinTop = Val(Mid(.TTopsFinPompe(b), 11, 2))
            
                        '--- calcul du point de début et de fin de tracer pour la pompe---
                        XDebut = X_MINI_POMPE + HeureDebutTop * (2 * LARGEUR_SEGMENT)
                        XDebut = XDebut - LARGEUR_SEGMENT * (MinuteDebutTop = 30)
                        XFin = X_MINI_POMPE + HeureFinTop * (2 * LARGEUR_SEGMENT)
                        XFin = XFin - LARGEUR_SEGMENT * (MinuteFinTop = 29) - (2 * LARGEUR_SEGMENT) * (MinuteFinTop = 59)
                
                        '--- couleur de remplissage pour la pompe---
                        Select Case .TCyclesPompe(b)
                            Case CYCLES_POMPES.CP_ARRET: CouleurRemplissagePompe = COULEURS_ECHELLES_GRAPHIQUES.C_ARRET_POMPE
                            Case CYCLES_POMPES.CP_MARCHE: CouleurRemplissagePompe = COULEURS_ECHELLES_GRAPHIQUES.C_MARCHE_POMPE
                            Case Else: CouleurRemplissagePompe = COULEURS_ECHELLES_GRAPHIQUES.C_ARRET_POMPE
                        End Select
                
                        '--- tracer de la pompe---
                        If XFin > X_MAXI_POMPE Then XFin = X_MAXI_POMPE
                        PBEchelle24H(a).AutoRedraw = True
                        PBEchelle24H(a).Line (XDebut, Y_MINI_POMPE)-(XFin, Y_MAXI_POMPE), CouleurRemplissagePompe, BF
                        PBEchelle24H(a).AutoRedraw = False
                
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
                XDebut = (XDebut - LARGEUR_SEGMENT * (MinuteDebutTop = 30))
                XFin = X_MINI_CHAUFFAGE + HeureFinTop * (2 * LARGEUR_SEGMENT)
                XFin = (XFin - LARGEUR_SEGMENT * (MinuteFinTop = 29) - (2 * LARGEUR_SEGMENT) * (MinuteFinTop = 59))
                
                '--- couleur de remplissage pour le chauffage ---
                Select Case .TModesChauffage(b)
                    Case MODES_PRODUCTION.M_ARRET: CouleurRemplissageChauffage = COULEURS_ECHELLES_GRAPHIQUES.C_MODE_ARRET
                    Case MODES_PRODUCTION.M_VEILLE: CouleurRemplissageChauffage = COULEURS_ECHELLES_GRAPHIQUES.C_MODE_VEILLE
                    Case MODES_PRODUCTION.M_PRODUCTION: CouleurRemplissageChauffage = COULEURS_ECHELLES_GRAPHIQUES.C_MODE_PRODUCTION
                    Case Else: CouleurRemplissageChauffage = COULEURS_ECHELLES_GRAPHIQUES.C_MODE_ARRET
                End Select
                
                '--- tracer du chauffage ---
                If XFin > X_MAXI_CHAUFFAGE Then XFin = X_MAXI_CHAUFFAGE
                PBEchelle24H(a).AutoRedraw = True
                PBEchelle24H(a).Line (XDebut, Y_MINI_CHAUFFAGE)-(XFin, Y_MAXI_CHAUFFAGE), CouleurRemplissageChauffage, BF
                PBEchelle24H(a).AutoRedraw = False
            
            Next b
            
        End With
            
    Next a
            
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
    VManipsGestionRegulation.AppareillageConcerne = True
    ChoixAppareillage
End Sub

Private Sub LTitreTousCyclesPompe_Click()
    On Error Resume Next
    VManipsGestionRegulation.AppareillageConcerne = False
    ChoixAppareillage
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
    VManipsGestionRegulation.CyclesPompe = Index
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
    VManipsGestionRegulation.ModesChauffage = Index
    CouleurChauffage = PBModesChauffage(Index).BackColor

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
            AcquittementAlarmes
        
        Case vbKeyEscape
            '--- touche échap ---
            CBQuitter_Click
        
        Case vbKeyPageUp
            '--- saut de page arrière ---
            If VSChangementCuve.value > VSChangementCuve.min Then
                VSChangementCuve.value = Pred(VSChangementCuve.value)
            End If
            KeyCode = 0
        
        Case vbKeyPageDown
            '--- saut de page avant ---
            If VSChangementCuve.value < VSChangementCuve.Max Then
                VSChangementCuve.value = Succ(VSChangementCuve.value)
            End If
            KeyCode = 0
            
        Case Else

    End Select

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

Private Sub OBNomsCuves_Click(Index As Integer)
    
    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next

    '--- déclaration ---
    Dim OCCBoutonOption As OptionButton

    For Each OCCBoutonOption In OBNomsCuves
        With OCCBoutonOption
            If .value = True Then
                .BackColor = COULEURS.ROUGE_3
                .ForeColor = COULEURS.NOIR
            Else
                Select Case .Index
  
                    Case 200
                        '--- préparation ---
                        .BackColor = COULEURS.BLANC
                        .ForeColor = COULEURS.NOIR
                       
                    Case CUVES_REGULATION.C_C00 To CUVES_REGULATION.C_C07
                        '--- préparation ---
                        .BackColor = COULEURS.BLANC
                        .ForeColor = COULEURS.NOIR
                    Case CUVES_REGULATION.C_C13 To CUVES_REGULATION.C_C15 ' SZP C16
                        '--- anodisation ---
                        .BackColor = COULEURS.VERT_5
                        .ForeColor = COULEURS.BLANC
                    Case Else
                        '--- coloration et fin de ligne ---
                        .BackColor = COULEURS.BLEU_5
                        .ForeColor = COULEURS.BLANC
                End Select
            End If
        End With
    Next
    
    '--- changement de l'ascenseur (si click direct dans un des boutons) ---
    If InterdireEvenements = False Then
        VSChangementCuve.value = Index
    End If

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
    Dim SegmentDepart As Integer, _
            ZoneTravail As Integer

    If Button = vbLeftButton Then
        
        '--- recherche de la zone de travail ---
        ZoneTravail = 0
        If VManipsGestionRegulation.AppareillageConcerne = False Then
            If X >= X_MINI_POMPE And X <= X_MAXI_POMPE And Y >= Y_MINI_POMPE And Y <= Y_MAXI_POMPE Then
                ZoneTravail = 1
            End If
        Else
            If X >= X_MINI_CHAUFFAGE And X <= X_MAXI_CHAUFFAGE And Y >= Y_MINI_CHAUFFAGE And Y <= Y_MAXI_CHAUFFAGE Then
                ZoneTravail = 2
            End If
        End If
        
        '--- analyse ---
        If ZoneTravail > 0 Then
        
            '--- calcul du point de départ ---
            SegmentDepart = Int(X / LARGEUR_SEGMENT)
            XDepart = (SegmentDepart * LARGEUR_SEGMENT) + 2
        
            '--- tracer de la zone de départ ---
            Select Case ZoneTravail
                Case 1
                    '--- zone pompe---
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
            BloqueChangementCuve True
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
        If VManipsGestionRegulation.AppareillageConcerne = False Then
            If X >= X_MINI_POMPE And X <= X_MAXI_POMPE And Y >= Y_MINI_POMPE And Y <= Y_MAXI_POMPE Then
                ZoneTravail = 1
            End If
        Else
            If X >= X_MINI_CHAUFFAGE And X <= X_MAXI_CHAUFFAGE And Y >= Y_MINI_CHAUFFAGE And Y <= Y_MAXI_CHAUFFAGE Then
                ZoneTravail = 2
            End If
        End If

        If (X <> MemX Or Y <> MemY) And ZoneTravail > 0 Then

            '--- calcul du point d'arrivée ---
            XArrivee = (Int(X / LARGEUR_SEGMENT) * LARGEUR_SEGMENT) + 2

            '--- tracer de la zone de départ ---
            Select Case ZoneTravail
                Case 1
                    '--- zone pompe---
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

'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Rôle      : Lecture des valeurs de la régulation
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Private Sub LectureRegulation()
    
    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next
    Dim cuve As Integer
    
    '--- affichage général ---
    With TEtatsCuves(CuveEnCours).Temperatures
        AffichageNombre TBRegulation(INDEX_CHAMPS.IDX_TEMP_VEILLE), .TempVeille, FORMAT_TEMPERATURE_1_DECIMALE
        AffichageNombre TBRegulation(INDEX_CHAMPS.IDX_TEMP_PRODUCTION), .TempProduction, FORMAT_TEMPERATURE_1_DECIMALE
        AffichageNombre TBRegulation(INDEX_CHAMPS.IDX_ECART_INFERIEUR_REGUL), .EcartInferieurRegul, FORMAT_TEMPERATURE_1_DECIMALE
        AffichageNombre TBRegulation(INDEX_CHAMPS.IDX_ECART_SUPERIEUR_REGUL), .EcartSuperieurRegul, FORMAT_TEMPERATURE_1_DECIMALE
        AffichageNombre TBRegulation(INDEX_CHAMPS.IDX_ECART_INFERIEUR_ALARME), .EcartInferieurAlarme, FORMAT_TEMPERATURE_1_DECIMALE
        AffichageNombre TBRegulation(INDEX_CHAMPS.IDX_ECART_SUPERIEUR_ALARME), .EcartSuperieurAlarme, FORMAT_TEMPERATURE_1_DECIMALE
    End With

End Sub

'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Rôle      : Enregistre les valeurs de la régulation
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Private Sub EnregistreRegulation()
    
    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next

    '--- affectation général ---
    With TEtatsCuves(CuveEnCours).Temperatures
        .TempVeille = ControleSiNombre(TBRegulation(INDEX_CHAMPS.IDX_TEMP_VEILLE), 0)
        .TempProduction = ControleSiNombre(TBRegulation(INDEX_CHAMPS.IDX_TEMP_PRODUCTION), 0)
        .EcartInferieurRegul = ControleSiNombre(TBRegulation(INDEX_CHAMPS.IDX_ECART_INFERIEUR_REGUL), 0)
        .EcartSuperieurRegul = ControleSiNombre(TBRegulation(INDEX_CHAMPS.IDX_ECART_SUPERIEUR_REGUL), 0)
        .EcartInferieurAlarme = ControleSiNombre(TBRegulation(INDEX_CHAMPS.IDX_ECART_INFERIEUR_ALARME), 0)
        .EcartSuperieurAlarme = ControleSiNombre(TBRegulation(INDEX_CHAMPS.IDX_ECART_SUPERIEUR_ALARME), 0)
    End With

End Sub

'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Rôle      : Restitue les dernières manipulations et valeurs sur la fenêtre
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

    '--- analyse de la présence de la pompe en fonction de la cuve ---
    Select Case CuveEnCours
    
        Case CUVES_REGULATION.C_DEC, CUVES_REGULATION.C_C31, CUVES_REGULATION.C_C32  'SZP CUVES_REGULATION.C_SAT,
            '--- cuves avec pompe ---
            PBTousCyclesPompe.Visible = True
            LTitreTousCyclesPompe.Visible = True
            For a = PBEchelle24H.LBound To PBEchelle24H.UBound
                Set PBEchelle24H(a).Picture = TImgEchelles24H(ECHELLES_24H.E_POMPE_CHAUFFAGE)
            Next a
    
        'Case CUVES_REGULATION.C_C37
            '--- cas de l'étuve ---
        '    PBTousCyclesPompe.Visible = False
        '    LTitreTousCyclesPompe.Visible = False
        '    For a = PBEchelle24H.LBound To PBEchelle24H.UBound
        '        Set PBEchelle24H(a).Picture = TImgEchelles24H(ECHELLES_24H.E_VENTILATION_CHAUFFAGE)
        '   Next a
            
        Case Else
            '--- cuves sans pompe ---
            PBTousCyclesPompe.Visible = False
            LTitreTousCyclesPompe.Visible = False
            For a = PBEchelle24H.LBound To PBEchelle24H.UBound
                Set PBEchelle24H(a).Picture = TImgEchelles24H(ECHELLES_24H.E_CHAUFFAGE)
            Next a
    
    End Select
    
    '--- sélecteurs ---
    With VManipsGestionRegulation
        OBCyclesPompe(.CyclesPompe).value = True
        OBModesChauffage(.ModesChauffage).value = True
    End With

    '--- choix de l'appareillage ---
    ChoixAppareillage

    '--- valeurs de la régulation ---
    LectureRegulation

    '--- lecture des journées types ---
    LectureJourneesTypes

    '--- affichage de la cuve gérée ---
    With TEtatsCuves(CuveEnCours).DefinitionCuve
        LRenseignementsFenetre.Caption = "Cuve " & .NomCuve & " - " & .LibelleCuve
    End With
    
    '--- affectation ---
    BloqueChangementCuve False
    CBValider.Enabled = False
    CBAnnuler.Enabled = False
    
    '--- autoriser les évènements ---
    InterdireEvenements = False

End Sub

Private Sub PBModesChauffage_Click(Index As Integer)
    
    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next
    
    '--- affectation ---
    OBModesChauffage(Index).value = True

End Sub

Private Sub PBCyclesPompe_Click(Index As Integer)
    
    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next
    
    '--- affectation ---
    OBCyclesPompe(Index).value = True

End Sub

Private Sub PBTousModesChauffage_Click()
    On Error Resume Next
    VManipsGestionRegulation.AppareillageConcerne = True
    ChoixAppareillage
End Sub

Private Sub PBTousCyclesPompe_Click()
    On Error Resume Next
    VManipsGestionRegulation.AppareillageConcerne = False
    ChoixAppareillage
End Sub

Private Sub TBRegulation_Change(Index As Integer)
    On Error Resume Next
    If InterdireEvenements = False Then
        BloqueChangementCuve True
        CBValider.Enabled = True
        CBAnnuler.Enabled = True
    End If
End Sub

Private Sub TBRegulation_GotFocus(Index As Integer)
    On Error Resume Next
    With TBRegulation(Index)
        .SelStart = 0          'met en surbrillance la sélection saisie
        .SelLength = Len(.Text)
    End With
End Sub

Private Sub TBRegulation_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    On Error Resume Next
    FiltreToucheFonction KeyCode, Shift
End Sub

Private Sub TBRegulation_KeyPress(Index As Integer, KeyAscii As Integer)
    On Error Resume Next
    FiltreToucheASCII KeyAscii, DONNEES.D_NBR_REELS_POSITIFS, 5
End Sub

'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Rôle      : Enregistre les dernières manipulations et valeurs de la fenêtre
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Private Sub EnregistreValeursfenetre()
    
    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next
    
    '--- enregistrements ---
    EnregistreJourneesTypes
    EnregistreRegulation

End Sub

'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Rôle      : Fixe les couleurs de traçage
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Private Sub CouleursDeTraçage()
    
    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next

    '--- modes de la pompe---
    PBCyclesPompe(CYCLES_POMPES.CP_ARRET).BackColor = COULEURS_ECHELLES_GRAPHIQUES.C_ARRET_POMPE
    PBCyclesPompe(CYCLES_POMPES.CP_MARCHE).BackColor = COULEURS_ECHELLES_GRAPHIQUES.C_MARCHE_POMPE

    '--- modes du chauffage ---
    PBModesChauffage(MODES_PRODUCTION.M_ARRET).BackColor = COULEURS_ECHELLES_GRAPHIQUES.C_MODE_ARRET
    PBModesChauffage(MODES_PRODUCTION.M_VEILLE).BackColor = COULEURS_ECHELLES_GRAPHIQUES.C_MODE_VEILLE
    PBModesChauffage(MODES_PRODUCTION.M_PRODUCTION).BackColor = COULEURS_ECHELLES_GRAPHIQUES.C_MODE_PRODUCTION

End Sub

Private Sub TBRegulation_LostFocus(Index As Integer)
    On Error Resume Next
    InterdireEvenements = True
    AffichageNombre TBRegulation(Index), TBRegulation(Index).Text, FORMAT_TEMPERATURE_1_DECIMALE
    InterdireEvenements = False
End Sub

'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Rôle      : Détermine la zone d'appareillage concerné
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Private Sub ChoixAppareillage()
    
    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next

    '--- affichage et effacement des zones concernées ---
    With VManipsGestionRegulation
        
        '--- pompe---
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

Private Sub TimerEtatsRegulation_Timer()
    
    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next
    
    '--- appel de la routine ---
    TimerEtatsRegulation.Enabled = False
    EtatsRegulation
    TimerEtatsRegulation.Enabled = True
    
    '--- bip de passage dans la routine UNIQUEMENT POUR LES TESTS ---
    If PROGRAMME_AVEC_AUTOMATE = False Then Beep

End Sub

Private Sub VSChangementCuve_Change()
    
    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next
    
    '--- changement de la cuve en cours ---
    If VSChangementCuve.value >= CUVES_REGULATION.C_C00 And VSChangementCuve.value <= DERNIERE_CUV_REGULATION Then
        InterdireEvenements = True
        CuveEnCours = VSChangementCuve.value
        OBNomsCuves(CuveEnCours).value = True
        LectureValeursFenetre
        InterdireEvenements = False
    End If

End Sub

'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Rôle      : Initialise la fenêtre (chargement ou en vue de la rendre visible)
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Sub InitialisationFenetre()
    
    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next

    '--- déclaration ---
    Dim a As Integer

    '--- divers sur la fenêtre ---
    With Me
        .Caption = TITRE_FENETRE
        .WindowState = vbMaximized
    End With
    
    '--- couleurs des fonds ---
    PBModesProgrammation.Picture = ImgFondVert1
    PBProgrammation.Picture = ImgFondVert1
    PBSelectionNomsCuves.Picture = ImgFondGris1
    PBJourneesTypes.Picture = ImgFondGris1
    PBRegulationAlarme.Picture = ImgFondGris1
    PBTousModesChauffage.Picture = ImgFondGris1
    PBTousCyclesPompe.Picture = ImgFondGris1
    
    PBBoutons.Picture = ImgFondDesBoutons
    
    '--- affichage des noms des cuves ---
    For a = OBNomsCuves.LBound To OBNomsCuves.UBound
        With OBNomsCuves(a)
            .Caption = "Cuve " & TEtatsCuves(a).DefinitionCuve.NomCuve & " - " & TEtatsCuves(a).DefinitionCuve.LibelleCuve
        End With
    Next a
    
    '--- calculs de l'emplacement de la barre de défilement ---
    With PBSelectionNomsCuves
        VSChangementCuve.Top = 0
        VSChangementCuve.Height = .ScaleHeight
    End With

    '--- fixer la valeur maxi de l'ascenseur ---
    With VSChangementCuve
        .min = CUVES_REGULATION.C_C00
        .Max = DERNIERE_CUV_REGULATION
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

    '--- neutralisation du timer ---
    With TimerEtatsRegulation
        .Enabled = False
        .Interval = 0
    End With
    
    '--- déchargement de la fenêtre ---
    Me.Visible = False
    Unload Me
    Set OccFGestionRegulation = Nothing

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
' Rôle      : Effectue le paramètrage de la fenêtre
' Entrées : NumCuve -> Numéro de la cuve
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Sub ParametrageFenetre(ByVal NumCuve As Integer)
    
    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next
    
    '--- numéro de la cuve en cours ---
    CuveEnCours = NumCuve
        
    '--- divers ---
    VSChangementCuve.value = CuveEnCours        'déclenche les évènements
    
    '--- visualisation des différents états de la régulation ---
    EtatsRegulation

    '--- lancement du timer ---
    TimerEtatsRegulation.Enabled = True

End Sub

Private Sub VSDeplacementFENETRE_Change()
    On Error Resume Next
    PBDeplacementFenetre(ZONES_DEPLACEMENT_FENETRE.Z_FILLE).Top = -VSDeplacementFenetre.value
End Sub

VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form FAnnexes 
   BackColor       =   &H00C0C0C0&
   ClientHeight    =   14565
   ClientLeft      =   -1665
   ClientTop       =   4890
   ClientWidth     =   28860
   Icon            =   "FAnnexes.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   14565
   ScaleWidth      =   28860
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox PBBoutons 
      Align           =   2  'Align Bottom
      DrawStyle       =   2  'Dot
      DrawWidth       =   16891
      Height          =   1095
      Left            =   0
      ScaleHeight     =   1035
      ScaleWidth      =   28800
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   13470
      Width           =   28860
      Begin VB.CommandButton CBMarcheArretLigne 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Mise en service de la ligne"
         DownPicture     =   "FAnnexes.frx":014A
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
         Index           =   1
         Left            =   10440
         MaskColor       =   &H00FF00FF&
         Picture         =   "FAnnexes.frx":084C
         Style           =   1  'Graphical
         TabIndex        =   12
         ToolTipText     =   " Mise en service de la ligne "
         Top             =   120
         UseMaskColor    =   -1  'True
         Width           =   3015
      End
      Begin VB.CommandButton CBMarcheArretLigne 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Arrêt de la ligne"
         DownPicture     =   "FAnnexes.frx":0F4E
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
         Index           =   0
         Left            =   6480
         MaskColor       =   &H00FF00FF&
         Picture         =   "FAnnexes.frx":1650
         Style           =   1  'Graphical
         TabIndex        =   11
         ToolTipText     =   " Arrêt de la ligne "
         Top             =   120
         UseMaskColor    =   -1  'True
         Width           =   3015
      End
      Begin VB.Timer TimerEtatsAnnexes 
         Enabled         =   0   'False
         Interval        =   250
         Left            =   1560
         Top             =   120
      End
      Begin VB.PictureBox PBOutilsDeplacementFenetre 
         BackColor       =   &H00E0E0E0&
         Height          =   1035
         Left            =   0
         ScaleHeight     =   975
         ScaleWidth      =   1155
         TabIndex        =   7
         Top             =   0
         Visible         =   0   'False
         Width           =   1215
         Begin VB.CommandButton CBAgrandirFenetre 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Agrandir"
            DownPicture     =   "FAnnexes.frx":1D52
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
            Picture         =   "FAnnexes.frx":1EFC
            Style           =   1  'Graphical
            TabIndex        =   10
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
            TabIndex        =   9
            Top             =   0
            Width           =   255
         End
         Begin VB.HScrollBar HSDeplacementFenetre 
            Height          =   255
            LargeChange     =   300
            Left            =   0
            SmallChange     =   100
            TabIndex        =   8
            Top             =   720
            Width           =   915
         End
      End
      Begin VB.CommandButton CBValider 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Valider"
         DownPicture     =   "FAnnexes.frx":20A6
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
         Picture         =   "FAnnexes.frx":27A8
         Style           =   1  'Graphical
         TabIndex        =   1
         ToolTipText     =   " Valider l'enregistrement "
         Top             =   120
         UseMaskColor    =   -1  'True
         Width           =   1515
      End
      Begin VB.CommandButton CBQuitter 
         BackColor       =   &H00FFFFFF&
         Cancel          =   -1  'True
         Caption         =   "&Quitter"
         DownPicture     =   "FAnnexes.frx":2EAA
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
         Picture         =   "FAnnexes.frx":35AC
         Style           =   1  'Graphical
         TabIndex        =   0
         ToolTipText     =   " Quitter cette fenêtre "
         Top             =   120
         UseMaskColor    =   -1  'True
         Width           =   1515
      End
      Begin MSComctlLib.ImageList ILIndicateursNiveaux 
         Left            =   3360
         Top             =   240
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   21
         ImageHeight     =   53
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   4
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FAnnexes.frx":3CAE
               Key             =   "niveau vide"
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FAnnexes.frx":4A40
               Key             =   "niveau vert"
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FAnnexes.frx":57D2
               Key             =   "niveau orange"
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FAnnexes.frx":6564
               Key             =   "niveau rouge"
            EndProperty
         EndProperty
      End
      Begin VB.Shape SFocus 
         BorderColor     =   &H000000FF&
         BorderWidth     =   5
         Height          =   465
         Left            =   2220
         Top             =   180
         Visible         =   0   'False
         Width           =   540
      End
   End
   Begin VB.PictureBox PBRenseignementsFenetre 
      Align           =   1  'Align Top
      BackColor       =   &H00FF0000&
      ForeColor       =   &H0000FFFF&
      Height          =   375
      Left            =   0
      Picture         =   "FAnnexes.frx":72F6
      ScaleHeight     =   315
      ScaleWidth      =   28800
      TabIndex        =   4
      Top             =   0
      Width           =   28860
      Begin VB.Label LRenseignementsFenetre 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "ANNEXES"
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
         TabIndex        =   5
         Top             =   0
         Width           =   11415
         WordWrap        =   -1  'True
      End
   End
   Begin VB.PictureBox PBDeplacementFenetre 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   13875
      Index           =   0
      Left            =   0
      ScaleHeight     =   13875
      ScaleWidth      =   28860
      TabIndex        =   2
      Top             =   375
      Width           =   28860
      Begin VB.PictureBox PBDeplacementFenetre 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   13515
         Index           =   1
         Left            =   0
         ScaleHeight     =   13515
         ScaleWidth      =   28575
         TabIndex        =   3
         Top             =   0
         Width           =   28575
         Begin VB.PictureBox PBElementsLigne 
            Height          =   11985
            Left            =   240
            ScaleHeight     =   11925
            ScaleWidth      =   10335
            TabIndex        =   45
            Top             =   600
            Width           =   10395
            Begin VB.PictureBox PBGestionEVBrillantage 
               Height          =   3855
               Left            =   1080
               ScaleHeight     =   3795
               ScaleWidth      =   8055
               TabIndex        =   53
               Top             =   540
               Width           =   8115
               Begin VB.PictureBox PBModeEVBrillantage 
                  Appearance      =   0  'Flat
                  BackColor       =   &H00C0E0FF&
                  ForeColor       =   &H80000008&
                  Height          =   2835
                  Left            =   300
                  ScaleHeight     =   2805
                  ScaleWidth      =   7425
                  TabIndex        =   54
                  TabStop         =   0   'False
                  Top             =   660
                  Width           =   7455
                  Begin VB.OptionButton OBModesEVBrillantage 
                     BackColor       =   &H00C0E0FF&
                     Caption         =   "MARCHE FORCEE"
                     BeginProperty Font 
                        Name            =   "MS Sans Serif"
                        Size            =   9.75
                        Charset         =   0
                        Weight          =   700
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     Height          =   252
                     Index           =   2
                     Left            =   540
                     TabIndex        =   70
                     Top             =   1020
                     Width           =   6735
                  End
                  Begin VB.OptionButton OBModesEVBrillantage 
                     BackColor       =   &H00C0E0FF&
                     Caption         =   "MARCHE par INTERMITTENCE + GAMME en AUTOMATIQUE"
                     BeginProperty Font 
                        Name            =   "MS Sans Serif"
                        Size            =   9.75
                        Charset         =   0
                        Weight          =   700
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     Height          =   252
                     Index           =   1
                     Left            =   540
                     TabIndex        =   58
                     Top             =   600
                     Width           =   6735
                  End
                  Begin VB.OptionButton OBModesEVBrillantage 
                     BackColor       =   &H00C0E0FF&
                     Caption         =   "ARRET TOTAL de l'ELECTRO-VANNE"
                     BeginProperty Font 
                        Name            =   "MS Sans Serif"
                        Size            =   9.75
                        Charset         =   0
                        Weight          =   700
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     Height          =   252
                     Index           =   0
                     Left            =   540
                     TabIndex        =   57
                     Top             =   180
                     Width           =   6735
                  End
                  Begin VB.TextBox TBPeriodiciteEVBrillantage 
                     Alignment       =   1  'Right Justify
                     Appearance      =   0  'Flat
                     BackColor       =   &H00C0FFFF&
                     BorderStyle     =   0  'None
                     BeginProperty DataFormat 
                        Type            =   0
                        Format          =   "#0.0 °C"
                        HaveTrueFalseNull=   0
                        FirstDayOfWeek  =   0
                        FirstWeekOfYear =   0
                        LCID            =   1036
                        SubFormatType   =   0
                     EndProperty
                     BeginProperty Font 
                        Name            =   "MS Sans Serif"
                        Size            =   9.75
                        Charset         =   0
                        Weight          =   700
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     ForeColor       =   &H000000FF&
                     Height          =   255
                     Left            =   5100
                     MaxLength       =   4
                     MultiLine       =   -1  'True
                     TabIndex        =   56
                     Top             =   1920
                     Width           =   855
                  End
                  Begin VB.TextBox TBTempsMarcheEVBrillantage 
                     Alignment       =   1  'Right Justify
                     Appearance      =   0  'Flat
                     BackColor       =   &H00C0FFFF&
                     BorderStyle     =   0  'None
                     BeginProperty DataFormat 
                        Type            =   0
                        Format          =   "#0.0 °C"
                        HaveTrueFalseNull=   0
                        FirstDayOfWeek  =   0
                        FirstWeekOfYear =   0
                        LCID            =   1036
                        SubFormatType   =   0
                     EndProperty
                     BeginProperty Font 
                        Name            =   "MS Sans Serif"
                        Size            =   9.75
                        Charset         =   0
                        Weight          =   700
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     ForeColor       =   &H000000FF&
                     Height          =   255
                     Left            =   5100
                     MaxLength       =   4
                     MultiLine       =   -1  'True
                     TabIndex        =   55
                     Top             =   2280
                     Width           =   855
                  End
                  Begin VB.Image IIndicateurEVBrillantage 
                     Height          =   270
                     Index           =   2
                     Left            =   180
                     Picture         =   "FAnnexes.frx":2BC38
                     Top             =   1020
                     Visible         =   0   'False
                     Width           =   255
                  End
                  Begin VB.Image IIndicateurEVBrillantage 
                     Height          =   270
                     Index           =   0
                     Left            =   180
                     Picture         =   "FAnnexes.frx":2C022
                     Top             =   180
                     Visible         =   0   'False
                     Width           =   255
                  End
                  Begin VB.Image IIndicateurEVBrillantage 
                     Height          =   270
                     Index           =   1
                     Left            =   180
                     Picture         =   "FAnnexes.frx":2C40C
                     Top             =   600
                     Visible         =   0   'False
                     Width           =   255
                  End
                  Begin VB.Label LLibelles 
                     Alignment       =   2  'Center
                     Appearance      =   0  'Flat
                     BackColor       =   &H00C0FFFF&
                     BorderStyle     =   1  'Fixed Single
                     Caption         =   "TEMPS de FONCTIONNEMENT"
                     BeginProperty Font 
                        Name            =   "Arial Narrow"
                        Size            =   12
                        Charset         =   0
                        Weight          =   700
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     ForeColor       =   &H80000008&
                     Height          =   375
                     Index           =   12
                     Left            =   180
                     TabIndex        =   63
                     Top             =   1500
                     Width           =   7035
                  End
                  Begin VB.Label LLibelles 
                     Appearance      =   0  'Flat
                     BackColor       =   &H00FFFFFF&
                     Caption         =   "PERIODICITE de MISE en MARCHE"
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
                     Index           =   0
                     Left            =   300
                     TabIndex        =   62
                     Top             =   1920
                     Width           =   4575
                  End
                  Begin VB.Label LLibelles 
                     Appearance      =   0  'Flat
                     BackColor       =   &H00FFFFFF&
                     Caption         =   "TEMPS de MARCHE"
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
                     Index           =   1
                     Left            =   300
                     TabIndex        =   61
                     Top             =   2280
                     Width           =   4575
                  End
                  Begin VB.Label LLibelles 
                     Appearance      =   0  'Flat
                     BackColor       =   &H00FFFFFF&
                     BackStyle       =   0  'Transparent
                     Caption         =   "Minutes"
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
                     Left            =   6060
                     TabIndex        =   60
                     Top             =   1920
                     Width           =   1095
                  End
                  Begin VB.Label LLibelles 
                     Appearance      =   0  'Flat
                     BackColor       =   &H00FFFFFF&
                     BackStyle       =   0  'Transparent
                     Caption         =   "Secondes"
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
                     Index           =   4
                     Left            =   6060
                     TabIndex        =   59
                     Top             =   2280
                     Width           =   1095
                  End
                  Begin VB.Shape SDecoration 
                     BackColor       =   &H00FFFFFF&
                     BackStyle       =   1  'Opaque
                     Height          =   375
                     Index           =   1
                     Left            =   180
                     Top             =   2220
                     Width           =   4815
                  End
                  Begin VB.Shape SDecoration 
                     BackColor       =   &H00FFFFFF&
                     BackStyle       =   1  'Opaque
                     Height          =   375
                     Index           =   0
                     Left            =   180
                     Top             =   1860
                     Width           =   4815
                  End
                  Begin VB.Shape SDecoration 
                     BackColor       =   &H00C0FFFF&
                     BackStyle       =   1  'Opaque
                     Height          =   375
                     Index           =   3
                     Left            =   4980
                     Top             =   1860
                     Width           =   2235
                  End
                  Begin VB.Shape SDecoration 
                     BackColor       =   &H00C0FFFF&
                     BackStyle       =   1  'Opaque
                     Height          =   375
                     Index           =   2
                     Left            =   4980
                     Top             =   2220
                     Width           =   2235
                  End
               End
               Begin VB.Label LEtats 
                  Alignment       =   2  'Center
                  Appearance      =   0  'Flat
                  BackColor       =   &H80000005&
                  BackStyle       =   0  'Transparent
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
                  Left            =   5340
                  TabIndex        =   65
                  Top             =   360
                  Width           =   2355
               End
               Begin VB.Label LLibellesEtats 
                  Appearance      =   0  'Flat
                  BackColor       =   &H00FFFFFF&
                  Caption         =   "ETAT ACTUEL de l'ELECTRO-VANNE"
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
                  Index           =   7
                  Left            =   420
                  TabIndex        =   64
                  Top             =   360
                  Width           =   4815
               End
               Begin VB.Shape SCadresEtats 
                  BackColor       =   &H00E0E0E0&
                  BackStyle       =   1  'Opaque
                  Height          =   375
                  Index           =   1
                  Left            =   5280
                  Top             =   300
                  Width           =   2475
               End
               Begin VB.Shape SCadresLibellesEtats 
                  BackColor       =   &H00FFFFFF&
                  BackStyle       =   1  'Opaque
                  Height          =   375
                  Index           =   0
                  Left            =   300
                  Top             =   300
                  Width           =   4995
               End
            End
            Begin VB.PictureBox PBGestionEVEauLigne 
               Height          =   2115
               Left            =   1080
               ScaleHeight     =   2055
               ScaleWidth      =   8055
               TabIndex        =   47
               Top             =   4920
               Width           =   8115
               Begin VB.PictureBox PBGestionModeEVEauLigne 
                  Appearance      =   0  'Flat
                  BackColor       =   &H00C0E0FF&
                  ForeColor       =   &H80000008&
                  Height          =   1095
                  Left            =   300
                  ScaleHeight     =   1065
                  ScaleWidth      =   7425
                  TabIndex        =   48
                  TabStop         =   0   'False
                  Top             =   660
                  Width           =   7455
                  Begin VB.OptionButton OBModesEVEauLigne 
                     BackColor       =   &H00C0E0FF&
                     Caption         =   "ARRET de l'ELECTRO-VANNE d'ARRIVEE d'EAU"
                     BeginProperty Font 
                        Name            =   "MS Sans Serif"
                        Size            =   9.75
                        Charset         =   0
                        Weight          =   700
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     Height          =   252
                     Index           =   0
                     Left            =   540
                     TabIndex        =   50
                     Top             =   180
                     Width           =   6735
                  End
                  Begin VB.OptionButton OBModesEVEauLigne 
                     BackColor       =   &H00C0E0FF&
                     Caption         =   "MARCHE de l'ELECTRO-VANNE d'ARRIVEE d'EAU"
                     BeginProperty Font 
                        Name            =   "MS Sans Serif"
                        Size            =   9.75
                        Charset         =   0
                        Weight          =   700
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     Height          =   252
                     Index           =   1
                     Left            =   540
                     TabIndex        =   49
                     Top             =   600
                     Width           =   6735
                  End
                  Begin VB.Image IIndicateurEVEauLigne 
                     Height          =   270
                     Index           =   1
                     Left            =   180
                     Picture         =   "FAnnexes.frx":2C7F6
                     Top             =   600
                     Visible         =   0   'False
                     Width           =   255
                  End
                  Begin VB.Image IIndicateurEVEauLigne 
                     Height          =   270
                     Index           =   0
                     Left            =   180
                     Picture         =   "FAnnexes.frx":2CBE0
                     Top             =   180
                     Visible         =   0   'False
                     Width           =   255
                  End
               End
               Begin VB.Label LLibellesEtats 
                  Appearance      =   0  'Flat
                  BackColor       =   &H00FFFFFF&
                  Caption         =   "ETAT ACTUEL de l'ELECTRO-VANNE"
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
                  Index           =   4
                  Left            =   420
                  TabIndex        =   52
                  Top             =   360
                  Width           =   4815
               End
               Begin VB.Label LEtats 
                  Alignment       =   2  'Center
                  Appearance      =   0  'Flat
                  BackColor       =   &H80000005&
                  BackStyle       =   0  'Transparent
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
                  Left            =   5340
                  TabIndex        =   51
                  Top             =   360
                  Width           =   2355
               End
               Begin VB.Shape SCadresEtats 
                  BackColor       =   &H00E0E0E0&
                  BackStyle       =   1  'Opaque
                  Height          =   375
                  Index           =   2
                  Left            =   5280
                  Top             =   300
                  Width           =   2475
               End
               Begin VB.Shape SCadresLibellesEtats 
                  BackColor       =   &H00FFFFFF&
                  BackStyle       =   1  'Opaque
                  Height          =   375
                  Index           =   5
                  Left            =   300
                  Top             =   300
                  Width           =   4995
               End
            End
            Begin VB.PictureBox PBNiveauxRetentions 
               AutoSize        =   -1  'True
               Height          =   4155
               Left            =   180
               Picture         =   "FAnnexes.frx":2CFCA
               ScaleHeight     =   4095
               ScaleWidth      =   9885
               TabIndex        =   46
               Top             =   7560
               Width           =   9945
               Begin VB.Image IIndicateursNiveaux 
                  Height          =   795
                  Index           =   3
                  Left            =   9240
                  Picture         =   "FAnnexes.frx":B0F88
                  Top             =   2970
                  Width           =   315
               End
               Begin VB.Image IIndicateursNiveaux 
                  Height          =   795
                  Index           =   2
                  Left            =   7920
                  Picture         =   "FAnnexes.frx":B1D0A
                  Top             =   2970
                  Width           =   315
               End
               Begin VB.Image IIndicateursNiveaux 
                  Height          =   795
                  Index           =   1
                  Left            =   7920
                  Picture         =   "FAnnexes.frx":B2A8C
                  Top             =   1650
                  Width           =   315
               End
               Begin VB.Image IIndicateursNiveaux 
                  Height          =   795
                  Index           =   0
                  Left            =   7920
                  Picture         =   "FAnnexes.frx":B380E
                  Top             =   330
                  Width           =   315
               End
            End
            Begin VB.Label LLibelles 
               Alignment       =   2  'Center
               BackColor       =   &H00C00000&
               BorderStyle     =   1  'Fixed Single
               Caption         =   "GESTION de l'ELECTRO-VANNE d'AIR dans le BAIN de BRILLANTAGE"
               BeginProperty Font 
                  Name            =   "Arial Narrow"
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
               Left            =   1080
               TabIndex        =   68
               Top             =   180
               Width           =   8115
            End
            Begin VB.Label LLibelles 
               Alignment       =   2  'Center
               BackColor       =   &H00C00000&
               BorderStyle     =   1  'Fixed Single
               Caption         =   "GESTION de l'ELECTRO-VANNE d'ARRIVEE d'EAU de la LIGNE"
               BeginProperty Font 
                  Name            =   "Arial Narrow"
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
               Left            =   1080
               TabIndex        =   67
               Top             =   4560
               Width           =   8115
            End
            Begin VB.Label LLibelles 
               Alignment       =   2  'Center
               BackColor       =   &H00C00000&
               BorderStyle     =   1  'Fixed Single
               Caption         =   "NIVEAUX des RETENTIONS"
               BeginProperty Font 
                  Name            =   "Arial Narrow"
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
               Left            =   180
               TabIndex        =   66
               Top             =   7200
               Width           =   9945
            End
         End
         Begin VB.PictureBox PBPont2 
            Height          =   5535
            Left            =   19740
            ScaleHeight     =   5475
            ScaleWidth      =   8475
            TabIndex        =   29
            Top             =   600
            Width           =   8535
            Begin VB.PictureBox PBGestionEclairageP2 
               Height          =   2115
               Left            =   180
               ScaleHeight     =   2055
               ScaleWidth      =   8055
               TabIndex        =   36
               Top             =   3180
               Width           =   8115
               Begin VB.PictureBox Picture2 
                  Appearance      =   0  'Flat
                  BackColor       =   &H00C0E0FF&
                  ForeColor       =   &H80000008&
                  Height          =   1095
                  Left            =   300
                  ScaleHeight     =   1065
                  ScaleWidth      =   7425
                  TabIndex        =   37
                  TabStop         =   0   'False
                  Top             =   660
                  Width           =   7455
                  Begin VB.OptionButton OBModesEclairageP2 
                     BackColor       =   &H00C0E0FF&
                     Caption         =   "MARCHE de l'ECLAIRAGE du PONT 2"
                     BeginProperty Font 
                        Name            =   "MS Sans Serif"
                        Size            =   9.75
                        Charset         =   0
                        Weight          =   700
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     Height          =   252
                     Index           =   1
                     Left            =   540
                     TabIndex        =   39
                     Top             =   600
                     Width           =   6735
                  End
                  Begin VB.OptionButton OBModesEclairageP2 
                     BackColor       =   &H00C0E0FF&
                     Caption         =   "ARRET de l'ECLAIRAGE du PONT 2"
                     BeginProperty Font 
                        Name            =   "MS Sans Serif"
                        Size            =   9.75
                        Charset         =   0
                        Weight          =   700
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     Height          =   252
                     Index           =   0
                     Left            =   540
                     TabIndex        =   38
                     Top             =   180
                     Width           =   6735
                  End
                  Begin VB.Image IIndicateurEclairageP2 
                     Height          =   270
                     Index           =   0
                     Left            =   180
                     Picture         =   "FAnnexes.frx":B4590
                     Top             =   180
                     Visible         =   0   'False
                     Width           =   255
                  End
                  Begin VB.Image IIndicateurEclairageP2 
                     Height          =   270
                     Index           =   1
                     Left            =   180
                     Picture         =   "FAnnexes.frx":B497A
                     Top             =   600
                     Visible         =   0   'False
                     Width           =   255
                  End
               End
               Begin VB.Label LEtats 
                  Alignment       =   2  'Center
                  Appearance      =   0  'Flat
                  BackColor       =   &H80000005&
                  BackStyle       =   0  'Transparent
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
                  Index           =   6
                  Left            =   5340
                  TabIndex        =   41
                  Top             =   360
                  Width           =   2355
               End
               Begin VB.Label LLibellesEtats 
                  Appearance      =   0  'Flat
                  BackColor       =   &H00FFFFFF&
                  Caption         =   "ETAT ACTUEL de l'ECLAIRAGE du PONT 2"
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
                  Index           =   1
                  Left            =   420
                  TabIndex        =   40
                  Top             =   360
                  Width           =   4815
               End
               Begin VB.Shape SCadresEtats 
                  BackColor       =   &H00E0E0E0&
                  BackStyle       =   1  'Opaque
                  Height          =   375
                  Index           =   6
                  Left            =   5280
                  Top             =   300
                  Width           =   2475
               End
               Begin VB.Shape SCadresLibellesEtats 
                  BackColor       =   &H00FFFFFF&
                  BackStyle       =   1  'Opaque
                  Height          =   375
                  Index           =   2
                  Left            =   300
                  Top             =   300
                  Width           =   4995
               End
            End
            Begin VB.PictureBox PBGestionCompresseurP2 
               Height          =   2115
               Left            =   180
               ScaleHeight     =   2055
               ScaleWidth      =   8055
               TabIndex        =   30
               Top             =   540
               Width           =   8115
               Begin VB.PictureBox PBGestionModeCompresseurP2 
                  Appearance      =   0  'Flat
                  BackColor       =   &H00C0E0FF&
                  ForeColor       =   &H80000008&
                  Height          =   1095
                  Left            =   300
                  ScaleHeight     =   1065
                  ScaleWidth      =   7425
                  TabIndex        =   31
                  TabStop         =   0   'False
                  Top             =   660
                  Width           =   7455
                  Begin VB.OptionButton OBModesCompresseurP2 
                     BackColor       =   &H00C0E0FF&
                     Caption         =   "ARRET du COMPRESSEUR du PONT 2"
                     BeginProperty Font 
                        Name            =   "MS Sans Serif"
                        Size            =   9.75
                        Charset         =   0
                        Weight          =   700
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     Height          =   252
                     Index           =   0
                     Left            =   540
                     TabIndex        =   33
                     Top             =   180
                     Width           =   6735
                  End
                  Begin VB.OptionButton OBModesCompresseurP2 
                     BackColor       =   &H00C0E0FF&
                     Caption         =   "MARCHE du COMPRESSEUR du PONT 2"
                     BeginProperty Font 
                        Name            =   "MS Sans Serif"
                        Size            =   9.75
                        Charset         =   0
                        Weight          =   700
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     Height          =   252
                     Index           =   1
                     Left            =   540
                     TabIndex        =   32
                     Top             =   600
                     Width           =   6735
                  End
                  Begin VB.Image IIndicateurCompresseurP2 
                     Height          =   270
                     Index           =   1
                     Left            =   180
                     Picture         =   "FAnnexes.frx":B4D64
                     Top             =   600
                     Visible         =   0   'False
                     Width           =   255
                  End
                  Begin VB.Image IIndicateurCompresseurP2 
                     Height          =   270
                     Index           =   0
                     Left            =   180
                     Picture         =   "FAnnexes.frx":B514E
                     Top             =   180
                     Visible         =   0   'False
                     Width           =   255
                  End
               End
               Begin VB.Label LLibellesEtats 
                  Appearance      =   0  'Flat
                  BackColor       =   &H00FFFFFF&
                  Caption         =   "ETAT ACTUEL du COMPRESSEUR du PONT 2"
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
                  Left            =   360
                  TabIndex        =   35
                  Top             =   360
                  Width           =   4875
               End
               Begin VB.Label LEtats 
                  Alignment       =   2  'Center
                  Appearance      =   0  'Flat
                  BackColor       =   &H80000005&
                  BackStyle       =   0  'Transparent
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
                  Left            =   5340
                  TabIndex        =   34
                  Top             =   360
                  Width           =   2355
               End
               Begin VB.Shape SCadresLibellesEtats 
                  BackColor       =   &H00FFFFFF&
                  BackStyle       =   1  'Opaque
                  Height          =   375
                  Index           =   4
                  Left            =   300
                  Top             =   300
                  Width           =   4995
               End
               Begin VB.Shape SCadresEtats 
                  BackColor       =   &H00E0E0E0&
                  BackStyle       =   1  'Opaque
                  Height          =   375
                  Index           =   5
                  Left            =   5280
                  Top             =   300
                  Width           =   2475
               End
            End
            Begin VB.Label LLibelles 
               Alignment       =   2  'Center
               BackColor       =   &H00C00000&
               BorderStyle     =   1  'Fixed Single
               Caption         =   "GESTION de l'ECLAIRAGE du PONT 2"
               BeginProperty Font 
                  Name            =   "Arial Narrow"
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
               Left            =   180
               TabIndex        =   43
               Top             =   2820
               Width           =   8115
            End
            Begin VB.Label LLibelles 
               Alignment       =   2  'Center
               BackColor       =   &H00C00000&
               BorderStyle     =   1  'Fixed Single
               Caption         =   "GESTION du COMPRESSEUR du PONT 2"
               BeginProperty Font 
                  Name            =   "Arial Narrow"
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
               Left            =   180
               TabIndex        =   42
               Top             =   180
               Width           =   8115
            End
         End
         Begin VB.PictureBox PBPont1 
            Height          =   5535
            Left            =   10920
            ScaleHeight     =   5475
            ScaleWidth      =   8475
            TabIndex        =   13
            Top             =   600
            Width           =   8535
            Begin VB.PictureBox PBGestionEclairageP1 
               Height          =   2115
               Left            =   180
               ScaleHeight     =   2055
               ScaleWidth      =   8055
               TabIndex        =   20
               Top             =   3180
               Width           =   8115
               Begin VB.PictureBox PBModeGestionEclairageP1 
                  Appearance      =   0  'Flat
                  BackColor       =   &H00C0E0FF&
                  ForeColor       =   &H80000008&
                  Height          =   1095
                  Left            =   300
                  ScaleHeight     =   1065
                  ScaleWidth      =   7425
                  TabIndex        =   21
                  TabStop         =   0   'False
                  Top             =   660
                  Width           =   7455
                  Begin VB.OptionButton OBModesEclairageP1 
                     BackColor       =   &H00C0E0FF&
                     Caption         =   "ARRET de l'ECLAIRAGE du PONT 1"
                     BeginProperty Font 
                        Name            =   "MS Sans Serif"
                        Size            =   9.75
                        Charset         =   0
                        Weight          =   700
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     Height          =   252
                     Index           =   0
                     Left            =   540
                     TabIndex        =   23
                     Top             =   180
                     Width           =   6735
                  End
                  Begin VB.OptionButton OBModesEclairageP1 
                     BackColor       =   &H00C0E0FF&
                     Caption         =   "MARCHE de l'ECLAIRAGE du PONT 1"
                     BeginProperty Font 
                        Name            =   "MS Sans Serif"
                        Size            =   9.75
                        Charset         =   0
                        Weight          =   700
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     Height          =   252
                     Index           =   1
                     Left            =   540
                     TabIndex        =   22
                     Top             =   600
                     Width           =   6735
                  End
                  Begin VB.Image IIndicateurEclairageP1 
                     Height          =   270
                     Index           =   1
                     Left            =   180
                     Picture         =   "FAnnexes.frx":B5538
                     Top             =   600
                     Visible         =   0   'False
                     Width           =   255
                  End
                  Begin VB.Image IIndicateurEclairageP1 
                     Height          =   270
                     Index           =   0
                     Left            =   180
                     Picture         =   "FAnnexes.frx":B5922
                     Top             =   180
                     Visible         =   0   'False
                     Width           =   255
                  End
               End
               Begin VB.Label LLibellesEtats 
                  Appearance      =   0  'Flat
                  BackColor       =   &H00FFFFFF&
                  Caption         =   "ETAT ACTUEL de l'ECLAIRAGE du PONT 1"
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
                  Index           =   0
                  Left            =   420
                  TabIndex        =   25
                  Top             =   360
                  Width           =   4815
               End
               Begin VB.Label LEtats 
                  Alignment       =   2  'Center
                  Appearance      =   0  'Flat
                  BackColor       =   &H80000005&
                  BackStyle       =   0  'Transparent
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
                  Left            =   5340
                  TabIndex        =   24
                  Top             =   360
                  Width           =   2355
               End
               Begin VB.Shape SCadresEtats 
                  BackColor       =   &H00E0E0E0&
                  BackStyle       =   1  'Opaque
                  Height          =   375
                  Index           =   4
                  Left            =   5280
                  Top             =   300
                  Width           =   2475
               End
               Begin VB.Shape SCadresLibellesEtats 
                  BackColor       =   &H00FFFFFF&
                  BackStyle       =   1  'Opaque
                  Height          =   375
                  Index           =   1
                  Left            =   300
                  Top             =   300
                  Width           =   4995
               End
            End
            Begin VB.PictureBox PBGestionCompresseurP1 
               Height          =   2115
               Left            =   180
               ScaleHeight     =   2055
               ScaleWidth      =   8055
               TabIndex        =   14
               Top             =   540
               Width           =   8115
               Begin VB.PictureBox PBGestionModeCompresseurP1 
                  Appearance      =   0  'Flat
                  BackColor       =   &H00C0E0FF&
                  ForeColor       =   &H80000008&
                  Height          =   1095
                  Left            =   300
                  ScaleHeight     =   1065
                  ScaleWidth      =   7425
                  TabIndex        =   15
                  TabStop         =   0   'False
                  Top             =   660
                  Width           =   7455
                  Begin VB.OptionButton OBModesCompresseurP1 
                     BackColor       =   &H00C0E0FF&
                     Caption         =   "MARCHE du COMPRESSEUR du PONT 1"
                     BeginProperty Font 
                        Name            =   "MS Sans Serif"
                        Size            =   9.75
                        Charset         =   0
                        Weight          =   700
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     Height          =   252
                     Index           =   1
                     Left            =   540
                     TabIndex        =   17
                     Top             =   600
                     Width           =   6735
                  End
                  Begin VB.OptionButton OBModesCompresseurP1 
                     BackColor       =   &H00C0E0FF&
                     Caption         =   "ARRET du COMPRESSEUR du PONT 1"
                     BeginProperty Font 
                        Name            =   "MS Sans Serif"
                        Size            =   9.75
                        Charset         =   0
                        Weight          =   700
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     Height          =   252
                     Index           =   0
                     Left            =   540
                     TabIndex        =   16
                     Top             =   180
                     Width           =   6735
                  End
                  Begin VB.Image IIndicateurCompresseurP1 
                     Height          =   270
                     Index           =   0
                     Left            =   180
                     Picture         =   "FAnnexes.frx":B5D0C
                     Top             =   180
                     Visible         =   0   'False
                     Width           =   255
                  End
                  Begin VB.Image IIndicateurCompresseurP1 
                     Height          =   270
                     Index           =   1
                     Left            =   180
                     Picture         =   "FAnnexes.frx":B60F6
                     Top             =   600
                     Visible         =   0   'False
                     Width           =   255
                  End
               End
               Begin VB.Label LEtats 
                  Alignment       =   2  'Center
                  Appearance      =   0  'Flat
                  BackColor       =   &H80000005&
                  BackStyle       =   0  'Transparent
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
                  Left            =   5340
                  TabIndex        =   19
                  Top             =   360
                  Width           =   2355
               End
               Begin VB.Label LLibellesEtats 
                  Appearance      =   0  'Flat
                  BackColor       =   &H00FFFFFF&
                  Caption         =   "ETAT ACTUEL du COMPRESSEUR du PONT 1"
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
                  Left            =   360
                  TabIndex        =   18
                  Top             =   360
                  Width           =   4875
               End
               Begin VB.Shape SCadresLibellesEtats 
                  BackColor       =   &H00FFFFFF&
                  BackStyle       =   1  'Opaque
                  Height          =   375
                  Index           =   3
                  Left            =   300
                  Top             =   300
                  Width           =   4995
               End
               Begin VB.Shape SCadresEtats 
                  BackColor       =   &H00E0E0E0&
                  BackStyle       =   1  'Opaque
                  Height          =   375
                  Index           =   3
                  Left            =   5280
                  Top             =   300
                  Width           =   2475
               End
            End
            Begin VB.Label LLibelles 
               Alignment       =   2  'Center
               BackColor       =   &H00C00000&
               BorderStyle     =   1  'Fixed Single
               Caption         =   "GESTION de l'ECLAIRAGE du PONT 1"
               BeginProperty Font 
                  Name            =   "Arial Narrow"
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
               Left            =   180
               TabIndex        =   27
               Top             =   2820
               Width           =   8115
            End
            Begin VB.Label LLibelles 
               Alignment       =   2  'Center
               BackColor       =   &H00C00000&
               BorderStyle     =   1  'Fixed Single
               Caption         =   "GESTION du COMPRESSEUR du PONT 1"
               BeginProperty Font 
                  Name            =   "Arial Narrow"
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
               Left            =   180
               TabIndex        =   26
               Top             =   180
               Width           =   8115
            End
         End
         Begin VB.Label LLibelles 
            Alignment       =   2  'Center
            BackColor       =   &H00C00000&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "ELEMENTS de la LIGNE et STATION des EAUX et VAPEURS"
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   375
            Index           =   14
            Left            =   240
            TabIndex        =   69
            Top             =   240
            Width           =   10395
         End
         Begin VB.Label LLibelles 
            Alignment       =   2  'Center
            BackColor       =   &H00C00000&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "PONT 2"
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   375
            Index           =   13
            Left            =   19740
            TabIndex        =   44
            Top             =   240
            Width           =   8535
         End
         Begin VB.Label LLibelles 
            Alignment       =   2  'Center
            BackColor       =   &H00C00000&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "PONT 1"
            BeginProperty Font 
               Name            =   "Arial Narrow"
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
            Left            =   10920
            TabIndex        =   28
            Top             =   240
            Width           =   8535
         End
      End
   End
End
Attribute VB_Name = "FAnnexes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Rôle                    : Fenêtre gérant les annexes
' Nom                    : FAnnexes.frm
' Date de création : 30/05/2011
' Détails                :
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

'--- déclarations obligatoires ---
Option Explicit

'--- options générales ---
Option Base 1
DefVar A-Z

'--- constantes privées ---
Private Const TITRE_FENETRE As String = "ANNEXES"
Private Const TITRE_MESSAGES As String = TITRE_FENETRE

Private Const ADR_API_MODE_EV_BRILLANTAGE As String = "MW_Mode_EV_Brillantage"
Private Const ADR_API_PERIODICITE_MARCHE_EV_BRILLANTAGE As String = "PeriodiciteEVBrillantage"
Private Const ADR_API_TEMPS_MARCHE_EV_BRILLANTAGE As String = "TempsMarcheEVBrillantage"

Private Const ADR_API_MODE_EV_EAU_LIGNE  As String = "MW_Mode_EV_Eau_Ligne"

Private Const ADR_API_MODE_COMPRESSEUR_P1  As String = "MW_Mode_Compres_P1"
Private Const ADR_API_MODE_ECLAIRAGE_P1  As String = "MW_Mode_Eclairage_P1"

Private Const ADR_API_MODE_COMPRESSEUR_P2  As String = "MW_Mode_Compres_P2"
Private Const ADR_API_MODE_ECLAIRAGE_P2  As String = "MW_Mode_Eclairage_P2"

'--- énumérations privées ---
Private Enum INDEX_CHAMPS
    
    IDX_CHAMP_ETATS_EV_BRILLANTAGE = 1
    
    IDX_CHAMP_ETATS_EV_EAU_LIGNE = 2
    
    IDX_CHAMP_ETATS_COMPRESSEUR_P1 = 3
    IDX_CHAMP_ETATS_ECLAIRAGE_P1 = 4
    
    IDX_CHAMP_ETATS_COMPRESSEUR_P2 = 5
    IDX_CHAMP_ETATS_ECLAIRAGE_P2 = 6
    
End Enum

'--- variables privées ---
Private PremiereActivation As Boolean
Private InterdireEvenements As Boolean                      'pour interdire certains évènements

Private MemModeEVBrillantage As Integer                   'mémoire du mode de l'électro-vanne d'air du brillantage
Private MemModeEVEauLigne As Integer                     'mémoire du mode de l'électro-vanne d'eau de la ligne

Private MemModeCompresseurP1 As Integer              'mémoire du mode du compresseur du pont 1
Private MemModeEclairageP1 As Integer                     'mémoire du mode de l'éclairage du pont 1

Private MemModeCompresseurP2 As Integer              'mémoire du mode du compresseur du pont 2
Private MemModeEclairageP2 As Integer                     'mémoire du mode de l'éclairage du pont 2

'--- tableaux privées ---

'--- variables publiques ---
Public NumFenetre As Long                                          'numéro de la fenêtre lorsqu'elle devient active

Private Sub CBAgrandirFENETRE_Click()
    On Error Resume Next
    Me.WindowState = vbMaximized
End Sub

Private Sub CBMarcheArretLigne_Click(Index As Integer)
    
    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next

    If AppelFenetre(F_MESSAGE, _
                             TITRE_MESSAGES, _
                             IIf(Index = 0, MESSAGE_132, MESSAGE_133), _
                             TYPES_MESSAGES.T_AVERTISSEMENT, _
                             TYPES_BOUTONS.T_OUI_NON, _
                             EMPLACEMENT_FOCUS.E_SUR_NON) = vbYes Then

        '--- forcer l'arrêt ---
        OBModesEVEauLigne(Index).Value = True
        
        OBModesCompresseurP1(Index).Value = True
        OBModesEclairageP1(Index).Value = True
        
        OBModesCompresseurP2(Index).Value = True
        OBModesEclairageP2(Index).Value = True

        '--- curseur de la souris ---
        SourisEnAttente True
    
        '--- enregistre les valeurs dans l'automate ---
        If EnregistreVersAutomate = True Then
    
            '--- enregistrement des valeurs de la fenêtre ---
            EnregistreValeursfenetre
    
            '--- sauvegarde des annexes ---
            SauveAnnexes
    
            '--- affectation ---
            CBValider.Enabled = False
    
        Else
    
            '--- lancer un message d'erreur ---
            Bidon = MessageErreur(TITRE_MESSAGES, MESSAGE_500)
    
            '--- chargement des annexes ---
            ChargeAnnexes
    
            '--- restitue les dernières manipulations et valeurs sur la fenêtre ---
            LectureValeursFenetre
    
        End If
        
        '--- curseur de la souris ---
        SourisEnAttente False

    End If

End Sub

Private Sub CBMarcheArretLigne_GotFocus(Index As Integer)
    
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

Private Sub CBMarcheArretLigne_LostFocus(Index As Integer)
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

        '--- sauvegarde des annexes ---
        SauveAnnexes

        '--- affectation ---
        CBValider.Enabled = False

    Else

        '--- lancer un message d'erreur ---
        Bidon = MessageErreur(TITRE_MESSAGES, MESSAGE_500)

        '--- chargement des annexes ---
        ChargeAnnexes

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

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    
    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next
    
    Select Case KeyCode
        
        Case vbKeyF1 To vbKeyF11
            '--- touches de fonctions ---
            OccFSynoptique.SetFocus
            Call OccFSynoptique.GestionTouches(KeyCode, Shift)
        
        Case vbKeyF12
            '--- acquittement des alarmes ---
            AcquittementAlarmes
        
        Case Else
    End Select

End Sub

Private Sub Form_Resize()
    
    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next
    
    '--- déclaration ---
    
    '--- zone mére et fille du déplacement de la fenetre ---
    PBDeplacementFenetre(ZONES_DEPLACEMENT_FENETRE.Z_MERE).Height = Abs(Me.ScaleHeight - PBBoutons.Height)
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

Private Sub OBModesCompresseurP1_Click(Index As Integer)
    On Error Resume Next
    MemModeCompresseurP1 = Index
    CBValider.Enabled = True
End Sub

Private Sub OBModesCompresseurP2_Click(Index As Integer)
    On Error Resume Next
    MemModeCompresseurP2 = Index
    CBValider.Enabled = True
End Sub

Private Sub OBModesEclairageP1_Click(Index As Integer)
    On Error Resume Next
    MemModeEclairageP1 = Index
    CBValider.Enabled = True
End Sub

Private Sub OBModesEclairageP2_Click(Index As Integer)
    On Error Resume Next
    MemModeEclairageP2 = Index
    CBValider.Enabled = True
End Sub

Private Sub OBModesEVBrillantage_Click(Index As Integer)
    On Error Resume Next
    MemModeEVBrillantage = Index
    CBValider.Enabled = True
End Sub

Private Sub OBModesEVEauLigne_Click(Index As Integer)
    On Error Resume Next
    MemModeEVEauLigne = Index
    CBValider.Enabled = True
End Sub

Private Sub PBBoutons_Resize()
    
    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next
    
    '--- calculs de l'emplacement des boutons ---
    CBQuitter.Left = PBBoutons.ScaleWidth - MARGES.M_BORD_DROIT - CBQuitter.Width
    CBValider.Left = CBQuitter.Left - MARGES.M_ENTRE_BOUTONS - CBValider.Width
    CBMarcheArretLigne(1).Left = CBValider.Left - 56 * MARGES.M_ENTRE_BOUTONS - CBMarcheArretLigne(1).Width
    CBMarcheArretLigne(0).Left = CBMarcheArretLigne(1).Left - MARGES.M_ENTRE_BOUTONS - CBMarcheArretLigne(0).Width
    
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
    
    With TEtatsAnnexes
        
        '--- électro-vanne d'air dans le bain de brillantage ---
        Select Case .ModeEVBrillantage
            Case MODES_EV_BRILLANTAGE.M_ARRET To MODES_EV_BRILLANTAGE.M_MARCHE_FORCEE
                OBModesEVBrillantage(.ModeEVBrillantage).Value = True
            Case Else
        End Select
        AffichageNombre TBPeriodiciteEVBrillantage, .PeriodiciteEVBrillantage, FORMAT_NATURELS_3_CHIFFRES
        AffichageNombre TBTempsMarcheEVBrillantage, .TempsMarcheEVBrillantage, FORMAT_NATURELS_4_CHIFFRES
    
        '--- électro-vanne d'arrivée d'eau de la ligne ---
        Select Case .ModeEVEauLigne
            Case MODES_EV_EAU_LIGNE.M_ARRET To MODES_EV_EAU_LIGNE.M_MARCHE
                OBModesEVEauLigne(.ModeEVEauLigne).Value = True
            Case Else
        End Select
        
        '--- compresseur du pont 1 ---
        Select Case .ModeCompresseurP1
            Case MODES_COMPRESSEURS_PONTS.M_ARRET To MODES_COMPRESSEURS_PONTS.M_MARCHE
                OBModesCompresseurP1(.ModeCompresseurP1).Value = True
            Case Else
        End Select
        
        '--- éclairage du pont 1 ---
        Select Case .ModeEclairageP1
            Case MODES_ECLAIRAGE_PONTS.M_ARRET To MODES_ECLAIRAGE_PONTS.M_MARCHE
                OBModesEclairageP1(.ModeEclairageP1).Value = True
            Case Else
        End Select
        
        '--- compresseur du pont 2 ---
        Select Case .ModeCompresseurP2
            Case MODES_COMPRESSEURS_PONTS.M_ARRET To MODES_COMPRESSEURS_PONTS.M_MARCHE
                OBModesCompresseurP2(.ModeCompresseurP2).Value = True
            Case Else
        End Select
        
        '--- éclairage du pont 2 ---
        Select Case .ModeEclairageP2
            Case MODES_ECLAIRAGE_PONTS.M_ARRET To MODES_ECLAIRAGE_PONTS.M_MARCHE
                OBModesEclairageP2(.ModeEclairageP2).Value = True
            Case Else
        End Select

    End With

    '--- affactation ---
    CBValider.Enabled = False
    
    '--- autoriser les évènements ---
    InterdireEvenements = False

End Sub

'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Rôle      : Enregistre les dernières manipulations et valeurs de la fenêtre
' Entrées :
' Retours :
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Private Sub EnregistreValeursfenetre()
    
    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next

    '--- déclaration ---
    Dim a As Integer

    With TEtatsAnnexes
        
        '--- électro-vanne d'air dans le bain de brillantage ---
        .ModeEVBrillantage = MemModeEVBrillantage
        .PeriodiciteEVBrillantage = ControleSiNombre(TBPeriodiciteEVBrillantage, 0)
        .TempsMarcheEVBrillantage = ControleSiNombre(TBTempsMarcheEVBrillantage, 0)
    
        '--- électro-vanne d'eau de la ligne ---
        .ModeEVEauLigne = MemModeEVEauLigne
        
        '--- compresseur du pont 1 ---
        .ModeCompresseurP1 = MemModeCompresseurP1
        
        '--- éclairage du pont 1 ---
        .ModeEclairageP1 = MemModeEclairageP1
        
        '--- compresseur du pont 2 ---
        .ModeCompresseurP2 = MemModeCompresseurP2
    
        '--- éclairage du pont 2 ---
        .ModeEclairageP2 = MemModeEclairageP2
    
    End With

End Sub

'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Rôle      : Visualisation des différents états des annexes
' Entrées :
' Retours :
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Private Sub EtatsAnnexes()
    
    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next

    '--- constantee privées ---
    
    '--- déclaration ---
    Dim a As Integer
    Dim ClignotantDefaut As Integer
    Static CptAppels As Integer
    Dim CouleurFond As Long, CouleurPlan As Long
    Dim Texte As String

    '--- contrôle du clignotement ---
    CptAppels = IIf(CptAppels > 11, 1, CptAppels + 1)
    ClignotantDefaut = Choose(CptAppels, 1, 1, 0, 0, 1, 1, 0, 0, 1, 1, 0, 0)
    
    With TEtatsAnnexes
        
        '****************************************************************************************************************
        '                                 ELECTRO-VANNE D'AIR DANS LE BAIN DE BRILLANTAGE
        '****************************************************************************************************************
        
        '--- contrôle automatique des changements de l'électro-vanne d'air dans le bain de brillantage ---
        If .API_ChangementsEVBrillantage = True Then

            '--- position de l'indicateur ---
            PositionneIndicateurEVBrillantage

            '--- valeurs de la fenêtre ---
            LectureValeursFenetre

            '--- affectation ---
            .API_ChangementsEVBrillantage = False
        
        End If
    
        '--- états de l'électro-vanne d'aire dans le brillantage ---
        Select Case .EtatsEVBrillantage
            Case ETATS_EV_BRILLANTAGE.E_ARRET: CouleurFond = COULEURS.VERT_3: CouleurPlan = COULEURS.NOIR: Texte = "ARRET"
            Case ETATS_EV_BRILLANTAGE.E_MARCHE: CouleurFond = COULEURS.ORANGE_2: CouleurPlan = COULEURS.NOIR: Texte = "MARCHE"
            Case Else
        End Select
        With LEtats(INDEX_CHAMPS.IDX_CHAMP_ETATS_EV_BRILLANTAGE)
            If .Caption <> Texte Then
                SCadresEtats(INDEX_CHAMPS.IDX_CHAMP_ETATS_EV_BRILLANTAGE).BackColor = CouleurFond
                .ForeColor = CouleurPlan
                .Caption = Texte
            End If
        End With
        
        '****************************************************************************************************************
        '                                              ELECTRO-VANNE D'ARRIVEE D'EAU DE LA LIGNE
        '****************************************************************************************************************
        
        '--- contrôle automatique des changements de l'électro-vanne d'arrivée d'eau de la ligne ---
        If .API_ChangementsEVEauLigne = True Then

            '--- position de l'indicateur ---
            PositionneIndicateurEVEauLigne

            '--- valeurs de la fenêtre ---
            LectureValeursFenetre

            '--- affectation ---
            .API_ChangementsEVEauLigne = False
        
        End If
    
        '--- états  de l'électro-vanne d'arrivée d'eau de la ligne ---
        Select Case .EtatsEVEauLigne
            Case ETATS_EV_EAU_LIGNE.E_ARRET: CouleurFond = COULEURS.VERT_3: CouleurPlan = COULEURS.NOIR: Texte = "ARRET"
            Case ETATS_EV_EAU_LIGNE.E_MARCHE: CouleurFond = COULEURS.ORANGE_2: CouleurPlan = COULEURS.NOIR: Texte = "MARCHE"
            Case Else
        End Select
        With LEtats(INDEX_CHAMPS.IDX_CHAMP_ETATS_EV_EAU_LIGNE)
            If .Caption <> Texte Then
                SCadresEtats(INDEX_CHAMPS.IDX_CHAMP_ETATS_EV_EAU_LIGNE).BackColor = CouleurFond
                .ForeColor = CouleurPlan
                .Caption = Texte
            End If
        End With
        
        
        '****************************************************************************************************************
        '                                                           COMPRESSEUR DU PONT 1
        '****************************************************************************************************************
        
        '--- contrôle automatique des changements du compresseur du pont 1 ---
        If .API_ChangementsCompresseurP1 = True Then

            '--- position de l'indicateur ---
            PositionneIndicateurCompresseurP1

            '--- valeurs de la fenêtre ---
            LectureValeursFenetre

            '--- affectation ---
            .API_ChangementsCompresseurP1 = False
        
        End If
    
        '--- états du compresseur du pont 1 ---
        Select Case .EtatsCompresseurP1
            Case ETATS_COMPRESSEURS_PONTS.E_ARRET: CouleurFond = COULEURS.VERT_3: CouleurPlan = COULEURS.NOIR: Texte = "ARRET"
            Case ETATS_COMPRESSEURS_PONTS.E_MARCHE: CouleurFond = COULEURS.ORANGE_2: CouleurPlan = COULEURS.NOIR: Texte = "MARCHE"
            Case Else
        End Select
        With LEtats(INDEX_CHAMPS.IDX_CHAMP_ETATS_COMPRESSEUR_P1)
            If .Caption <> Texte Then
                SCadresEtats(INDEX_CHAMPS.IDX_CHAMP_ETATS_COMPRESSEUR_P1).BackColor = CouleurFond
                .ForeColor = CouleurPlan
                .Caption = Texte
            End If
        End With
        
        '****************************************************************************************************************
        '                                                          ECLAIRAGE DU PONT 1
        '****************************************************************************************************************
        
        '--- contrôle automatique des changements de l'éclairage du pont 1 ---
        If .API_ChangementsEclairageP1 = True Then

            '--- position de l'indicateur ---
            PositionneIndicateurEclairageP1

            '--- valeurs de la fenêtre ---
            LectureValeursFenetre

            '--- affectation ---
            .API_ChangementsEclairageP1 = False
        
        End If
    
        '--- états de l'éclairage du pont 1 ---
        Select Case .EtatsEclairageP1
            Case ETATS_ECLAIRAGE_PONTS.E_ARRET: CouleurFond = COULEURS.VERT_3: CouleurPlan = COULEURS.NOIR: Texte = "ARRET"
            Case ETATS_ECLAIRAGE_PONTS.E_MARCHE: CouleurFond = COULEURS.ORANGE_2: CouleurPlan = COULEURS.NOIR: Texte = "MARCHE"
            Case Else
        End Select
        With LEtats(INDEX_CHAMPS.IDX_CHAMP_ETATS_ECLAIRAGE_P1)
            If .Caption <> Texte Then
                SCadresEtats(INDEX_CHAMPS.IDX_CHAMP_ETATS_ECLAIRAGE_P1).BackColor = CouleurFond
                .ForeColor = CouleurPlan
                .Caption = Texte
            End If
        End With

        '****************************************************************************************************************
        '                                                           COMPRESSEUR DU PONT 2
        '****************************************************************************************************************
        
        '--- contrôle automatique des changements du compresseur du pont 2 ---
        If .API_ChangementsCompresseurP2 = True Then

            '--- position de l'indicateur ---
            PositionneIndicateurCompresseurP2

            '--- valeurs de la fenêtre ---
            LectureValeursFenetre

            '--- affectation ---
            .API_ChangementsCompresseurP2 = False
        
        End If
    
        '--- états du compresseur du pont 2 ---
        Select Case .EtatsCompresseurP2
            Case ETATS_COMPRESSEURS_PONTS.E_ARRET: CouleurFond = COULEURS.VERT_3: CouleurPlan = COULEURS.NOIR: Texte = "ARRET"
            Case ETATS_COMPRESSEURS_PONTS.E_MARCHE: CouleurFond = COULEURS.ORANGE_2: CouleurPlan = COULEURS.NOIR: Texte = "MARCHE"
            Case Else
        End Select
        With LEtats(INDEX_CHAMPS.IDX_CHAMP_ETATS_COMPRESSEUR_P2)
            If .Caption <> Texte Then
                SCadresEtats(INDEX_CHAMPS.IDX_CHAMP_ETATS_COMPRESSEUR_P2).BackColor = CouleurFond
                .ForeColor = CouleurPlan
                .Caption = Texte
            End If
        End With
        
        '****************************************************************************************************************
        '                                                           ECLAIRAGE DU PONT 2
        '****************************************************************************************************************
        
        '--- contrôle automatique des changements de l'éclairage du pont 2 ---
        If .API_ChangementsEclairageP2 = True Then

            '--- position de l'indicateur ---
            PositionneIndicateurEclairageP2

            '--- valeurs de la fenêtre ---
            LectureValeursFenetre

            '--- affectation ---
            .API_ChangementsEclairageP2 = False
        
        End If
    
        '--- états de l'éclairage du pont 2 ---
        Select Case .EtatsEclairageP2
            Case ETATS_ECLAIRAGE_PONTS.E_ARRET: CouleurFond = COULEURS.VERT_3: CouleurPlan = COULEURS.NOIR: Texte = "ARRET"
            Case ETATS_ECLAIRAGE_PONTS.E_MARCHE: CouleurFond = COULEURS.ORANGE_2: CouleurPlan = COULEURS.NOIR: Texte = "MARCHE"
            Case Else
        End Select
        With LEtats(INDEX_CHAMPS.IDX_CHAMP_ETATS_ECLAIRAGE_P2)
            If .Caption <> Texte Then
                SCadresEtats(INDEX_CHAMPS.IDX_CHAMP_ETATS_ECLAIRAGE_P2).BackColor = CouleurFond
                .ForeColor = CouleurPlan
                .Caption = Texte
            End If
        End With

        '****************************************************************************************************************
        '                                                           NIVEAUX DES RETENTIONS
        '****************************************************************************************************************
    
        For a = LBound(.TEtatsNiveauxRetentions()) To UBound(.TEtatsNiveauxRetentions())
    
            '--- contrôle automatique des changements de niveaux des rétentions ---
            If .TAPI_ChangementsNiveauxRetentions(a) = True Then
            
                '--- position du niveau concerné ---
                PositionneNiveauxRetentions (a)
            
                '--- affectation ---
                .TAPI_ChangementsNiveauxRetentions(a) = False
            
            End If
            
        Next a
    
    End With

End Sub

'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Rôle      : Positionne l'indicateur de l'électro-vanne d'air dans le bain du brillantage
' Entrées :
' Retours :
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Private Sub PositionneIndicateurEVBrillantage()
    
    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next

    '--- déclaration ---
    Dim a As Integer

    '--- effacement de tous les indicateurs ---
    For a = MODES_EV_BRILLANTAGE.M_ARRET To MODES_EV_BRILLANTAGE.M_MARCHE_FORCEE
        IIndicateurEVBrillantage(a).Visible = False
    Next a
    
    '--- affichage du bon indicateur ---
    With TEtatsAnnexes
        Select Case .ModeEVBrillantage
            Case MODES_EV_BRILLANTAGE.M_ARRET To MODES_EV_BRILLANTAGE.M_MARCHE_FORCEE: IIndicateurEVBrillantage(.ModeEVBrillantage).Visible = True
            Case Else
        End Select
   End With

End Sub

'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Rôle      : Positionne l'indicateur de l'éclairage du pont 1
' Entrées :
' Retours :
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Private Sub PositionneIndicateurEclairageP1()
    
    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next

    '--- déclaration ---
    Dim a As Integer

    '--- effacement de tous les indicateurs ---
    For a = MODES_ECLAIRAGE_PONTS.M_ARRET To MODES_ECLAIRAGE_PONTS.M_MARCHE
        IIndicateurEclairageP1(a).Visible = False
    Next a
    
    '--- affichage du bon indicateur ---
    With TEtatsAnnexes
        Select Case .ModeEclairageP1
            Case MODES_ECLAIRAGE_PONTS.M_ARRET To MODES_ECLAIRAGE_PONTS.M_MARCHE: IIndicateurEclairageP1(.ModeEclairageP1).Visible = True
            Case Else
        End Select
   End With

End Sub

'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Rôle      : Positionne l'indicateur du compresseur du pont 1
' Entrées :
' Retours :
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Private Sub PositionneIndicateurCompresseurP1()
    
    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next

    '--- déclaration ---
    Dim a As Integer

    '--- effacement de tous les indicateurs ---
    For a = MODES_COMPRESSEURS_PONTS.M_ARRET To MODES_COMPRESSEURS_PONTS.M_MARCHE
        IIndicateurCompresseurP1(a).Visible = False
    Next a
    
    '--- affichage du bon indicateur ---
    With TEtatsAnnexes
        Select Case .ModeCompresseurP1
            Case MODES_COMPRESSEURS_PONTS.M_ARRET To MODES_COMPRESSEURS_PONTS.M_MARCHE: IIndicateurCompresseurP1(.ModeCompresseurP1).Visible = True
            Case Else
        End Select
   End With

End Sub

'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Rôle      : Positionne l'indicateur de l'électro-vanne d'arrivée d'eau de la ligne
' Entrées :
' Retours :
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Private Sub PositionneIndicateurEVEauLigne()
    
    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next

    '--- déclaration ---
    Dim a As Integer

    '--- effacement de tous les indicateurs ---
    For a = MODES_EV_EAU_LIGNE.M_ARRET To MODES_EV_EAU_LIGNE.M_MARCHE
        IIndicateurEVEauLigne(a).Visible = False
    Next a
    
    '--- affichage du bon indicateur ---
    With TEtatsAnnexes
        Select Case .ModeEVEauLigne
            Case MODES_EV_EAU_LIGNE.M_ARRET To MODES_EV_EAU_LIGNE.M_MARCHE: IIndicateurEVEauLigne(.ModeEVEauLigne).Visible = True
            Case Else
        End Select
   End With

End Sub

'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Rôle      : Positionne l'indicateur du compresseur du pont 2
' Entrées :
' Retours :
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Private Sub PositionneIndicateurCompresseurP2()
    
    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next

    '--- déclaration ---
    Dim a As Integer

    '--- effacement de tous les indicateurs ---
    For a = MODES_COMPRESSEURS_PONTS.M_ARRET To MODES_COMPRESSEURS_PONTS.M_MARCHE
        IIndicateurCompresseurP2(a).Visible = False
    Next a
    
    '--- affichage du bon indicateur ---
    With TEtatsAnnexes
        Select Case .ModeCompresseurP2
            Case MODES_COMPRESSEURS_PONTS.M_ARRET To MODES_COMPRESSEURS_PONTS.M_MARCHE: IIndicateurCompresseurP2(.ModeCompresseurP2).Visible = True
            Case Else
        End Select
   End With

End Sub

'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Rôle      : Positionne l'indicateur de l'éclairage du pont 2
' Entrées :
' Retours :
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Private Sub PositionneIndicateurEclairageP2()
    
    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next

    '--- déclaration ---
    Dim a As Integer

    '--- effacement de tous les indicateurs ---
    For a = MODES_ECLAIRAGE_PONTS.M_ARRET To MODES_ECLAIRAGE_PONTS.M_MARCHE
        IIndicateurEclairageP2(a).Visible = False
    Next a
    
    '--- affichage du bon indicateur ---
    With TEtatsAnnexes
        Select Case .ModeEclairageP2
            Case MODES_ECLAIRAGE_PONTS.M_ARRET To MODES_ECLAIRAGE_PONTS.M_MARCHE: IIndicateurEclairageP2(.ModeEclairageP2).Visible = True
            Case Else
        End Select
   End With

End Sub

'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Rôle      : Positionne les niveaus des rétentions
' Entrées : NiveauRetentionConcerne -> fonction de l'énumération NIVEAUX_RETENTIONS
' Retours :
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Private Sub PositionneNiveauxRetentions(ByVal NiveauRetentionConcerne As NIVEAUX_RETENTIONS)
    
    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next

    '--- déclaration ---
    Dim a As Integer
    
    Dim NomDessinNiveau As String           'représente le nom du dessin du niveau
    
    Select Case TEtatsAnnexes.TEtatsNiveauxRetentions(NiveauRetentionConcerne)
    
        
        Case ETATS_NIVEAUX_RETENTIONS.ENR_BON
            '--- pas d'eau dans la rétention ---
            NomDessinNiveau = "niveau vert"
        
        Case ETATS_NIVEAUX_RETENTIONS.ENR_EN_DETECTION
            '--- en détection de liquide ---
            NomDessinNiveau = "niveau orange"
            
        Case ETATS_NIVEAUX_RETENTIONS.ENR_HAUT
            '--- niveau d'alarme de la rétention ---
            NomDessinNiveau = "niveau rouge"
        
        Case ETATS_NIVEAUX_RETENTIONS.ENR_INCONNU
            '--- niveau inconnu ---
            NomDessinNiveau = "niveau vide"

        Case Else
    End Select
    
    '--- affichage du bon dessin ---
    IIndicateursNiveaux(NiveauRetentionConcerne).Picture = ILIndicateursNiveaux.ListImages(NomDessinNiveau).Picture

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

'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Rôle      : Enregistre les valeurs dans l'automate
' Entrées :
' Retours : EnregistreVersAutomate -> FALSE = Incident de communication ou de transfert
'                                                             TRUE = Pas de problèmes
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Private Function EnregistreVersAutomate() As Boolean
    
    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next

    '--- constantes privées ---
    Const NOM_GROUPE As String = "ANNEXES"
    
    '--- déclaration ---
    Dim ValeurRetourneeAPI As Long        'valeur retournée par une fonction concernant le dialogue avec l'automate
    
    '--- blocage du timer ---
    TimerEtatsAnnexes.Enabled = False

    '--- affectation par défaut ---
    EnregistreVersAutomate = False
    
    '--- écriture dans l'API ---
    If PROGRAMME_AVEC_AUTOMATE = True Then
        
        ValeurRetourneeAPI = APIEcritureVariableNommee(NOM_GROUPE, ADR_API_MODE_EV_BRILLANTAGE, CStr(MemModeEVBrillantage))
        If ValeurRetourneeAPI <> 0 Then Exit Function
        
        ValeurRetourneeAPI = APIEcritureVariableNommee(NOM_GROUPE, ADR_API_PERIODICITE_MARCHE_EV_BRILLANTAGE, CStr(CInt(TBPeriodiciteEVBrillantage.Text) * 60))             'temps en minutes
        If ValeurRetourneeAPI <> 0 Then Exit Function
        
        ValeurRetourneeAPI = APIEcritureVariableNommee(NOM_GROUPE, ADR_API_TEMPS_MARCHE_EV_BRILLANTAGE, TBTempsMarcheEVBrillantage.Text)                                         'temps en secondes
        If ValeurRetourneeAPI <> 0 Then Exit Function
        
        ValeurRetourneeAPI = APIEcritureVariableNommee(NOM_GROUPE, ADR_API_MODE_EV_EAU_LIGNE, CStr(MemModeEVEauLigne))
        If ValeurRetourneeAPI <> 0 Then Exit Function
        
        ValeurRetourneeAPI = APIEcritureVariableNommee(NOM_GROUPE, ADR_API_MODE_COMPRESSEUR_P1, CStr(MemModeCompresseurP1))
        If ValeurRetourneeAPI <> 0 Then Exit Function
        
        ValeurRetourneeAPI = APIEcritureVariableNommee(NOM_GROUPE, ADR_API_MODE_ECLAIRAGE_P1, CStr(MemModeEclairageP1))
        If ValeurRetourneeAPI <> 0 Then Exit Function
        
        ValeurRetourneeAPI = APIEcritureVariableNommee(NOM_GROUPE, ADR_API_MODE_COMPRESSEUR_P2, CStr(MemModeCompresseurP2))
        If ValeurRetourneeAPI <> 0 Then Exit Function
        
        ValeurRetourneeAPI = APIEcritureVariableNommee(NOM_GROUPE, ADR_API_MODE_ECLAIRAGE_P2, CStr(MemModeEclairageP2))
        If ValeurRetourneeAPI <> 0 Then Exit Function
 
    End If
    
    '--- affectation ---
    EnregistreVersAutomate = True

    '--- déblocage du timer ---
    TimerEtatsAnnexes.Enabled = True
    
End Function

'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Rôle      : Initialise la fenêtre (chargement ou en vue de la rendre visible)
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Sub InitialisationFenetre()
    
    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next

    '--- constantes privées ---
    
    '--- déclaration ---
    Dim a As Integer

    '--- affectation ---

    '--- divers sur la fenêtre ---
    With Me
        .Caption = TITRE_FENETRE
        .WindowState = vbMaximized
    End With
    
    '--- couleurs des fonds ---
    PBDeplacementFenetre(ZONES_DEPLACEMENT_FENETRE.Z_FILLE).Picture = ImgFondVert1
    
    PBElementsLigne.Picture = ImgFondOrange1
    PBPont1.Picture = ImgFondOrange1
    PBPont2.Picture = ImgFondOrange1
    
    PBGestionEVBrillantage.Picture = ImgFondGris1
    PBGestionEVEauLigne.Picture = ImgFondGris1
    
    PBGestionCompresseurP1.Picture = ImgFondGris1
    PBGestionEclairageP1.Picture = ImgFondGris1
    
    PBGestionCompresseurP2.Picture = ImgFondGris1
    PBGestionEclairageP2.Picture = ImgFondGris1
    
    PBBoutons.Picture = ImgFondDesBoutons
    
    '--- valeurs de la fenêtre ---
    LectureValeursFenetre
                
    '--- position de la totalité des indicateurs ---
    PositionneIndicateurEVBrillantage
    
    PositionneIndicateurEVEauLigne
    
    PositionneIndicateurCompresseurP1
    PositionneIndicateurEclairageP1
    
    PositionneIndicateurCompresseurP2
    PositionneIndicateurEclairageP2
    
    '--- visualisation des différents états des annexes ---
    EtatsAnnexes

    '--- lancement du timer ---
    TimerEtatsAnnexes.Enabled = True

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
    With TimerEtatsAnnexes
        .Enabled = False
        .Interval = 0
    End With
    
    '--- déchargement de la fenêtre ---
    Me.Visible = False
    Unload Me
    Set OccFAnnexes = Nothing
    
End Sub

Private Sub TBPeriodiciteEVBrillantage_Change()
    On Error Resume Next
    If InterdireEvenements = False Then
        CBValider.Enabled = True
    End If
End Sub

Private Sub TBPeriodiciteEVBrillantage_GotFocus()
    On Error Resume Next
    With TBPeriodiciteEVBrillantage
        .SelStart = 0          'met en surbrillance la sélection saisie
        .SelLength = Len(.Text)
    End With
End Sub

Private Sub TBPeriodiciteEVBrillantage_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error Resume Next
    FiltreToucheFonction KeyCode, Shift
End Sub

Private Sub TBPeriodiciteEVBrillantage_KeyPress(KeyAscii As Integer)
    On Error Resume Next
    FiltreToucheASCII KeyAscii, DONNEES.D_NBR_NATURELS, 3
End Sub

Private Sub TBPeriodiciteEVBrillantage_LostFocus()
    On Error Resume Next
    InterdireEvenements = True
    AffichageNombre TBPeriodiciteEVBrillantage, TBPeriodiciteEVBrillantage.Text, FORMAT_NATURELS_3_CHIFFRES
    InterdireEvenements = False
End Sub

Private Sub TBTempsMarcheEVBrillantage_Change()
    On Error Resume Next
    If InterdireEvenements = False Then
        CBValider.Enabled = True
    End If
End Sub

Private Sub TBTempsMarcheEVBrillantage_GotFocus()
    On Error Resume Next
    With TBTempsMarcheEVBrillantage
        .SelStart = 0          'met en surbrillance la sélection saisie
        .SelLength = Len(.Text)
    End With
End Sub

Private Sub TBTempsMarcheEVBrillantage_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error Resume Next
    FiltreToucheFonction KeyCode, Shift
End Sub

Private Sub TBTempsMarcheEVBrillantage_KeyPress(KeyAscii As Integer)
    On Error Resume Next
    FiltreToucheASCII KeyAscii, DONNEES.D_NBR_NATURELS, 4
End Sub

Private Sub TBTempsMarcheEVBrillantage_LostFocus()
    On Error Resume Next
    InterdireEvenements = True
    AffichageNombre TBTempsMarcheEVBrillantage, TBTempsMarcheEVBrillantage.Text, FORMAT_NATURELS_4_CHIFFRES
    InterdireEvenements = False
End Sub

Private Sub TimerEtatsAnnexes_Timer()
    
    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next
    
    '--- appel de la routine ---
    TimerEtatsAnnexes.Enabled = False
    EtatsAnnexes
    TimerEtatsAnnexes.Enabled = True
    
    '--- bip de passage dans la routine UNIQUEMENT POUR LES TESTS ---
    If PROGRAMME_AVEC_AUTOMATE = False Then Beep

End Sub

Private Sub VSDeplacementFENETRE_Change()
    On Error Resume Next
    PBDeplacementFenetre(ZONES_DEPLACEMENT_FENETRE.Z_FILLE).Top = -VSDeplacementFenetre.Value
End Sub


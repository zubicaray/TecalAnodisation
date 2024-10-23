VERSION 5.00
Begin VB.Form FGestionCuves 
   BackColor       =   &H00C0C0C0&
   ClientHeight    =   13050
   ClientLeft      =   -6945
   ClientTop       =   3180
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
   Icon            =   "FGestionCuves.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   13050
   ScaleWidth      =   13440
   ShowInTaskbar   =   0   'False
   WindowState     =   2  'Maximized
   Begin VB.PictureBox PBRenseignementsFenetre 
      Align           =   1  'Align Top
      BackColor       =   &H00FF0000&
      ForeColor       =   &H0000FFFF&
      Height          =   375
      Left            =   0
      Picture         =   "FGestionCuves.frx":014A
      ScaleHeight     =   315
      ScaleWidth      =   13380
      TabIndex        =   85
      Top             =   0
      Width           =   13440
      Begin VB.Label LRenseignementsFenetre 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
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
         Height          =   225
         Left            =   240
         TabIndex        =   86
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
      ScaleWidth      =   13380
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   11955
      Width           =   13440
      Begin VB.Timer TimerEtatsCuve 
         Enabled         =   0   'False
         Interval        =   500
         Left            =   1560
         Top             =   120
      End
      Begin VB.CommandButton CBQuitter 
         BackColor       =   &H00FFFFFF&
         Cancel          =   -1  'True
         Caption         =   "&Quitter"
         DownPicture     =   "FGestionCuves.frx":24A8C
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
         Picture         =   "FGestionCuves.frx":2518E
         Style           =   1  'Graphical
         TabIndex        =   9
         ToolTipText     =   " Quitter cette fenêtre "
         Top             =   120
         UseMaskColor    =   -1  'True
         Width           =   1515
      End
      Begin VB.CommandButton CBValider 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Valider"
         DownPicture     =   "FGestionCuves.frx":25890
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
         Picture         =   "FGestionCuves.frx":25F92
         Style           =   1  'Graphical
         TabIndex        =   8
         ToolTipText     =   " Valider l'enregistrement "
         Top             =   120
         UseMaskColor    =   -1  'True
         Width           =   1515
      End
      Begin VB.CommandButton CBAnnuler 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Annuler"
         DownPicture     =   "FGestionCuves.frx":26694
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
         Picture         =   "FGestionCuves.frx":26D96
         Style           =   1  'Graphical
         TabIndex        =   7
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
         TabIndex        =   3
         Top             =   0
         Visible         =   0   'False
         Width           =   1215
         Begin VB.HScrollBar HSDeplacementFenetre 
            Height          =   255
            LargeChange     =   300
            Left            =   0
            SmallChange     =   100
            TabIndex        =   6
            Top             =   720
            Width           =   915
         End
         Begin VB.VScrollBar VSDeplacementFenetre 
            Height          =   975
            LargeChange     =   300
            Left            =   900
            SmallChange     =   100
            TabIndex        =   5
            Top             =   0
            Width           =   255
         End
         Begin VB.CommandButton CBAgrandirFenetre 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Agrandir"
            DownPicture     =   "FGestionCuves.frx":27498
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
            Picture         =   "FGestionCuves.frx":27642
            Style           =   1  'Graphical
            TabIndex        =   4
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
         Left            =   2220
         Top             =   180
         Visible         =   0   'False
         Width           =   780
      End
   End
   Begin VB.PictureBox PBDeplacementFenetre 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   13455
      Index           =   0
      Left            =   0
      ScaleHeight     =   13455
      ScaleWidth      =   13440
      TabIndex        =   0
      Top             =   375
      Width           =   13440
      Begin VB.PictureBox PBDeplacementFenetre 
         Height          =   12975
         Index           =   1
         Left            =   0
         ScaleHeight     =   12915
         ScaleWidth      =   26115
         TabIndex        =   1
         Top             =   0
         Width           =   26175
         Begin VB.PictureBox PBTemperature 
            BackColor       =   &H00C0E0FF&
            Height          =   855
            Left            =   9840
            ScaleHeight     =   795
            ScaleWidth      =   5175
            TabIndex        =   72
            Top             =   4680
            Width           =   5235
            Begin VB.Label LEtats 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               BackStyle       =   0  'Transparent
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   14.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   315
               Index           =   12
               Left            =   240
               TabIndex        =   73
               Top             =   240
               Width           =   4695
            End
            Begin VB.Shape SCadresEtats 
               BackColor       =   &H00E0E0E0&
               BackStyle       =   1  'Opaque
               Height          =   435
               Index           =   12
               Left            =   180
               Top             =   180
               Width           =   4815
            End
         End
         Begin VB.PictureBox PBEtatsElements 
            BackColor       =   &H00C0E0FF&
            Height          =   6195
            Left            =   9840
            ScaleHeight     =   6135
            ScaleWidth      =   5175
            TabIndex        =   48
            Top             =   6240
            Width           =   5235
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
               Left            =   2220
               TabIndex        =   70
               Top             =   240
               Width           =   2715
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
               Left            =   2220
               TabIndex        =   69
               Top             =   600
               Width           =   2715
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
               Left            =   2220
               TabIndex        =   68
               Top             =   960
               Width           =   2715
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
               Left            =   2220
               TabIndex        =   67
               Top             =   1320
               Width           =   2715
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
               Left            =   2220
               TabIndex        =   66
               Top             =   1680
               Width           =   2715
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
               Left            =   2220
               TabIndex        =   65
               Top             =   2040
               Width           =   2715
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
               Index           =   7
               Left            =   2220
               TabIndex        =   64
               Top             =   2400
               Width           =   2715
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
               Index           =   8
               Left            =   2220
               TabIndex        =   63
               Top             =   2760
               Width           =   2715
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
               Index           =   9
               Left            =   2220
               TabIndex        =   62
               Top             =   3120
               Width           =   2715
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
               Index           =   10
               Left            =   2220
               TabIndex        =   61
               Top             =   3480
               Width           =   2715
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
               Index           =   11
               Left            =   2220
               TabIndex        =   60
               Top             =   3840
               Width           =   2715
            End
            Begin VB.Label LLibellesEtats 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
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
               Left            =   240
               TabIndex        =   59
               Top             =   240
               Width           =   1875
            End
            Begin VB.Label LLibellesEtats 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
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
               Left            =   240
               TabIndex        =   58
               Top             =   600
               Width           =   1875
            End
            Begin VB.Label LLibellesEtats 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
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
               Left            =   240
               TabIndex        =   57
               Top             =   960
               Width           =   1875
            End
            Begin VB.Label LLibellesEtats 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
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
               Left            =   240
               TabIndex        =   56
               Top             =   1320
               Width           =   1875
            End
            Begin VB.Label LLibellesEtats 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
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
               Left            =   240
               TabIndex        =   55
               Top             =   1680
               Width           =   1875
            End
            Begin VB.Label LLibellesEtats 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
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
               Left            =   240
               TabIndex        =   54
               Top             =   2040
               Width           =   1875
            End
            Begin VB.Label LLibellesEtats 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
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
               Left            =   240
               TabIndex        =   53
               Top             =   2400
               Width           =   1875
            End
            Begin VB.Label LLibellesEtats 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
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
               Left            =   240
               TabIndex        =   52
               Top             =   2760
               Width           =   1875
            End
            Begin VB.Label LLibellesEtats 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
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
               Left            =   240
               TabIndex        =   51
               Top             =   3120
               Width           =   1875
            End
            Begin VB.Label LLibellesEtats 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
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
               Index           =   10
               Left            =   240
               TabIndex        =   50
               Top             =   3480
               Width           =   1875
            End
            Begin VB.Label LLibellesEtats 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
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
               Index           =   11
               Left            =   240
               TabIndex        =   49
               Top             =   3840
               Width           =   1875
            End
            Begin VB.Shape SCadresLibellesEtats 
               BackColor       =   &H00FFFFFF&
               BackStyle       =   1  'Opaque
               Height          =   375
               Index           =   11
               Left            =   180
               Top             =   3780
               Width           =   1995
            End
            Begin VB.Shape SCadresLibellesEtats 
               BackColor       =   &H00FFFFFF&
               BackStyle       =   1  'Opaque
               Height          =   375
               Index           =   10
               Left            =   180
               Top             =   3420
               Width           =   1995
            End
            Begin VB.Shape SCadresLibellesEtats 
               BackColor       =   &H00FFFFFF&
               BackStyle       =   1  'Opaque
               Height          =   375
               Index           =   9
               Left            =   180
               Top             =   3060
               Width           =   1995
            End
            Begin VB.Shape SCadresLibellesEtats 
               BackColor       =   &H00FFFFFF&
               BackStyle       =   1  'Opaque
               Height          =   375
               Index           =   8
               Left            =   180
               Top             =   2700
               Width           =   1995
            End
            Begin VB.Shape SCadresLibellesEtats 
               BackColor       =   &H00FFFFFF&
               BackStyle       =   1  'Opaque
               Height          =   375
               Index           =   7
               Left            =   180
               Top             =   2340
               Width           =   1995
            End
            Begin VB.Shape SCadresLibellesEtats 
               BackColor       =   &H00FFFFFF&
               BackStyle       =   1  'Opaque
               Height          =   375
               Index           =   6
               Left            =   180
               Top             =   1980
               Width           =   1995
            End
            Begin VB.Shape SCadresLibellesEtats 
               BackColor       =   &H00FFFFFF&
               BackStyle       =   1  'Opaque
               Height          =   375
               Index           =   5
               Left            =   180
               Top             =   1620
               Width           =   1995
            End
            Begin VB.Shape SCadresLibellesEtats 
               BackColor       =   &H00FFFFFF&
               BackStyle       =   1  'Opaque
               Height          =   375
               Index           =   4
               Left            =   180
               Top             =   1260
               Width           =   1995
            End
            Begin VB.Shape SCadresLibellesEtats 
               BackColor       =   &H00FFFFFF&
               BackStyle       =   1  'Opaque
               Height          =   375
               Index           =   3
               Left            =   180
               Top             =   900
               Width           =   1995
            End
            Begin VB.Shape SCadresLibellesEtats 
               BackColor       =   &H00FFFFFF&
               BackStyle       =   1  'Opaque
               Height          =   375
               Index           =   2
               Left            =   180
               Top             =   540
               Width           =   1995
            End
            Begin VB.Shape SCadresLibellesEtats 
               BackColor       =   &H00FFFFFF&
               BackStyle       =   1  'Opaque
               Height          =   375
               Index           =   1
               Left            =   180
               Top             =   180
               Width           =   1995
            End
            Begin VB.Shape SCadresEtats 
               BackColor       =   &H00E0E0E0&
               BackStyle       =   1  'Opaque
               Height          =   375
               Index           =   11
               Left            =   2160
               Top             =   3780
               Width           =   2835
            End
            Begin VB.Shape SCadresEtats 
               BackColor       =   &H00E0E0E0&
               BackStyle       =   1  'Opaque
               Height          =   375
               Index           =   10
               Left            =   2160
               Top             =   3420
               Width           =   2835
            End
            Begin VB.Shape SCadresEtats 
               BackColor       =   &H00E0E0E0&
               BackStyle       =   1  'Opaque
               Height          =   375
               Index           =   9
               Left            =   2160
               Top             =   3060
               Width           =   2835
            End
            Begin VB.Shape SCadresEtats 
               BackColor       =   &H00E0E0E0&
               BackStyle       =   1  'Opaque
               Height          =   375
               Index           =   8
               Left            =   2160
               Top             =   2700
               Width           =   2835
            End
            Begin VB.Shape SCadresEtats 
               BackColor       =   &H00E0E0E0&
               BackStyle       =   1  'Opaque
               Height          =   375
               Index           =   7
               Left            =   2160
               Top             =   2340
               Width           =   2835
            End
            Begin VB.Shape SCadresEtats 
               BackColor       =   &H00E0E0E0&
               BackStyle       =   1  'Opaque
               Height          =   375
               Index           =   6
               Left            =   2160
               Top             =   1980
               Width           =   2835
            End
            Begin VB.Shape SCadresEtats 
               BackColor       =   &H00E0E0E0&
               BackStyle       =   1  'Opaque
               Height          =   375
               Index           =   5
               Left            =   2160
               Top             =   1620
               Width           =   2835
            End
            Begin VB.Shape SCadresEtats 
               BackColor       =   &H00E0E0E0&
               BackStyle       =   1  'Opaque
               Height          =   375
               Index           =   4
               Left            =   2160
               Top             =   1260
               Width           =   2835
            End
            Begin VB.Shape SCadresEtats 
               BackColor       =   &H00E0E0E0&
               BackStyle       =   1  'Opaque
               Height          =   375
               Index           =   3
               Left            =   2160
               Top             =   900
               Width           =   2835
            End
            Begin VB.Shape SCadresEtats 
               BackColor       =   &H00E0E0E0&
               BackStyle       =   1  'Opaque
               Height          =   375
               Index           =   2
               Left            =   2160
               Top             =   540
               Width           =   2835
            End
            Begin VB.Shape SCadresEtats 
               BackColor       =   &H00E0E0E0&
               BackStyle       =   1  'Opaque
               Height          =   375
               Index           =   1
               Left            =   2160
               Top             =   180
               Width           =   2835
            End
         End
         Begin VB.PictureBox PBProgrammationEnCours 
            BackColor       =   &H00C0E0FF&
            Height          =   3315
            Left            =   9840
            ScaleHeight     =   3255
            ScaleWidth      =   10635
            TabIndex        =   33
            Top             =   660
            Width           =   10695
            Begin VB.Frame FLegendeChauffage 
               BackColor       =   &H00C0E0FF&
               Caption         =   " Légende du chauffage "
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   975
               Left            =   4080
               TabIndex        =   41
               Top             =   1920
               Width           =   6195
               Begin VB.PictureBox PBModesChauffage 
                  Appearance      =   0  'Flat
                  BackColor       =   &H00FFFFFF&
                  ForeColor       =   &H000080FF&
                  Height          =   375
                  Index           =   2
                  Left            =   3180
                  ScaleHeight     =   345
                  ScaleWidth      =   465
                  TabIndex        =   44
                  TabStop         =   0   'False
                  Top             =   360
                  Width           =   495
               End
               Begin VB.PictureBox PBModesChauffage 
                  Appearance      =   0  'Flat
                  BackColor       =   &H00FFFFFF&
                  ForeColor       =   &H00FFFF00&
                  Height          =   375
                  Index           =   1
                  Left            =   1680
                  ScaleHeight     =   345
                  ScaleWidth      =   465
                  TabIndex        =   43
                  TabStop         =   0   'False
                  Top             =   360
                  Width           =   495
               End
               Begin VB.PictureBox PBModesChauffage 
                  Appearance      =   0  'Flat
                  BackColor       =   &H00FFFFFF&
                  ForeColor       =   &H00FFFFFF&
                  Height          =   375
                  Index           =   0
                  Left            =   240
                  ScaleHeight     =   345
                  ScaleWidth      =   465
                  TabIndex        =   42
                  TabStop         =   0   'False
                  Top             =   360
                  Width           =   495
               End
               Begin VB.Label LLegendes 
                  BackColor       =   &H00C0E0FF&
                  Caption         =   "Production"
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
                  Left            =   3840
                  TabIndex        =   47
                  Top             =   420
                  Width           =   1335
               End
               Begin VB.Label LLegendes 
                  BackColor       =   &H00C0E0FF&
                  Caption         =   "Veille"
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
                  Left            =   2340
                  TabIndex        =   46
                  Top             =   420
                  Width           =   615
               End
               Begin VB.Label LLegendes 
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
                  Height          =   255
                  Index           =   2
                  Left            =   900
                  TabIndex        =   45
                  Top             =   420
                  Width           =   495
               End
            End
            Begin VB.Frame FLegendePompe 
               BackColor       =   &H00C0E0FF&
               Caption         =   " Légende de la pompe "
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   975
               Left            =   360
               TabIndex        =   36
               Top             =   1920
               Width           =   3375
               Begin VB.PictureBox PBModesPompe 
                  Appearance      =   0  'Flat
                  BackColor       =   &H00FFFFFF&
                  ForeColor       =   &H000080FF&
                  Height          =   375
                  Index           =   1
                  Left            =   1680
                  ScaleHeight     =   345
                  ScaleWidth      =   465
                  TabIndex        =   38
                  TabStop         =   0   'False
                  Top             =   360
                  Width           =   495
               End
               Begin VB.PictureBox PBModesPompe 
                  Appearance      =   0  'Flat
                  BackColor       =   &H00FFFFFF&
                  ForeColor       =   &H00FFFFFF&
                  Height          =   375
                  Index           =   0
                  Left            =   240
                  ScaleHeight     =   345
                  ScaleWidth      =   465
                  TabIndex        =   37
                  TabStop         =   0   'False
                  Top             =   360
                  Width           =   495
               End
               Begin VB.Label LLegendes 
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
                  Height          =   255
                  Index           =   0
                  Left            =   900
                  TabIndex        =   40
                  Top             =   420
                  Width           =   555
               End
               Begin VB.Label LLegendes 
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
                  Height          =   255
                  Index           =   1
                  Left            =   2340
                  TabIndex        =   39
                  Top             =   420
                  Width           =   795
               End
            End
            Begin VB.PictureBox PBEchelle24H 
               Appearance      =   0  'Flat
               AutoSize        =   -1  'True
               BackColor       =   &H80000005&
               BorderStyle     =   0  'None
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   945
               Left            =   360
               Picture         =   "FGestionCuves.frx":277EC
               ScaleHeight     =   63
               ScaleMode       =   3  'Pixel
               ScaleWidth      =   661
               TabIndex        =   34
               TabStop         =   0   'False
               Top             =   360
               Width           =   9915
            End
            Begin VB.Label LCycleTravailEnCours 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H00C0FFFF&
               BorderStyle     =   1  'Fixed Single
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   12
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   315
               Left            =   360
               TabIndex        =   35
               Top             =   1305
               Width           =   9915
            End
         End
         Begin VB.PictureBox PBSelectionNomsCuves 
            Height          =   11775
            Left            =   300
            ScaleHeight     =   11715
            ScaleWidth      =   9135
            TabIndex        =   10
            Top             =   660
            Width           =   9195
            Begin VB.OptionButton OBNomsCuves 
               BackColor       =   &H00C00000&
               Caption         =   "C37"
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
               Index           =   19
               Left            =   360
               Style           =   1  'Graphical
               TabIndex        =   19
               Top             =   10800
               Visible         =   0   'False
               Width           =   8115
            End
            Begin VB.VScrollBar VSChangementCuve 
               Height          =   5160
               Left            =   8820
               Max             =   14
               Min             =   1
               TabIndex        =   30
               Top             =   840
               Value           =   1
               Width           =   315
            End
            Begin VB.OptionButton OBNomsCuves 
               BackColor       =   &H00FFFFFF&
               Caption         =   "C04"
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
               TabIndex        =   29
               Top             =   2040
               Width           =   8115
            End
            Begin VB.OptionButton OBNomsCuves 
               BackColor       =   &H00FFFFFF&
               Caption         =   "SAT"
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
               TabIndex        =   28
               Top             =   1440
               Width           =   8115
            End
            Begin VB.OptionButton OBNomsCuves 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               Caption         =   "DEC"
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
               TabIndex        =   27
               Top             =   900
               Width           =   8115
            End
            Begin VB.OptionButton OBNomsCuves 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               Caption         =   "C00"
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
               TabIndex        =   26
               Top             =   360
               UseMaskColor    =   -1  'True
               Width           =   8115
            End
            Begin VB.OptionButton OBNomsCuves 
               BackColor       =   &H00008000&
               Caption         =   "C16"
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
               TabIndex        =   25
               Top             =   6180
               Width           =   8115
            End
            Begin VB.OptionButton OBNomsCuves 
               BackColor       =   &H00008000&
               Caption         =   "C15"
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
               TabIndex        =   24
               Top             =   5640
               Width           =   8115
            End
            Begin VB.OptionButton OBNomsCuves 
               BackColor       =   &H00008000&
               Caption         =   "C14"
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
               TabIndex        =   23
               Top             =   5100
               Width           =   8115
            End
            Begin VB.OptionButton OBNomsCuves 
               BackColor       =   &H00008000&
               Caption         =   "C13"
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
               Index           =   8
               Left            =   360
               Style           =   1  'Graphical
               TabIndex        =   22
               Top             =   4560
               Width           =   8115
            End
            Begin VB.OptionButton OBNomsCuves 
               BackColor       =   &H00FFFFFF&
               Caption         =   "C07"
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
               Left            =   360
               Style           =   1  'Graphical
               TabIndex        =   21
               Top             =   3840
               Width           =   8115
            End
            Begin VB.OptionButton OBNomsCuves 
               BackColor       =   &H00FFFFFF&
               Caption         =   "C06"
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
               TabIndex        =   20
               Top             =   3300
               Width           =   8115
            End
            Begin VB.OptionButton OBNomsCuves 
               BackColor       =   &H00FFFFFF&
               Caption         =   "C05"
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
               TabIndex        =   18
               Top             =   2760
               Width           =   8115
            End
            Begin VB.OptionButton OBNomsCuves 
               BackColor       =   &H00C00000&
               Caption         =   "C33"
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
               TabIndex        =   17
               Top             =   10260
               Visible         =   0   'False
               Width           =   8115
            End
            Begin VB.OptionButton OBNomsCuves 
               BackColor       =   &H00C00000&
               Caption         =   "C32"
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
               TabIndex        =   16
               Top             =   9600
               Visible         =   0   'False
               Width           =   8115
            End
            Begin VB.OptionButton OBNomsCuves 
               BackColor       =   &H00C00000&
               Caption         =   "C31"
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
               TabIndex        =   15
               Top             =   9060
               Visible         =   0   'False
               Width           =   8115
            End
            Begin VB.OptionButton OBNomsCuves 
               BackColor       =   &H00C00000&
               Caption         =   "C28"
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
               TabIndex        =   14
               Top             =   8520
               Visible         =   0   'False
               Width           =   8115
            End
            Begin VB.OptionButton OBNomsCuves 
               BackColor       =   &H00C00000&
               Caption         =   "C27"
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
               TabIndex        =   13
               Top             =   7980
               Visible         =   0   'False
               Width           =   8115
            End
            Begin VB.OptionButton OBNomsCuves 
               BackColor       =   &H00C00000&
               Caption         =   "C22"
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
               TabIndex        =   12
               Top             =   7440
               Visible         =   0   'False
               Width           =   8115
            End
            Begin VB.OptionButton OBNomsCuves 
               BackColor       =   &H00C00000&
               Caption         =   "C19"
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
               TabIndex        =   11
               Top             =   6900
               Visible         =   0   'False
               Width           =   8115
            End
         End
         Begin VB.PictureBox PBModesGestion 
            BackColor       =   &H00C0E0FF&
            Height          =   6195
            Left            =   15420
            ScaleHeight     =   6135
            ScaleWidth      =   5055
            TabIndex        =   75
            Top             =   6240
            Visible         =   0   'False
            Width           =   5115
            Begin VB.PictureBox PBZonePompe 
               Appearance      =   0  'Flat
               BackColor       =   &H00C0E0FF&
               ForeColor       =   &H80000008&
               Height          =   1575
               Left            =   180
               ScaleHeight     =   1545
               ScaleWidth      =   3405
               TabIndex        =   81
               TabStop         =   0   'False
               Top             =   180
               Visible         =   0   'False
               Width           =   3435
               Begin VB.OptionButton OBModesPompe 
                  BackColor       =   &H00C0E0FF&
                  Caption         =   "Forcer l'ARRET"
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
                  TabIndex        =   84
                  Top             =   780
                  Width           =   2775
               End
               Begin VB.OptionButton OBModesPompe 
                  BackColor       =   &H00C0E0FF&
                  Caption         =   "Forcer la MARCHE"
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
                  Index           =   3
                  Left            =   540
                  TabIndex        =   83
                  Top             =   1140
                  Width           =   2775
               End
               Begin VB.OptionButton OBModesPompe 
                  BackColor       =   &H00C0E0FF&
                  Caption         =   "En AUTOMATIQUE"
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
                  TabIndex        =   82
                  Top             =   435
                  Width           =   2775
               End
               Begin VB.Label LLibelles 
                  Alignment       =   2  'Center
                  Appearance      =   0  'Flat
                  BackColor       =   &H00808000&
                  BorderStyle     =   1  'Fixed Single
                  Caption         =   "Gestion de la POMPE"
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
                  Height          =   285
                  Index           =   7
                  Left            =   -15
                  TabIndex        =   88
                  Top             =   -15
                  Width           =   3435
               End
               Begin VB.Image IIndicateur 
                  Height          =   270
                  Index           =   0
                  Left            =   180
                  Picture         =   "FGestionCuves.frx":4606E
                  Top             =   420
                  Width           =   255
               End
               Begin VB.Image IIndicateur 
                  Height          =   270
                  Index           =   4
                  Left            =   180
                  Picture         =   "FGestionCuves.frx":46458
                  Top             =   780
                  Width           =   255
               End
               Begin VB.Image IIndicateur 
                  Height          =   270
                  Index           =   5
                  Left            =   180
                  Picture         =   "FGestionCuves.frx":46842
                  Top             =   1140
                  Width           =   255
               End
            End
            Begin VB.PictureBox PBZoneCouvercles 
               Appearance      =   0  'Flat
               BackColor       =   &H00C0E0FF&
               ForeColor       =   &H80000008&
               Height          =   1575
               Left            =   180
               ScaleHeight     =   1545
               ScaleWidth      =   3405
               TabIndex        =   77
               TabStop         =   0   'False
               Top             =   1920
               Visible         =   0   'False
               Width           =   3435
               Begin VB.OptionButton OBModesCouvercles 
                  BackColor       =   &H00C0E0FF&
                  Caption         =   "AUTOMATIQUE"
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
                  TabIndex        =   80
                  Top             =   435
                  Width           =   2775
               End
               Begin VB.OptionButton OBModesCouvercles 
                  BackColor       =   &H00C0E0FF&
                  Caption         =   "Forcer l'OUVERTURE"
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
                  Index           =   3
                  Left            =   540
                  TabIndex        =   79
                  Top             =   1140
                  Width           =   2775
               End
               Begin VB.OptionButton OBModesCouvercles 
                  BackColor       =   &H00C0E0FF&
                  Caption         =   "Forcer la FERMETURE"
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
                  TabIndex        =   78
                  Top             =   780
                  Width           =   2775
               End
               Begin VB.Label LLibelles 
                  Alignment       =   2  'Center
                  Appearance      =   0  'Flat
                  BackColor       =   &H00808000&
                  BorderStyle     =   1  'Fixed Single
                  Caption         =   "Gestion des COUVERCLES"
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
                  Height          =   285
                  Index           =   5
                  Left            =   -15
                  TabIndex        =   87
                  Top             =   -15
                  Width           =   3435
               End
               Begin VB.Image IIndicateur 
                  Height          =   270
                  Index           =   3
                  Left            =   180
                  Picture         =   "FGestionCuves.frx":46C2C
                  Top             =   1140
                  Width           =   255
               End
               Begin VB.Image IIndicateur 
                  Height          =   270
                  Index           =   2
                  Left            =   180
                  Picture         =   "FGestionCuves.frx":47016
                  Top             =   780
                  Width           =   255
               End
               Begin VB.Image IIndicateur 
                  Height          =   270
                  Index           =   1
                  Left            =   180
                  Picture         =   "FGestionCuves.frx":47400
                  Top             =   420
                  Width           =   255
               End
            End
         End
         Begin VB.Label LLibelles 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFF80&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "MODES de GESTION"
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
            Height          =   375
            Index           =   4
            Left            =   15420
            TabIndex        =   76
            Top             =   5880
            Visible         =   0   'False
            Width           =   5115
         End
         Begin VB.Label LLibelles 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFF80&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "TEMPERATURE"
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
            Height          =   375
            Index           =   3
            Left            =   9840
            TabIndex        =   74
            Top             =   4320
            Width           =   5235
         End
         Begin VB.Label LLibelles 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFF80&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "ETATS des ELEMENTS"
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
            Height          =   375
            Index           =   1
            Left            =   9840
            TabIndex        =   71
            Top             =   5880
            Width           =   5235
         End
         Begin VB.Label LLibelles 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFF80&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "PROGRAMMATION EN COURS"
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
            Height          =   375
            Index           =   0
            Left            =   9840
            TabIndex        =   32
            Top             =   300
            Width           =   10695
         End
         Begin VB.Label LLibelles 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFF80&
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
            ForeColor       =   &H00000000&
            Height          =   375
            Index           =   2
            Left            =   300
            TabIndex        =   31
            Top             =   300
            Width           =   9195
         End
      End
   End
End
Attribute VB_Name = "FGestionCuves"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Rôle                    : Fenêtre gérant les cuves
' Nom                    : FGestionCuves.frm
' Date de création : 11/01/2001
' Détails                :
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

'--- déclarations obligatoires ---
Option Explicit

'--- options générales ---
Option Base 1
DefVar A-Z
    
'--- constantes privées ---
Private Const NBR_CHAMPS_ETATS As Integer = 11

Private Const IndexSelectionGestionAutoPompe As Integer = 1
Private Const IndexSelectionArretPompe As Integer = 2
Private Const IndexSelectionMarchePompe As Integer = 3

Private Const X_MINI_POMPE As Single = 2
Private Const X_MAXI_POMPE As Single = 626                            'encadrement de la zone pompe
Private Const Y_MINI_POMPE As Single = 2                                 '(X mini, Y mini, X maxi, Y Maxi) de la zone
Private Const Y_MAXI_POMPE As Single = 16

Private Const X_MINI_CHAUFFAGE As Single = 2
Private Const X_MAXI_CHAUFFAGE As Single = 626                      'encadrement de la zone chauffage
Private Const Y_MINI_CHAUFFAGE As Single = 46                        '(X mini, Y mini, X maxi, Y Maxi) de la zone
Private Const Y_MAXI_CHAUFFAGE As Single = 60

Private Const LARGEUR_SEGMENT As Single = 13                       'largeur d'un segment correspondant à 1/2 heure

Private Const TITRE_FENETRE As String = "GESTION DES CUVES"
Private Const TITRE_MESSAGES As String = TITRE_FENETRE

'--- variables privées ---
Private PremiereActivation As Boolean
Private InterdireEvenements As Boolean             'pour interdire certains évènements
Private MemModeEnCoursPompe As Integer

Private IdxChampEtatsChauffage As Integer
Private IdxChampEtatsRefroidissementBain As Integer
Private IdxChampEtatsPompe As Integer
Private IdxChampEtatsNiveaux As Integer
Private IdxChampEtatsEVEau As Integer
Private IdxChampEtatsAgitationBain As Integer
Private IdxChampEtatsCouvercles As Integer
Private IdxChampTemperature As Integer

Private DateFicProgCyclique As Date                   'date du fichier du programmateur cyclique
Private MemDateFicProgCyclique As Date           'mémoire de la date du fichier du programmateur cyclique

'--- variables publiques ---
Public NumFenetre As Long                             'numéro de la fenêtre lorsqu'elle devient active
Public CuveEnCours As Integer                      'cuve en cours

Private Sub CBAgrandirFENETRE_Click()
    On Error Resume Next
    Me.WindowState = vbMaximized
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
    
    '--- enregistre les valeurs dans l'automate ---
    If EnregistreVersAutomate = True Then
        
        '--- enregistrement des valeurs de la fenêtre ---
        EnregistreValeursfenetre
        
        '--- sauvegarde de la configuration ---
        'SauveConfiguration_Old
        
        '--- affectation ---
        CBValider.Enabled = False
    
    Else
        
        '--- lancer un message d'erreur ---
        'Bidon = MessageErreur(TITRE_MESSAGES, MESSAGE_500)
        
        '--- restitue les dernières manipulations et valeurs sur la fenêtre ---
        LectureValeursFenetre
        
    End If

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

'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Rôle      : Enregistre les valeurs dans l'automate
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Private Function EnregistreVersAutomate() As Boolean
    
    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next
    
    '--- constantes privées ---
    Const NOM_GROUPE As String = "MODES_POMPES_CUVES"
    
    '--- déclaration ---
    Dim ValeurRetourneeAPI As Long                  'valeur retournée par une fonction concernant le dialogue avec l'automate
    Dim NomVariable As String                            'nom de la variable
    
    '--- valeur par défaut ---
    EnregistreVersAutomate = True
    
    '--- blocage du timer ---
    TimerEtatsCuve.Enabled = False
    
    Dim NumCuve As Integer
    NumCuve = TEtatsCuves(CuveEnCours).IndexAutomate

    '--- affectation du nom de la variable ---
    NomVariable = "ModePompePCCuve" & Right("00" & NumCuve, 2)
                
    '--- écriture ---
    If PROGRAMME_AVEC_AUTOMATE = True Then
        ValeurRetourneeAPI = APIEcritureVariableNommee(NOM_GROUPE, NomVariable, MemModeEnCoursPompe)
        If ValeurRetourneeAPI <> 0 Then
            EnregistreVersAutomate = False
        End If
    End If

    '--- déblocage du timer ---
    TimerEtatsCuve.Enabled = True
    
End Function

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

    '--- couleurs correspondantes aux légendes ---
    CouleursLegendes

    '--- RAZ des dates (pour le contrôle du programmateur cyclique) ---
    DateFicProgCyclique = Empty
    MemDateFicProgCyclique = Empty
    
    '--- analyse de la présence de la pompe en fonction de la cuve ---
    Select Case CuveEnCours
    
        Case CUVES_REGULATION.C_DEC, CUVES_REGULATION.C_C31, CUVES_REGULATION.C_C32  'CUVES_REGULATION.C_SAT,
            '--- cuves avec pompe ---
            Set PBEchelle24H.Picture = TImgEchelles24H(ECHELLES_24H.E_POMPE_CHAUFFAGE)
    
        'Case CUVES_REGULATION.C_C37 'SZP
            '--- cas de l'étuve ---
            'Set PBEchelle24H.Picture = TImgEchelles24H(ECHELLES_24H.E_VENTILATION_CHAUFFAGE)
            
        Case Else
            '--- cuves sans pompe ---
            Set PBEchelle24H.Picture = TImgEchelles24H(ECHELLES_24H.E_CHAUFFAGE)
    
    End Select
    
    '--- affichage des libellés ---
    AffichageLibellesEtats

    '--- sélecteurs ---
    PositionneSelecteurs
    
    '--- affichage de la cuve gérée ---
    With TEtatsCuves(CuveEnCours).DefinitionCuve
        LRenseignementsFenetre.Caption = "Cuve " & .NomCuve & " - " & .LibelleCuve
    End With
    
    '--- affectation ---
    VSChangementCuve.Enabled = True
    CBValider.Enabled = False
    CBAnnuler.Enabled = False
    
    '--- autoriser les évènements ---
    InterdireEvenements = False
 
End Sub

'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Rôle      : Enregistre les dernières manipulations et valeurs de la fenêtre
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Private Sub EnregistreValeursfenetre()
    
    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next
    
    '--- pompe ---
    'With TEtatsCuves(CuveEnCours)
    '    .ModeEnCoursPompe = MemModeEnCoursPompe
   ' End With
    
End Sub

'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Rôle      : Positionne l'indicateur de contrôle
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Private Sub PositionneIndicateur()
    
    '--- aiguillage en cas d'erreurs ---
'    On Error Resume Next
'
'    '--- déclaration ---
'    Dim a As Integer
'
'    '--- analyse de la présence de la pompe en fonction de la cuve ---
'    Select Case CuveEnCours
'
'        Case CUVE_4, CUVE_5
'
'        Case Else
'            '--- effacement de tous les indicateurs ---
'            For a = IndexSelectionGestionAutoPompe To IndexSelectionMarchePompe
'                IIndicateur(a).Visible = False
'            Next a
'
'            '--- affichage du bon indicateur ---
'            With TEtatsCuves(CuveEnCours)
'                Select Case .ModeEnCoursPompe
'                    Case 1 To 3: IIndicateur(.ModeEnCoursPompe).Visible = True
'                    Case Else
'                End Select
'            End With
'
'    End Select

End Sub

'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Rôle      : Positionne les sélecteurs des modes de gestion
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Private Sub PositionneSelecteurs()
    
    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next

    '--- déclaration ---
    Dim a As Integer, _
           NumPoste As Integer
    
    '--- positionnement en fonction de l'existence par cuve ---
    With TEtatsCuves(CuveEnCours)
    
        '--- recherche du n°de poste ---
        NumPoste = CorrespondanceCuvesAPIPostes(CuveEnCours)
    
        '--- pompes ---
        If .DefinitionCuve.PresencePompe = True Then
                
            '--- présence d'une pompe au poste ---
            With PBZonePompe
                .Visible = True
                .Refresh
            End With
                
            '--- positionnement du sélecteur ---
            Select Case .ModePompe
                Case MODES_POMPES.M_AUTO To MODES_POMPES.M_FORCER_MARCHE
                    OBModesPompe(.ModePompe).value = True
                Case Else
            End Select
                
        Else
                    
            '--- pas de pompe au poste ---
            With PBZonePompe
                .Visible = False
                .Refresh
            End With
                
        End If
    
    End With
    
 End Sub

'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Rôle      : Lecture du cycle de la cuve
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Private Sub LectureCycleCuve()
    
    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next

    '--- déclaration ---
    Dim a As Integer, b As Integer
    Dim HeureDebutTop As Integer, MinuteDebutTop As Integer
    Dim CouleurRemplissagePompe As Long, CouleurRemplissageChauffage As Long
    Dim HeureFinTop As Integer, MinuteFinTop As Integer
    Dim XDebut As Single, XFin As Single

    '--- analyse de l'échelle graphique ---
    With TProgCyclique(1, CuveEnCours)

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
            
            '--- couleur de remplissage pour une pompe ---
            Select Case .TCyclesPompe(b)
                Case CYCLES_POMPES.CP_ARRET: CouleurRemplissagePompe = COULEURS_ECHELLES_GRAPHIQUES.C_ARRET_POMPE
                Case CYCLES_POMPES.CP_MARCHE: CouleurRemplissagePompe = COULEURS_ECHELLES_GRAPHIQUES.C_MARCHE_POMPE
                Case Else: CouleurRemplissagePompe = COULEURS_ECHELLES_GRAPHIQUES.C_ARRET_POMPE
            End Select
                
            '--- tracer de la pompe ---
            If XFin > X_MAXI_POMPE Then XFin = X_MAXI_POMPE
            PBEchelle24H.AutoRedraw = True
            PBEchelle24H.Line (XDebut, Y_MINI_POMPE)-(XFin, Y_MAXI_POMPE), CouleurRemplissagePompe, BF
            PBEchelle24H.AutoRedraw = False
                
        Next b
                
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
            PBEchelle24H.AutoRedraw = True
            PBEchelle24H.Line (XDebut, Y_MINI_CHAUFFAGE)-(XFin, Y_MAXI_CHAUFFAGE), CouleurRemplissageChauffage, BF
            PBEchelle24H.AutoRedraw = False
        
        Next b
            
    End With
            
End Sub

'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Rôle      : Visualisation des différents états de la cuve
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Private Sub EtatsCuve()
    
    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next

    '--- constantes privées ---
    Const ARRET_GENERAL As String = "ARRET GENERAL"
    
    '--- déclaration ---
    Dim a As Integer
    Dim ClignotantDefaut As Integer, _
           RetourControleTemperature As Integer, _
           NumPoste As Integer
    Static CptAppels As Integer
    Dim CouleurFond As Long, CouleurPlan As Long
    Dim Texte As String

    '--- contrôle du clignotement ---
    CptAppels = IIf(CptAppels > 11, 1, CptAppels + 1)
    ClignotantDefaut = Choose(CptAppels, 1, 1, 0, 0, 1, 1, 0, 0, 1, 1, 0, 0)
    
    '--- contrôle automatique des changements ---
    'If IndexEtatsPompe > 0 Then
        'With TEtatsCuves(CuveEnCours)
            'If .API_Changements = True Then
                
                '--- position de l'indicateur ---
                'PositionneIndicateur
                
                '--- sélecteur Pompe ---
              '  PositionneSelecteur
                
                '--- affectation ---
               ' .API_Changements = False
        
            'End If
        'End With
    'End If

    '--- échelle des 24 heures ---
    DateFicProgCyclique = FileDateTime(RepFicAnodisation & FIC_PROG_CYCLIQUE)
    If DateFicProgCyclique <> MemDateFicProgCyclique Then
        LectureCycleCuve
        MemDateFicProgCyclique = DateFicProgCyclique
    End If
    
    '--- cycle de travail en cours ---
    Texte = "Mode actuel = "
    Select Case TEtatsCuves(CuveEnCours).ModeProduction
        Case MODES_PRODUCTION.M_ARRET: Texte = Texte & "ARRET"
        Case MODES_PRODUCTION.M_VEILLE: Texte = Texte & "VEILLE"
        Case MODES_PRODUCTION.M_PRODUCTION: Texte = Texte & "PRODUCTION"
        Case Else: Texte = "MODE INCONNU NON GERE"
    End Select
    With LCycleTravailEnCours
        If .Caption <> Texte Then
            .Caption = Texte
            .Refresh
        End If
    End With
    
    With TEtatsCuves(CuveEnCours)
        
        '----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
        
        '--- cannes chauffantes ---
        If .DefinitionCuve.NbrChauffages > 0 Then
            
            If TEtatsLigne.MarcheGenerale = True Then
            
                Select Case .EtatsChauffage
                    Case ETATS_CHAUFFAGES.M_ARRET: CouleurFond = COULEURS.VERT_3: CouleurPlan = COULEURS.NOIR: Texte = "ARRET"
                    Case ETATS_CHAUFFAGES.M_MARCHE: CouleurFond = COULEURS.ORANGE_2: CouleurPlan = COULEURS.NOIR: Texte = "MARCHE"
                    Case ETATS_CHAUFFAGES.M_DEFAUT
                        If ClignotantDefaut = 0 Then
                            CouleurFond = COULEURS.BLANC: CouleurPlan = COULEURS.NOIR: Texte = ""
                        Else
                            CouleurFond = COULEURS.ROUGE_3: CouleurPlan = COULEURS.JAUNE_3: Texte = "DEFAUT"
                        End If
                    Case Else
                End Select
                   
            Else
                
                '--- arrêt général de la ligne ---
                CouleurFond = COULEURS.MAGENTA_0: CouleurPlan = COULEURS.NOIR: Texte = ARRET_GENERAL
            
            End If
    
             '--- affichage ---
            With LEtats(IdxChampEtatsChauffage)
                If .Caption <> Texte Then
                    SCadresEtats(IdxChampEtatsChauffage).BackColor = CouleurFond
                    .ForeColor = CouleurPlan
                    .Caption = Texte
                End If
            End With
             
        End If
        
        '----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
        
        '--- refroidissement du bain ---
        If IdxChampEtatsRefroidissementBain > 0 Then
            
            If TEtatsLigne.MarcheGenerale = True Then
                
                Select Case .EtatsRefroidissementBain
                    Case ETATS_REFROIDISSEMENT_BAIN.M_ARRET: CouleurFond = COULEURS.VERT_3: CouleurPlan = COULEURS.NOIR: Texte = "ARRET"
                    Case ETATS_REFROIDISSEMENT_BAIN.M_MARCHE: CouleurFond = COULEURS.ORANGE_2: CouleurPlan = COULEURS.NOIR: Texte = "MARCHE"
                    Case Else
                        If ClignotantDefaut = 0 Then
                            CouleurFond = COULEURS.BLANC: CouleurPlan = COULEURS.NOIR: Texte = ""
                        Else
                            CouleurFond = COULEURS.ROUGE_3: CouleurPlan = COULEURS.JAUNE_3: Texte = "DEFAUT"
                        End If
                End Select
            
            Else
            
                '--- arrêt général de la ligne ---
                CouleurFond = COULEURS.MAGENTA_0: CouleurPlan = COULEURS.NOIR: Texte = ARRET_GENERAL
            
            End If
            
            '--- affichage ---
            With LEtats(IdxChampEtatsRefroidissementBain)
                If .Caption <> Texte Then
                    SCadresEtats(IdxChampEtatsRefroidissementBain).BackColor = CouleurFond
                    .ForeColor = CouleurPlan
                    .Caption = Texte
                End If
            End With
        
        End If
        
        '----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    
        '--- la pompe ---
        If IdxChampEtatsPompe > 0 Then
            
            If TEtatsLigne.MarcheGenerale = True Then
               
               Select Case .EtatsPompe
                   Case ETATS_POMPES.E_ARRET: CouleurFond = COULEURS.VERT_3: CouleurPlan = COULEURS.NOIR: Texte = "ARRET"
                   Case ETATS_POMPES.E_MARCHE: CouleurFond = COULEURS.ORANGE_2: CouleurPlan = COULEURS.NOIR: Texte = "MARCHE"
                   Case Else
                        If ClignotantDefaut = 0 Then
                            CouleurFond = COULEURS.BLANC: CouleurPlan = COULEURS.NOIR: Texte = ""
                        Else
                            CouleurFond = COULEURS.ROUGE_3: CouleurPlan = COULEURS.JAUNE_3: Texte = "DEFAUT"
                        End If
               End Select
            
            Else
            
               '--- arrêt général de la ligne ---
               CouleurFond = COULEURS.MAGENTA_0: CouleurPlan = COULEURS.NOIR: Texte = ARRET_GENERAL
            
            End If
            
            '--- affichage ---
            With LEtats(IdxChampEtatsPompe)
               If .Caption <> Texte Then
                   SCadresEtats(IdxChampEtatsPompe).BackColor = CouleurFond
                   .ForeColor = CouleurPlan
                   .Caption = Texte
               End If
            End With
        
        End If
       
        '----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
       
        '--- niveaux ---
        If IdxChampEtatsNiveaux > 0 Then
            
            If TEtatsLigne.MarcheGenerale = True Then
                
                Select Case .EtatsNiveaux
                    Case ETATS_NIVEAUX.E_INTERMEDIAIRE_BAS: CouleurFond = COULEURS.VERT_3: CouleurPlan = COULEURS.NOIR: Texte = "INTERMEDIAIRE BAS"
                    Case ETATS_NIVEAUX.E_NORMAL: CouleurFond = COULEURS.VERT_3: CouleurPlan = COULEURS.NOIR: Texte = "NORMAL"
                    Case ETATS_NIVEAUX.E_INTERMEDIAIRE_HAUT: CouleurFond = COULEURS.VERT_3: CouleurPlan = COULEURS.NOIR: Texte = "INTERMEDIAIRE HAUT"
                    Case ETATS_NIVEAUX.E_TRES_HAUT
                        If ClignotantDefaut = 0 Then
                            CouleurFond = COULEURS.BLANC: CouleurPlan = COULEURS.NOIR: Texte = ""
                        Else
                            CouleurFond = COULEURS.ROUGE_3: CouleurPlan = COULEURS.JAUNE_3: Texte = "TROP HAUT"
                        End If
                    Case ETATS_NIVEAUX.E_TRES_BAS
                        If ClignotantDefaut = 0 Then
                            CouleurFond = COULEURS.BLANC: CouleurPlan = COULEURS.NOIR: Texte = ""
                        Else
                            CouleurFond = COULEURS.ROUGE_3: CouleurPlan = COULEURS.JAUNE_3: Texte = "TROP BAS"
                        End If
                    Case Else
                End Select
            
            Else
                
                '--- arrêt général de la ligne ---
                CouleurFond = COULEURS.MAGENTA_0: CouleurPlan = COULEURS.NOIR: Texte = ARRET_GENERAL
            
            End If
            
            '--- affichage ---
            With LEtats(IdxChampEtatsNiveaux)
                If .Caption <> Texte Then
                    SCadresEtats(IdxChampEtatsNiveaux).BackColor = CouleurFond
                    .ForeColor = CouleurPlan
                    .Caption = Texte
                End If
            End With
        
        End If
               
        '----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
               
        '--- électro-vanne d'arrivée d'eau ---
        If IdxChampEtatsEVEau > 0 Then
            
            If TEtatsLigne.MarcheGenerale = True Then
                
                Select Case .EtatsEVEau
                    Case ETATS_EV_EAU.E_FERMEE: CouleurFond = COULEURS.VERT_3: CouleurPlan = COULEURS.NOIR: Texte = "FERMEE"
                    Case ETATS_EV_EAU.E_OUVERTE: CouleurFond = COULEURS.ORANGE_2: CouleurPlan = COULEURS.NOIR: Texte = "OUVERTE"
                    Case Else
                        If ClignotantDefaut = 0 Then
                            CouleurFond = COULEURS.BLANC: CouleurPlan = COULEURS.NOIR: Texte = ""
                        Else
                            CouleurFond = COULEURS.ROUGE_3: CouleurPlan = COULEURS.JAUNE_3: Texte = "DEFAUT"
                        End If
                End Select
            
            Else
            
                '--- arrêt général de la ligne ---
                CouleurFond = COULEURS.MAGENTA_0: CouleurPlan = COULEURS.NOIR: Texte = ARRET_GENERAL
            
            End If
            
            '--- affichage ---
            With LEtats(IdxChampEtatsEVEau)
                If .Caption <> Texte Then
                    SCadresEtats(IdxChampEtatsEVEau).BackColor = CouleurFond
                    .ForeColor = CouleurPlan
                    .Caption = Texte
                End If
            End With
        
        End If

        '----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
        
        '--- agitation du bain ---
        'If IdxChampEtatsAgitationBain > 0 Then
            
            '--- recherche du numéro de poste (car les couvercles se trouvent sur les postes) ---
            'NumPoste = 0
            'Select Case CuveEnCours
                'Case CUVES_REGULATION.C_C13: NumPoste = POSTES.P_C13
                'Case CUVES_REGULATION.C_C14: NumPoste = POSTES.P_C14
                'Case CUVES_REGULATION.C_C15: NumPoste = POSTES.P_C15
                'Case Else
            'End Select
        
            'If NumPoste > 0 Then
                
                'If TEtatsLigne.MarcheGenerale = True Then
                    
                    'With TEtatsPostes(NumPoste)
                    '    Select Case .EtatsAgitationCharge
                    '        Case ETATS_AGITATION_CHARGE.E_ARRET_CENTRE: CouleurFond = COULEURS.VERT_3: CouleurPlan = COULEURS.NOIR: Texte = "ARRET CENTRE"
                    '        Case ETATS_AGITATION_CHARGE.E_MARCHE: CouleurFond = COULEURS.ORANGE_2: CouleurPlan = COULEURS.NOIR: Texte = "MARCHE"
                    '        Case ETATS_AGITATION_CHARGE.E_ARRET_NON_CENTRE
                    '            If ClignotantDefaut = 0 Then
                    '                CouleurFond = COULEURS.BLANC: CouleurPlan = COULEURS.NOIR: Texte = ""
                    '            Else
                    '                CouleurFond = COULEURS.ROUGE_3: CouleurPlan = COULEURS.JAUNE_3: Texte = "NON CENTRE"
                    '            End If
                    '        Case Else
                    '            If TEtatsCuves(CuveEnCours).TNumDefauts.DisjonctionAgitationCharge = True Then
                    '                If ClignotantDefaut = 0 Then
                    '                    CouleurFond = COULEURS.BLANC: CouleurPlan = COULEURS.NOIR: Texte = ""
                    '                Else
                    '                    CouleurFond = COULEURS.ROUGE_3: CouleurPlan = COULEURS.JAUNE_3: Texte = "DEFAUT"
                    '                End If
                    '            Else
                     '               CouleurFond = COULEURS.ROUGE_4: CouleurPlan = COULEURS.JAUNE_3: Texte = "Défaut"
                     '           End If
                     '   End Select
                    'End With
                
                'Else
                
                    '--- arrêt général de la ligne ---
                    'CouleurFond = COULEURS.MAGENTA_0: CouleurPlan = COULEURS.NOIR: Texte = ARRET_GENERAL
                    
                'End If
                
                '--- affichage ---
                'With LEtats(IdxChampEtatsAgitationBain)
                '    If .Caption <> Texte Then
                '        SCadresEtats(IdxChampEtatsAgitationBain).BackColor = CouleurFond
                '        .ForeColor = CouleurPlan
                '        .Caption = Texte
                '    End If
                'End With

            'End If

        'End If
        
        '----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
        
        '--- températures ---
        RetourControleTemperature = ControleTemperature(CuveEnCours)
        With .Temperatures
            Select Case RetourControleTemperature
                Case 0: CouleurFond = COULEURS.VERT_3: CouleurPlan = COULEURS.NOIR: Texte = Format(.TempActuelle, "#00.0°C")
                Case 1: CouleurFond = ORANGE_CUVE: CouleurPlan = COULEURS.NOIR: Texte = Format(.TempActuelle, "#00.0°C")
                Case 2, 3, 4
                    If ClignotantDefaut = 0 Then
                        CouleurFond = COULEURS.BLANC: CouleurPlan = COULEURS.BLANC: Texte = ""
                    Else
                        CouleurFond = COULEURS.ROUGE_3: CouleurPlan = COULEURS.JAUNE_3
                        If RetourControleTemperature = 4 Then
                            Texte = "PT100"
                        Else
                            Texte = Format(.TempActuelle, "#00.0°C")
                        End If
                    End If
                Case 5
                    If ClignotantDefaut = 0 Then
                        CouleurFond = COULEURS.BLANC: CouleurPlan = COULEURS.BLANC: Texte = ""
                    Else
                        CouleurFond = COULEURS.BLEU_4: CouleurPlan = COULEURS.JAUNE_3: Texte = Format(.TempActuelle, "#00.0°C")
                    End If
                Case Else
            End Select
        End With
        With LEtats(IdxChampTemperature)
            If .Caption <> Texte Then
                SCadresEtats(IdxChampTemperature).BackColor = CouleurFond
                .ForeColor = CouleurPlan
                .Caption = Texte
            End If
        End With
    
    End With

End Sub

Private Sub Form_Resize()
    
    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next
    
    '--- déclaration ---
    
    '--- zone mére et fille du déplacement de la fenetre ---
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

Private Sub OBModesPompe_Click(Index As Integer)
    
    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next
    
    '--- affectation ---
    MemModeEnCoursPompe = Index
    CBValider.Enabled = True

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
                    'TODO:  à corriger !!
                    Case 200
                        '--- préparation ---
                        .BackColor = COULEURS.BLANC
                        .ForeColor = COULEURS.NOIR
                    Case CUVES_REGULATION.C_C00 To CUVES_REGULATION.C_C07
                        '--- préparation ---
                        .BackColor = COULEURS.BLANC
                        .ForeColor = COULEURS.NOIR
                    Case CUVES_REGULATION.C_C13 To CUVES_REGULATION.C_C15 ' CUVES_REGULATION.C_C16 SZP
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

Private Sub OBNomsCuves_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next
        
    '--- déclaration ---
    Dim TitreCuve As String
    Static MemTitreCuve As String
    
    With OBNomsCuves(Index)
        
        '--- affectation ---
        TitreCuve = .Caption
        
        If TitreCuve <> MemTitreCuve Then
            
            '--- gestion de la bulle ---
            If TitreCuve <> "" Then
                .ToolTipText = UN_ESPACE & PosteParNom(.Caption).LibellePoste & UN_ESPACE
            End If
        
            '--- affectation ---
            MemTitreCuve = TitreCuve

        End If
    
    End With

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
        
    End If
            
    '--- valeur des curseurs ---
    If Me.WindowState <> vbMaximized And Me.WindowState <> vbMinimized Then
        HSDeplacementFenetre.Max = PBDeplacementFenetre(ZONES_DEPLACEMENT_FENETRE.Z_FILLE).Width - _
                                                         PBDeplacementFenetre(ZONES_DEPLACEMENT_FENETRE.Z_MERE).Width
        VSDeplacementFenetre.Max = PBDeplacementFenetre(ZONES_DEPLACEMENT_FENETRE.Z_FILLE).Height - _
                                                        PBDeplacementFenetre(ZONES_DEPLACEMENT_FENETRE.Z_MERE).Height
    End If

End Sub

Private Sub TimerEtatsCuve_Timer()
    
    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next
    
    '--- appel de la routine ---
    TimerEtatsCuve.Enabled = False
    EtatsCuve
    TimerEtatsCuve.Enabled = True
    
    '--- bip de passage dans la routine UNIQUEMENT POUR LES TESTS ---
    If PROGRAMME_AVEC_AUTOMATE = False Then Beep
    
End Sub

'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Rôle      : Visualise les couleurs correspondantes aux légendes
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Private Sub CouleursLegendes()
    
    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next

    '--- Pompe ---
    PBModesPompe(ARRET).BackColor = COULEURS_ECHELLES_GRAPHIQUES.C_ARRET_POMPE
    PBModesPompe(MARCHE).BackColor = COULEURS_ECHELLES_GRAPHIQUES.C_MARCHE_POMPE

    '--- chauffage ---
    PBModesChauffage(MODES_PRODUCTION.M_ARRET).BackColor = COULEURS_ECHELLES_GRAPHIQUES.C_MODE_ARRET
    PBModesChauffage(MODES_PRODUCTION.M_VEILLE).BackColor = COULEURS_ECHELLES_GRAPHIQUES.C_MODE_VEILLE
    PBModesChauffage(MODES_PRODUCTION.M_PRODUCTION).BackColor = COULEURS_ECHELLES_GRAPHIQUES.C_MODE_PRODUCTION

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
        EtatsCuve
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

    '--- affectation ---
    Dim a As Integer                                'pour les boucles FOR...NEXT

    '--- divers sur la fenêtre ---
    With Me
        .Caption = TITRE_FENETRE
        .WindowState = vbMaximized
    End With
    
    '--- couleurs des fonds ---
    PBDeplacementFenetre(ZONES_DEPLACEMENT_FENETRE.Z_FILLE).Picture = ImgFondVert1
    PBSelectionNomsCuves.Picture = ImgFondGris1
    PBBoutons.Picture = ImgFondDesBoutons
    
    '--- affichage des noms des cuves ---
    'TMP
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
    With TimerEtatsCuve
        .Enabled = False
        .Interval = 0
    End With
    
    '--- déchargement de la fenêtre ---
    Me.Visible = False
    Unload Me
    Set OccFGestionCuves = Nothing
    
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
    
    '--- valeurs de la fenêtre ---
    LectureValeursFenetre
    
    '--- position de l'indicateur ---
    PositionneIndicateur
    
    '--- visualisation des différents états de la cuves ---
    EtatsCuve

    '--- lancement du timer ---
    TimerEtatsCuve.Enabled = True
                
End Sub

Private Sub VSDeplacementFENETRE_Change()
    On Error Resume Next
    PBDeplacementFenetre(ZONES_DEPLACEMENT_FENETRE.Z_FILLE).Top = -VSDeplacementFenetre.value
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
    For a = OBNomsCuves.UBound To OBNomsCuves.LBound
        OBNomsCuves(a).Enabled = Not (Blocage)
    Next a
    
    '--- blocage de la barre de scrolling ---
    VSChangementCuve.Enabled = Not (Blocage)
    
End Sub

'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Rôle      : Affiche les libellés des états
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Private Sub AffichageLibellesEtats()
    
    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next

    '--- déclaration ---
    Dim a As Integer, _
            IdxChamps As Integer, _
            IndexSuppression As Integer
    
    '--- affectation ---
    IdxChamps = 0
    IdxChampEtatsChauffage = 0
    IdxChampEtatsRefroidissementBain = 0
    IdxChampEtatsPompe = 0
    IdxChampEtatsNiveaux = 0
    IdxChampEtatsEVEau = 0
    IdxChampEtatsAgitationBain = 0
    IdxChampEtatsCouvercles = 0
    IdxChampTemperature = 12
    
    With TEtatsCuves(CuveEnCours)
    
        '--- cannes chauffantes ---
        If .DefinitionCuve.NbrChauffages >= 1 Then
            Inc IdxChamps
            IdxChampEtatsChauffage = IdxChamps
        End If
        
        '--- refroidissement du bain ---
        If .DefinitionCuve.PresenceRefroidissementBain = True Then
            Inc IdxChamps
            IdxChampEtatsRefroidissementBain = IdxChamps
        End If
    
        '--- pompe ---
        If .DefinitionCuve.PresencePompe = True Then
            Inc IdxChamps
            IdxChampEtatsPompe = IdxChamps
        End If
    
        '--- niveaux ---
        If .DefinitionCuve.PresenceNiveauBas = True Or .DefinitionCuve.PresenceNiveauHaut = True Then
            Inc IdxChamps
            IdxChampEtatsNiveaux = IdxChamps
        End If
    
        '--- EV eau ---
        If .DefinitionCuve.PresenceEVEau = True Then
            Inc IdxChamps
            IdxChampEtatsEVEau = IdxChamps
        End If
        
    End With
    
    '--- affichage des libellés ---
    If IdxChampEtatsChauffage <> 0 Then
        With LLibellesEtats(IdxChampEtatsChauffage)
            .Caption = "Chauffage"
            .Visible = True
            .Refresh
        End With
    End If
    If IdxChampEtatsRefroidissementBain <> 0 Then
        With LLibellesEtats(IdxChampEtatsRefroidissementBain)
            .Caption = "Refroidissement"
            .Visible = True
            .Refresh
        End With
    End If
    If IdxChampEtatsPompe <> 0 Then
        With LLibellesEtats(IdxChampEtatsPompe)
            .Caption = "Pompe"
            .Visible = True
            .Refresh
        End With
    End If
    If IdxChampEtatsNiveaux <> 0 Then
        With LLibellesEtats(IdxChampEtatsNiveaux)
            .Caption = "Niveau"
            .Visible = True
            .Refresh
        End With
    End If
    If IdxChampEtatsEVEau <> 0 Then
        With LLibellesEtats(IdxChampEtatsEVEau)
            .Caption = "Arrivée d'eau"
            .Visible = True
            .Refresh
        End With
    End If
    If IdxChampEtatsAgitationBain <> 0 Then
        With LLibellesEtats(IdxChampEtatsAgitationBain)
            .Caption = "Agitation du bain"
            .Visible = True
            .Refresh
        End With
    End If
    If IdxChampEtatsCouvercles <> 0 Then
        With LLibellesEtats(IdxChampEtatsCouvercles)
            .Caption = "Couvercles"
            .Visible = True
            .Refresh
        End With
    End If

    '--- suppression des champs inutilisés ---
    For a = 1 To NBR_CHAMPS_ETATS
        If a <= IdxChamps Then
            With SCadresLibellesEtats(a)
                .Visible = True
                .Refresh
            End With
            With LLibellesEtats(a)
                .Visible = True
                .Refresh
            End With
            With SCadresEtats(a)
                .Visible = True
                .Refresh
            End With
            With LEtats(a)
                .Visible = True
                .Refresh
            End With
        Else
            With SCadresLibellesEtats(a)
                .Visible = False
                .Refresh
            End With
            With LLibellesEtats(a)
                .Visible = False
                .Refresh
            End With
            With SCadresEtats(a)
                .Visible = False
                .Refresh
            End With
            With LEtats(a)
                .Visible = False
                .Refresh
            End With
        End If
    Next a
    
End Sub



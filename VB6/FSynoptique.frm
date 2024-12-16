VERSION 5.00
Object = "{1C0489F8-9EFD-423D-887A-315387F18C8F}#1.0#0"; "vsflex8l.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.ocx"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "richtx32.ocx"
Object = "{38911DA0-E448-11D0-84A3-00DD01104159}#1.1#0"; "comct332.ocx"
Begin VB.Form FSynoptique 
   Appearance      =   0  'Flat
   BorderStyle     =   0  'None
   ClientHeight    =   15600
   ClientLeft      =   -60
   ClientTop       =   2670
   ClientWidth     =   28860
   FillColor       =   &H80000005&
   ForeColor       =   &H00000000&
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   1040
   ScaleMode       =   0  'User
   ScaleWidth      =   1960
   ShowInTaskbar   =   0   'False
   Begin VB.Timer TimerEtatsLigne 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   12300
      Top             =   13740
   End
   Begin MSComctlLib.ImageList ILOutilsDialogues2 
      Left            =   10380
      Top             =   13740
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   169
      ImageHeight     =   19
      MaskColor       =   16711935
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   8
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FSynoptique.frx":0000
            Key             =   "renseignements"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FSynoptique.frx":2606
            Key             =   "renseignements en selection"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FSynoptique.frx":4C0C
            Key             =   "questions reponses"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FSynoptique.frx":7212
            Key             =   "questions reponses en selection"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FSynoptique.frx":9818
            Key             =   "previsionnel"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FSynoptique.frx":BE1E
            Key             =   "previsionnel en selection"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FSynoptique.frx":E424
            Key             =   "entrees charges"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FSynoptique.frx":10A2A
            Key             =   "entrees charges en selection"
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox PBEtatsPrincipaux 
      Height          =   12855
      Left            =   21480
      ScaleHeight     =   853
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   989
      TabIndex        =   106
      Top             =   5460
      Width           =   14895
      Begin VB.PictureBox PBManuelP2 
         BackColor       =   &H00C0FFC0&
         Height          =   1395
         Left            =   3600
         ScaleHeight     =   1335
         ScaleWidth      =   3075
         TabIndex        =   161
         Top             =   6540
         Width           =   3135
         Begin VB.CommandButton CBMonteeDescenteAccrochesPCP2 
            BackColor       =   &H0080FFFF&
            Caption         =   "DESCENTE"
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
            Left            =   1620
            Style           =   1  'Graphical
            TabIndex        =   163
            Top             =   780
            Width           =   1215
         End
         Begin VB.CommandButton CBMonteeDescenteAccrochesPCP2 
            BackColor       =   &H0080FFFF&
            Caption         =   "MONTEE"
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
            Index           =   0
            Left            =   240
            Style           =   1  'Graphical
            TabIndex        =   162
            Top             =   780
            Width           =   1215
         End
         Begin VB.Label LLibelles 
            Alignment       =   2  'Center
            BackColor       =   &H00FF0000&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "MANUEL du PONT 2"
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
            Height          =   315
            Index           =   16
            Left            =   0
            TabIndex        =   165
            Top             =   0
            Width           =   3075
         End
         Begin VB.Label LLibelles 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FFC0C0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "ACCROCHES"
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
            Height          =   255
            Index           =   21
            Left            =   120
            TabIndex        =   164
            Top             =   420
            Width           =   2835
         End
      End
      Begin VB.PictureBox PBManuelP1 
         BackColor       =   &H00C0FFC0&
         Height          =   1395
         Left            =   240
         ScaleHeight     =   1335
         ScaleWidth      =   3075
         TabIndex        =   156
         Top             =   6540
         Width           =   3135
         Begin VB.CommandButton CBMonteeDescenteAccrochesPCP1 
            BackColor       =   &H0080FFFF&
            Caption         =   "MONTEE"
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
            Index           =   0
            Left            =   240
            Style           =   1  'Graphical
            TabIndex        =   158
            Top             =   780
            Width           =   1215
         End
         Begin VB.CommandButton CBMonteeDescenteAccrochesPCP1 
            BackColor       =   &H0080FFFF&
            Caption         =   "DESCENTE"
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
            Left            =   1620
            Style           =   1  'Graphical
            TabIndex        =   157
            Top             =   780
            Width           =   1215
         End
         Begin VB.Label LLibelles 
            Alignment       =   2  'Center
            BackColor       =   &H00FF0000&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "MANUEL du PONT 1"
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
            Height          =   315
            Index           =   15
            Left            =   0
            TabIndex        =   160
            Top             =   0
            Width           =   3075
         End
         Begin VB.Label LLibelles 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FFC0C0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "ACCROCHES"
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
            Height          =   255
            Index           =   22
            Left            =   120
            TabIndex        =   159
            Top             =   420
            Width           =   2835
         End
      End
      Begin VB.Frame FPont2 
         BackColor       =   &H00C0C0C0&
         Height          =   1470
         Left            =   11880
         TabIndex        =   111
         Top             =   2640
         Width           =   2595
         Begin VB.Label LLibelles 
            Alignment       =   2  'Center
            BackColor       =   &H00FF0000&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "PONT 2"
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
            Height          =   315
            Index           =   1
            Left            =   0
            TabIndex        =   114
            Top             =   0
            Width           =   2595
         End
         Begin VB.Label LMaintenanceAAutomatique 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
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
            ForeColor       =   &H80000008&
            Height          =   315
            Index           =   2
            Left            =   120
            TabIndex        =   113
            Top             =   420
            Width           =   1875
         End
         Begin VB.Label LTypeSequence 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
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
            ForeColor       =   &H80000008&
            Height          =   495
            Index           =   2
            Left            =   120
            TabIndex        =   112
            Top             =   840
            Width           =   2355
         End
         Begin VB.Image IEtatsPonts 
            Appearance      =   0  'Flat
            BorderStyle     =   1  'Fixed Single
            Height          =   315
            Index           =   2
            Left            =   2100
            Picture         =   "FSynoptique.frx":13030
            Stretch         =   -1  'True
            Top             =   420
            Width           =   375
         End
      End
      Begin VB.Frame FPont1 
         BackColor       =   &H00C0C0C0&
         Height          =   1470
         Index           =   0
         Left            =   11880
         TabIndex        =   107
         Top             =   960
         Width           =   2595
         Begin VB.Label LMaintenanceAAutomatique 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
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
            ForeColor       =   &H80000008&
            Height          =   315
            Index           =   1
            Left            =   120
            TabIndex        =   110
            Top             =   420
            Width           =   1875
         End
         Begin VB.Label LLibelles 
            Alignment       =   2  'Center
            BackColor       =   &H00FF0000&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "PONT 1"
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
            Height          =   315
            Index           =   0
            Left            =   0
            TabIndex        =   109
            Top             =   0
            Width           =   2595
         End
         Begin VB.Label LTypeSequence 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
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
            ForeColor       =   &H80000008&
            Height          =   495
            Index           =   1
            Left            =   120
            TabIndex        =   108
            Top             =   840
            Width           =   2355
         End
         Begin VB.Image IEtatsPonts 
            Appearance      =   0  'Flat
            BorderStyle     =   1  'Fixed Single
            Height          =   315
            Index           =   1
            Left            =   2100
            Picture         =   "FSynoptique.frx":133F6
            Stretch         =   -1  'True
            Top             =   420
            Width           =   375
         End
      End
      Begin VSFlex8LCtl.VSFlexGrid VSFGEnCours 
         Height          =   5775
         Left            =   240
         TabIndex        =   116
         Top             =   240
         Width           =   11235
         _cx             =   19817
         _cy             =   10186
         Appearance      =   0
         BorderStyle     =   1
         Enabled         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MousePointer    =   0
         BackColor       =   16761024
         ForeColor       =   -2147483640
         BackColorFixed  =   12582912
         ForeColorFixed  =   -2147483639
         BackColorSel    =   255
         ForeColorSel    =   -2147483634
         BackColorBkg    =   -2147483636
         BackColorAlternate=   16761024
         GridColor       =   8421504
         GridColorFixed  =   16777215
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483642
         FocusRect       =   3
         HighLight       =   0
         AllowSelection  =   0   'False
         AllowBigSelection=   0   'False
         AllowUserResizing=   0
         SelectionMode   =   0
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   0
         Rows            =   50
         Cols            =   3
         FixedRows       =   1
         FixedCols       =   1
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"FSynoptique.frx":137BC
         ScrollTrack     =   0   'False
         ScrollBars      =   3
         ScrollTips      =   0   'False
         MergeCells      =   0
         MergeCompare    =   111
         AutoResize      =   0   'False
         AutoSizeMode    =   0
         AutoSearch      =   0
         AutoSearchDelay =   2
         MultiTotals     =   0   'False
         SubtotalPosition=   1
         OutlineBar      =   0
         OutlineCol      =   0
         Ellipsis        =   0
         ExplorerBar     =   0
         PicturesOver    =   0   'False
         FillStyle       =   0
         RightToLeft     =   0   'False
         PictureType     =   0
         TabBehavior     =   0
         OwnerDraw       =   0
         Editable        =   2
         ShowComboButton =   1
         WordWrap        =   -1  'True
         TextStyle       =   0
         TextStyleFixed  =   0
         OleDragMode     =   0
         OleDropMode     =   0
         ComboSearch     =   0
         AutoSizeMouse   =   0   'False
         FrozenRows      =   0
         FrozenCols      =   0
         AllowUserFreezing=   0
         BackColorFrozen =   0
         ForeColorFrozen =   0
         WallPaperAlignment=   9
         AccessibleName  =   ""
         AccessibleDescription=   ""
         AccessibleValue =   ""
         AccessibleRole  =   24
      End
      Begin VB.Line LDecoration 
         Index           =   1
         X1              =   784
         X2              =   972
         Y1              =   168
         Y2              =   168
      End
      Begin VB.Line LDecoration 
         Index           =   0
         X1              =   784
         X2              =   972
         Y1              =   56
         Y2              =   56
      End
      Begin VB.Label LControleOperateurPonts 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H000000C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Contrôle par l'opérateur du PONT 2"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   555
         Index           =   2
         Left            =   11940
         TabIndex        =   254
         Top             =   5280
         Visible         =   0   'False
         Width           =   2475
      End
      Begin VB.Label LControleOperateurPonts 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H000000C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Contrôle par l'opérateur du PONT 1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   555
         Index           =   1
         Left            =   11940
         TabIndex        =   253
         Top             =   4560
         Visible         =   0   'False
         Width           =   2475
      End
      Begin VB.Label LMarcheGenerale 
         Alignment       =   2  'Center
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " MARCHE GENERALE "
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
         Left            =   11880
         TabIndex        =   115
         Top             =   405
         Width           =   2595
      End
      Begin VB.Shape SDecoration 
         BackColor       =   &H00C0E0FF&
         BackStyle       =   1  'Opaque
         Height          =   3975
         Index           =   0
         Left            =   11760
         Top             =   240
         Width           =   2835
      End
      Begin VB.Shape SControleOperateur 
         BackColor       =   &H00C0E0FF&
         BackStyle       =   1  'Opaque
         Height          =   1635
         Left            =   11760
         Shape           =   4  'Rounded Rectangle
         Top             =   4380
         Visible         =   0   'False
         Width           =   2835
      End
   End
   Begin VB.Timer TimerSynoptique 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   11760
      Top             =   13740
   End
   Begin VB.PictureBox PBSynoptique 
      AutoRedraw      =   -1  'True
      Height          =   5475
      Left            =   -360
      ScaleHeight     =   361
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   1919
      TabIndex        =   1
      Top             =   720
      Width           =   28845
      Begin VB.PictureBox PBEtatsLigne 
         BackColor       =   &H00E0E0E0&
         Height          =   12195
         Left            =   420
         ScaleHeight     =   809
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   1877
         TabIndex        =   12
         Top             =   5400
         Width           =   28215
         Begin VB.Frame FApresAnodisation 
            Height          =   10335
            Left            =   10680
            TabIndex        =   170
            Top             =   240
            Width           =   10095
            Begin Anodisation.ImageMask IMProgrammateurCycliqueCuves 
               Height          =   360
               Index           =   12
               Left            =   5280
               TabIndex        =   237
               Top             =   840
               Visible         =   0   'False
               Width           =   330
               _ExtentX        =   582
               _ExtentY        =   635
               Picture         =   "FSynoptique.frx":1380B
               MaskPicture     =   "FSynoptique.frx":13EBD
               MaskColor       =   16711935
            End
            Begin Anodisation.ImageMask IMProgrammateurCycliqueCuves 
               Height          =   360
               Index           =   13
               Left            =   5280
               TabIndex        =   238
               Top             =   2100
               Width           =   330
               _ExtentX        =   582
               _ExtentY        =   635
               Picture         =   "FSynoptique.frx":1456F
               MaskPicture     =   "FSynoptique.frx":14C21
               MaskColor       =   16711935
            End
            Begin Anodisation.ImageMask IMProgrammateurCycliqueCuves 
               Height          =   360
               Index           =   14
               Left            =   5280
               TabIndex        =   243
               Top             =   4200
               Width           =   330
               _ExtentX        =   582
               _ExtentY        =   635
               Picture         =   "FSynoptique.frx":152D3
               MaskPicture     =   "FSynoptique.frx":15985
               MaskColor       =   16711935
            End
            Begin Anodisation.ImageMask IMProgrammateurCycliqueCuves 
               Height          =   360
               Index           =   15
               Left            =   5280
               TabIndex        =   244
               Top             =   4620
               Width           =   330
               _ExtentX        =   582
               _ExtentY        =   635
               Picture         =   "FSynoptique.frx":16037
               MaskPicture     =   "FSynoptique.frx":166E9
               MaskColor       =   16711935
            End
            Begin Anodisation.ImageMask IMProgrammateurCycliqueCuves 
               Height          =   360
               Index           =   16
               Left            =   5280
               TabIndex        =   245
               Top             =   5880
               Width           =   330
               _ExtentX        =   582
               _ExtentY        =   635
               Picture         =   "FSynoptique.frx":16D9B
               MaskPicture     =   "FSynoptique.frx":1744D
               MaskColor       =   16711935
            End
            Begin Anodisation.ImageMask IMProgrammateurCycliqueCuves 
               Height          =   360
               Index           =   17
               Left            =   5280
               TabIndex        =   246
               Top             =   6300
               Width           =   330
               _ExtentX        =   582
               _ExtentY        =   635
               Picture         =   "FSynoptique.frx":17AFF
               MaskPicture     =   "FSynoptique.frx":181B1
               MaskColor       =   16711935
            End
            Begin Anodisation.ImageMask IMProgrammateurCycliqueCuves 
               Height          =   360
               Index           =   18
               Left            =   5280
               TabIndex        =   247
               Top             =   6900
               Visible         =   0   'False
               Width           =   330
               _ExtentX        =   582
               _ExtentY        =   635
               Picture         =   "FSynoptique.frx":18863
               MaskPicture     =   "FSynoptique.frx":18F15
               MaskColor       =   16711935
            End
            Begin VB.Label LNomsPostes 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H00C0E0FF&
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
               ForeColor       =   &H80000008&
               Height          =   315
               Index           =   44
               Left            =   180
               TabIndex        =   313
               Top             =   8520
               Width           =   615
            End
            Begin VB.Label LNumCharges 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
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
               ForeColor       =   &H80000008&
               Height          =   315
               Index           =   44
               Left            =   4260
               TabIndex        =   312
               Top             =   8520
               Width           =   435
            End
            Begin VB.Image IEtatsPostes 
               Appearance      =   0  'Flat
               BorderStyle     =   1  'Fixed Single
               Height          =   315
               Index           =   44
               Left            =   4680
               Picture         =   "FSynoptique.frx":195C7
               Stretch         =   -1  'True
               Top             =   8520
               Width           =   435
            End
            Begin VB.Label LLibellesPostes 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
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
               ForeColor       =   &H80000008&
               Height          =   315
               Index           =   44
               Left            =   780
               TabIndex        =   311
               Top             =   8520
               Width           =   3495
            End
            Begin VB.Label LLibellesPostes 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
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
               ForeColor       =   &H80000008&
               Height          =   315
               Index           =   43
               Left            =   780
               TabIndex        =   310
               Top             =   8040
               Width           =   3495
            End
            Begin VB.Image IEtatsPostes 
               Appearance      =   0  'Flat
               BorderStyle     =   1  'Fixed Single
               Height          =   315
               Index           =   43
               Left            =   4680
               Picture         =   "FSynoptique.frx":1998D
               Stretch         =   -1  'True
               Top             =   8040
               Width           =   435
            End
            Begin VB.Label LNumCharges 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
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
               ForeColor       =   &H80000008&
               Height          =   315
               Index           =   43
               Left            =   4260
               TabIndex        =   309
               Top             =   8040
               Width           =   435
            End
            Begin VB.Label LNomsPostes 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H00C0E0FF&
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
               ForeColor       =   &H80000008&
               Height          =   315
               Index           =   43
               Left            =   180
               TabIndex        =   308
               Top             =   8040
               Width           =   615
            End
            Begin VB.Label LManuAutoRegulation 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H00C0E0FF&
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
               ForeColor       =   &H80000008&
               Height          =   315
               Index           =   18
               Left            =   5760
               TabIndex        =   298
               Top             =   6900
               Visible         =   0   'False
               Width           =   1155
            End
            Begin VB.Label LManuAutoRegulation 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H00C0E0FF&
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
               ForeColor       =   &H80000008&
               Height          =   315
               Index           =   17
               Left            =   5760
               TabIndex        =   297
               Top             =   6300
               Width           =   1155
            End
            Begin VB.Label LManuAutoRegulation 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H00C0E0FF&
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
               ForeColor       =   &H80000008&
               Height          =   315
               Index           =   16
               Left            =   5760
               TabIndex        =   296
               Top             =   5880
               Width           =   1155
            End
            Begin VB.Label LManuAutoRegulation 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H00C0E0FF&
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
               ForeColor       =   &H80000008&
               Height          =   315
               Index           =   15
               Left            =   5760
               TabIndex        =   295
               Top             =   4620
               Width           =   1155
            End
            Begin VB.Label LManuAutoRegulation 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H00C0E0FF&
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
               ForeColor       =   &H80000008&
               Height          =   315
               Index           =   14
               Left            =   5760
               TabIndex        =   294
               Top             =   4200
               Width           =   1155
            End
            Begin VB.Label LManuAutoRegulation 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H00C0E0FF&
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
               ForeColor       =   &H80000008&
               Height          =   315
               Index           =   13
               Left            =   5760
               TabIndex        =   293
               Top             =   2100
               Width           =   1155
            End
            Begin VB.Label LManuAutoRegulation 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H00C0E0FF&
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
               ForeColor       =   &H80000008&
               Height          =   315
               Index           =   12
               Left            =   5760
               TabIndex        =   292
               Top             =   840
               Visible         =   0   'False
               Width           =   1155
            End
            Begin VB.Label LLibelles 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H00FF0000&
               BorderStyle     =   1  'Fixed Single
               Caption         =   "Régulation"
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
               Height          =   315
               Index           =   31
               Left            =   5760
               TabIndex        =   280
               Top             =   420
               Width           =   1155
            End
            Begin VB.Label LLibelles 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H00FF0000&
               BorderStyle     =   1  'Fixed Single
               Caption         =   "Températures"
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
               Height          =   315
               Index           =   28
               Left            =   7020
               TabIndex        =   277
               Top             =   420
               Width           =   1395
            End
            Begin VB.Label LLibelles 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H00FF0000&
               BorderStyle     =   1  'Fixed Single
               Caption         =   "Niveaux"
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
               Height          =   315
               Index           =   27
               Left            =   8520
               TabIndex        =   276
               Top             =   420
               Width           =   1395
            End
            Begin VB.Label LNiveaux 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H00C0E0FF&
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
               ForeColor       =   &H80000008&
               Height          =   315
               Index           =   17
               Left            =   8520
               TabIndex        =   271
               Top             =   6300
               Width           =   1395
            End
            Begin VB.Label LNiveaux 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H00C0E0FF&
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
               ForeColor       =   &H80000008&
               Height          =   315
               Index           =   16
               Left            =   8520
               TabIndex        =   270
               Top             =   5880
               Width           =   1395
            End
            Begin VB.Label LNiveaux 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H00C0E0FF&
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
               ForeColor       =   &H80000008&
               Height          =   315
               Index           =   15
               Left            =   8520
               TabIndex        =   269
               Top             =   4620
               Width           =   1395
            End
            Begin VB.Label LNiveaux 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H00C0E0FF&
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
               ForeColor       =   &H80000008&
               Height          =   315
               Index           =   14
               Left            =   8520
               TabIndex        =   268
               Top             =   4200
               Width           =   1395
            End
            Begin VB.Label LNiveaux 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H00C0E0FF&
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
               ForeColor       =   &H80000008&
               Height          =   315
               Index           =   13
               Left            =   8520
               TabIndex        =   267
               Top             =   2100
               Width           =   1395
            End
            Begin VB.Label LNiveaux 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H00C0E0FF&
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
               ForeColor       =   &H80000008&
               Height          =   315
               Index           =   12
               Left            =   8520
               TabIndex        =   266
               Top             =   840
               Visible         =   0   'False
               Width           =   1395
            End
            Begin VB.Label LTemperatures 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H00C0E0FF&
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
               ForeColor       =   &H80000008&
               Height          =   315
               Index           =   18
               Left            =   7020
               TabIndex        =   248
               Top             =   6915
               Visible         =   0   'False
               Width           =   1395
            End
            Begin VB.Image IEtatsPostes 
               Appearance      =   0  'Flat
               BorderStyle     =   1  'Fixed Single
               Height          =   315
               Index           =   40
               Left            =   4680
               Picture         =   "FSynoptique.frx":19D53
               Stretch         =   -1  'True
               Top             =   7560
               Width           =   435
            End
            Begin VB.Label LNomsPostes 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H00C0E0FF&
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
               ForeColor       =   &H80000008&
               Height          =   315
               Index           =   40
               Left            =   180
               TabIndex        =   228
               Top             =   7560
               Width           =   615
            End
            Begin VB.Label LLibellesPostes 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
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
               ForeColor       =   &H80000008&
               Height          =   315
               Index           =   40
               Left            =   780
               TabIndex        =   227
               Top             =   7560
               Width           =   3495
            End
            Begin VB.Label LNumCharges 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
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
               ForeColor       =   &H80000008&
               Height          =   315
               Index           =   40
               Left            =   4260
               TabIndex        =   226
               Top             =   7560
               Width           =   435
            End
            Begin VB.Label LNomsPostes 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H00C0E0FF&
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
               ForeColor       =   &H80000008&
               Height          =   315
               Index           =   39
               Left            =   180
               TabIndex        =   225
               Top             =   7140
               Width           =   615
            End
            Begin VB.Label LNumCharges 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
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
               ForeColor       =   &H80000008&
               Height          =   315
               Index           =   39
               Left            =   4260
               TabIndex        =   224
               Top             =   7140
               Width           =   435
            End
            Begin VB.Image IEtatsPostes 
               Appearance      =   0  'Flat
               BorderStyle     =   1  'Fixed Single
               Height          =   315
               Index           =   39
               Left            =   4680
               Picture         =   "FSynoptique.frx":1A119
               Stretch         =   -1  'True
               Top             =   7140
               Width           =   435
            End
            Begin VB.Label LLibellesPostes 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
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
               ForeColor       =   &H80000008&
               Height          =   315
               Index           =   39
               Left            =   780
               TabIndex        =   223
               Top             =   7140
               Width           =   3495
            End
            Begin VB.Label LTemperatures 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H00C0E0FF&
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
               ForeColor       =   &H80000008&
               Height          =   315
               Index           =   16
               Left            =   7020
               TabIndex        =   222
               Top             =   5880
               Width           =   1395
            End
            Begin VB.Label LNumCharges 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
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
               ForeColor       =   &H80000008&
               Height          =   315
               Index           =   35
               Left            =   4260
               TabIndex        =   221
               Top             =   5460
               Width           =   435
            End
            Begin VB.Label LNumCharges 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
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
               ForeColor       =   &H80000008&
               Height          =   315
               Index           =   34
               Left            =   4260
               TabIndex        =   220
               Top             =   5040
               Width           =   435
            End
            Begin VB.Label LNumCharges 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
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
               ForeColor       =   &H80000008&
               Height          =   315
               Index           =   33
               Left            =   4260
               TabIndex        =   219
               Top             =   4620
               Width           =   435
            End
            Begin VB.Label LNumCharges 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
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
               ForeColor       =   &H80000008&
               Height          =   315
               Index           =   32
               Left            =   4260
               TabIndex        =   218
               Top             =   4200
               Width           =   435
            End
            Begin VB.Label LNumCharges 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
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
               ForeColor       =   &H80000008&
               Height          =   315
               Index           =   31
               Left            =   4260
               TabIndex        =   217
               Top             =   3780
               Width           =   435
            End
            Begin VB.Label LNumCharges 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
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
               ForeColor       =   &H80000008&
               Height          =   315
               Index           =   30
               Left            =   4260
               TabIndex        =   216
               Top             =   3360
               Width           =   435
            End
            Begin VB.Label LNumCharges 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
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
               ForeColor       =   &H80000008&
               Height          =   315
               Index           =   29
               Left            =   4260
               TabIndex        =   215
               Top             =   2940
               Width           =   435
            End
            Begin VB.Label LNumCharges 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
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
               ForeColor       =   &H80000008&
               Height          =   315
               Index           =   28
               Left            =   4260
               TabIndex        =   214
               Top             =   2520
               Width           =   435
            End
            Begin VB.Label LNumCharges 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
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
               ForeColor       =   &H80000008&
               Height          =   315
               Index           =   27
               Left            =   4260
               TabIndex        =   213
               Top             =   2100
               Width           =   435
            End
            Begin VB.Label LNumCharges 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
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
               ForeColor       =   &H80000008&
               Height          =   315
               Index           =   26
               Left            =   4260
               TabIndex        =   212
               Top             =   1680
               Width           =   435
            End
            Begin VB.Label LNumCharges 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
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
               ForeColor       =   &H80000008&
               Height          =   315
               Index           =   25
               Left            =   4260
               TabIndex        =   211
               Top             =   1260
               Width           =   435
            End
            Begin VB.Label LNomsPostes 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H00C0E0FF&
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
               ForeColor       =   &H80000008&
               Height          =   315
               Index           =   35
               Left            =   180
               TabIndex        =   210
               Top             =   5460
               Width           =   615
            End
            Begin VB.Label LNomsPostes 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H00C0E0FF&
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
               ForeColor       =   &H80000008&
               Height          =   315
               Index           =   34
               Left            =   180
               TabIndex        =   209
               Top             =   5040
               Width           =   615
            End
            Begin VB.Label LNomsPostes 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H00C0E0FF&
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
               ForeColor       =   &H80000008&
               Height          =   315
               Index           =   33
               Left            =   180
               TabIndex        =   208
               Top             =   4620
               Width           =   615
            End
            Begin VB.Label LNomsPostes 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H00C0E0FF&
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
               ForeColor       =   &H80000008&
               Height          =   315
               Index           =   32
               Left            =   180
               TabIndex        =   207
               Top             =   4200
               Width           =   615
            End
            Begin VB.Label LNomsPostes 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H00C0E0FF&
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
               ForeColor       =   &H80000008&
               Height          =   315
               Index           =   31
               Left            =   180
               TabIndex        =   206
               Top             =   3780
               Width           =   615
            End
            Begin VB.Label LNomsPostes 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H00C0E0FF&
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
               ForeColor       =   &H80000008&
               Height          =   315
               Index           =   30
               Left            =   180
               TabIndex        =   205
               Top             =   3360
               Width           =   615
            End
            Begin VB.Label LNomsPostes 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H00C0E0FF&
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
               ForeColor       =   &H80000008&
               Height          =   315
               Index           =   29
               Left            =   180
               TabIndex        =   204
               Top             =   2940
               Width           =   615
            End
            Begin VB.Label LNomsPostes 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H00C0E0FF&
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
               ForeColor       =   &H80000008&
               Height          =   315
               Index           =   28
               Left            =   180
               TabIndex        =   203
               Top             =   2520
               Width           =   615
            End
            Begin VB.Label LNomsPostes 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H00C0E0FF&
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
               ForeColor       =   &H80000008&
               Height          =   315
               Index           =   27
               Left            =   180
               TabIndex        =   202
               Top             =   2100
               Width           =   615
            End
            Begin VB.Label LNomsPostes 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H00C0E0FF&
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
               ForeColor       =   &H80000008&
               Height          =   315
               Index           =   26
               Left            =   180
               TabIndex        =   201
               Top             =   1680
               Width           =   615
            End
            Begin VB.Label LNomsPostes 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H00C0E0FF&
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
               ForeColor       =   &H80000008&
               Height          =   315
               Index           =   25
               Left            =   180
               TabIndex        =   200
               Top             =   1260
               Width           =   615
            End
            Begin VB.Image IEtatsPostes 
               Appearance      =   0  'Flat
               BorderStyle     =   1  'Fixed Single
               Height          =   315
               Index           =   35
               Left            =   4680
               Picture         =   "FSynoptique.frx":1A4DF
               Stretch         =   -1  'True
               Top             =   5460
               Width           =   435
            End
            Begin VB.Image IEtatsPostes 
               Appearance      =   0  'Flat
               BorderStyle     =   1  'Fixed Single
               Height          =   315
               Index           =   34
               Left            =   4680
               Picture         =   "FSynoptique.frx":1A8A5
               Stretch         =   -1  'True
               Top             =   5040
               Width           =   435
            End
            Begin VB.Image IEtatsPostes 
               Appearance      =   0  'Flat
               BorderStyle     =   1  'Fixed Single
               Height          =   315
               Index           =   33
               Left            =   4680
               Picture         =   "FSynoptique.frx":1AC6B
               Stretch         =   -1  'True
               Top             =   4620
               Width           =   435
            End
            Begin VB.Image IEtatsPostes 
               Appearance      =   0  'Flat
               BorderStyle     =   1  'Fixed Single
               Height          =   315
               Index           =   32
               Left            =   4680
               Picture         =   "FSynoptique.frx":1B031
               Stretch         =   -1  'True
               Top             =   4200
               Width           =   435
            End
            Begin VB.Image IEtatsPostes 
               Appearance      =   0  'Flat
               BorderStyle     =   1  'Fixed Single
               Height          =   315
               Index           =   31
               Left            =   4680
               Picture         =   "FSynoptique.frx":1B3F7
               Stretch         =   -1  'True
               Top             =   3780
               Width           =   435
            End
            Begin VB.Image IEtatsPostes 
               Appearance      =   0  'Flat
               BorderStyle     =   1  'Fixed Single
               Height          =   315
               Index           =   30
               Left            =   4680
               Picture         =   "FSynoptique.frx":1B7BD
               Stretch         =   -1  'True
               Top             =   3360
               Width           =   435
            End
            Begin VB.Image IEtatsPostes 
               Appearance      =   0  'Flat
               BorderStyle     =   1  'Fixed Single
               Height          =   315
               Index           =   29
               Left            =   4680
               Picture         =   "FSynoptique.frx":1BB83
               Stretch         =   -1  'True
               Top             =   2940
               Width           =   435
            End
            Begin VB.Image IEtatsPostes 
               Appearance      =   0  'Flat
               BorderStyle     =   1  'Fixed Single
               Height          =   315
               Index           =   28
               Left            =   4680
               Picture         =   "FSynoptique.frx":1BF49
               Stretch         =   -1  'True
               Top             =   2520
               Width           =   435
            End
            Begin VB.Image IEtatsPostes 
               Appearance      =   0  'Flat
               BorderStyle     =   1  'Fixed Single
               Height          =   315
               Index           =   27
               Left            =   4680
               Picture         =   "FSynoptique.frx":1C30F
               Stretch         =   -1  'True
               Top             =   2100
               Width           =   435
            End
            Begin VB.Image IEtatsPostes 
               Appearance      =   0  'Flat
               BorderStyle     =   1  'Fixed Single
               Height          =   315
               Index           =   26
               Left            =   4680
               Picture         =   "FSynoptique.frx":1C6D5
               Stretch         =   -1  'True
               Top             =   1680
               Width           =   435
            End
            Begin VB.Image IEtatsPostes 
               Appearance      =   0  'Flat
               BorderStyle     =   1  'Fixed Single
               Height          =   315
               Index           =   25
               Left            =   4680
               Picture         =   "FSynoptique.frx":1CA9B
               Stretch         =   -1  'True
               Top             =   1260
               Width           =   435
            End
            Begin VB.Label LLibellesPostes 
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
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
               ForeColor       =   &H80000008&
               Height          =   315
               Index           =   35
               Left            =   780
               TabIndex        =   199
               Top             =   5460
               Width           =   3495
            End
            Begin VB.Label LLibellesPostes 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
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
               ForeColor       =   &H80000008&
               Height          =   315
               Index           =   34
               Left            =   780
               TabIndex        =   198
               Top             =   5040
               Width           =   3495
            End
            Begin VB.Label LLibellesPostes 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
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
               ForeColor       =   &H80000008&
               Height          =   315
               Index           =   33
               Left            =   780
               TabIndex        =   197
               Top             =   4620
               Width           =   3495
            End
            Begin VB.Label LLibellesPostes 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
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
               ForeColor       =   &H80000008&
               Height          =   315
               Index           =   32
               Left            =   780
               TabIndex        =   196
               Top             =   4200
               Width           =   3495
            End
            Begin VB.Label LLibellesPostes 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
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
               ForeColor       =   &H80000008&
               Height          =   315
               Index           =   31
               Left            =   780
               TabIndex        =   195
               Top             =   3780
               Width           =   3495
            End
            Begin VB.Label LLibellesPostes 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
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
               ForeColor       =   &H80000008&
               Height          =   315
               Index           =   30
               Left            =   780
               TabIndex        =   194
               Top             =   3360
               Width           =   3495
            End
            Begin VB.Label LLibellesPostes 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
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
               ForeColor       =   &H80000008&
               Height          =   315
               Index           =   29
               Left            =   780
               TabIndex        =   193
               Top             =   2940
               Width           =   3495
            End
            Begin VB.Label LLibellesPostes 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
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
               ForeColor       =   &H80000008&
               Height          =   315
               Index           =   28
               Left            =   780
               TabIndex        =   192
               Top             =   2520
               Width           =   3495
            End
            Begin VB.Label LLibellesPostes 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
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
               ForeColor       =   &H80000008&
               Height          =   315
               Index           =   27
               Left            =   780
               TabIndex        =   191
               Top             =   2100
               Width           =   3495
            End
            Begin VB.Label LLibellesPostes 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
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
               ForeColor       =   &H80000008&
               Height          =   315
               Index           =   26
               Left            =   780
               TabIndex        =   190
               Top             =   1680
               Width           =   3495
            End
            Begin VB.Label LLibellesPostes 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
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
               ForeColor       =   &H80000008&
               Height          =   315
               Index           =   25
               Left            =   780
               TabIndex        =   189
               Top             =   1260
               Width           =   3495
            End
            Begin VB.Label LTemperatures 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H00C0E0FF&
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
               ForeColor       =   &H80000008&
               Height          =   315
               Index           =   15
               Left            =   7020
               TabIndex        =   188
               Top             =   4620
               Width           =   1395
            End
            Begin VB.Label LTemperatures 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H00C0E0FF&
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
               ForeColor       =   &H80000008&
               Height          =   315
               Index           =   14
               Left            =   7020
               TabIndex        =   187
               Top             =   4200
               Width           =   1395
            End
            Begin VB.Label LTemperatures 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H00C0E0FF&
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
               ForeColor       =   &H80000008&
               Height          =   315
               Index           =   13
               Left            =   7020
               TabIndex        =   186
               Top             =   2100
               Width           =   1395
            End
            Begin VB.Label LTemperatures 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H00C0E0FF&
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
               ForeColor       =   &H80000008&
               Height          =   315
               Index           =   12
               Left            =   7020
               TabIndex        =   185
               Top             =   840
               Visible         =   0   'False
               Width           =   1395
            End
            Begin VB.Label LLibelles 
               Alignment       =   2  'Center
               BackColor       =   &H00FF0000&
               BorderStyle     =   1  'Fixed Single
               Caption         =   "APRES ANODISATION"
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
               Height          =   315
               Index           =   3
               Left            =   0
               TabIndex        =   184
               Top             =   0
               Width           =   10095
            End
            Begin VB.Label LNomsPostes 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H00C0E0FF&
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
               ForeColor       =   &H80000008&
               Height          =   315
               Index           =   36
               Left            =   180
               TabIndex        =   183
               Top             =   5880
               Width           =   615
            End
            Begin VB.Label LNomsPostes 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H00C0E0FF&
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
               ForeColor       =   &H80000008&
               Height          =   315
               Index           =   37
               Left            =   180
               TabIndex        =   182
               Top             =   6300
               Width           =   615
            End
            Begin VB.Label LNomsPostes 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H00C0E0FF&
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
               ForeColor       =   &H80000008&
               Height          =   315
               Index           =   38
               Left            =   180
               TabIndex        =   181
               Top             =   6720
               Width           =   615
            End
            Begin VB.Label LNumCharges 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
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
               ForeColor       =   &H80000008&
               Height          =   315
               Index           =   36
               Left            =   4260
               TabIndex        =   180
               Top             =   5880
               Width           =   435
            End
            Begin VB.Label LNumCharges 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
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
               ForeColor       =   &H80000008&
               Height          =   315
               Index           =   37
               Left            =   4260
               TabIndex        =   179
               Top             =   6300
               Width           =   435
            End
            Begin VB.Label LNumCharges 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
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
               ForeColor       =   &H80000008&
               Height          =   315
               Index           =   38
               Left            =   4260
               TabIndex        =   178
               Top             =   6720
               Width           =   435
            End
            Begin VB.Image IEtatsPostes 
               Appearance      =   0  'Flat
               BorderStyle     =   1  'Fixed Single
               Height          =   315
               Index           =   36
               Left            =   4680
               Picture         =   "FSynoptique.frx":1CE61
               Stretch         =   -1  'True
               Top             =   5880
               Width           =   435
            End
            Begin VB.Image IEtatsPostes 
               Appearance      =   0  'Flat
               BorderStyle     =   1  'Fixed Single
               Height          =   315
               Index           =   37
               Left            =   4680
               Picture         =   "FSynoptique.frx":1D227
               Stretch         =   -1  'True
               Top             =   6300
               Width           =   435
            End
            Begin VB.Image IEtatsPostes 
               Appearance      =   0  'Flat
               BorderStyle     =   1  'Fixed Single
               Height          =   315
               Index           =   38
               Left            =   4680
               Picture         =   "FSynoptique.frx":1D5ED
               Stretch         =   -1  'True
               Top             =   6720
               Width           =   435
            End
            Begin VB.Label LLibellesPostes 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
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
               ForeColor       =   &H80000008&
               Height          =   315
               Index           =   36
               Left            =   780
               TabIndex        =   177
               Top             =   5880
               Width           =   3495
            End
            Begin VB.Label LLibellesPostes 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
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
               ForeColor       =   &H80000008&
               Height          =   315
               Index           =   37
               Left            =   780
               TabIndex        =   176
               Top             =   6300
               Width           =   3495
            End
            Begin VB.Label LLibellesPostes 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
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
               ForeColor       =   &H80000008&
               Height          =   315
               Index           =   38
               Left            =   780
               TabIndex        =   175
               Top             =   6720
               Width           =   3495
            End
            Begin VB.Label LTemperatures 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H00C0E0FF&
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
               ForeColor       =   &H80000008&
               Height          =   315
               Index           =   17
               Left            =   7020
               TabIndex        =   174
               Top             =   6300
               Width           =   1395
            End
            Begin VB.Label LLibellesPostes 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
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
               ForeColor       =   &H80000008&
               Height          =   315
               Index           =   24
               Left            =   780
               TabIndex        =   173
               Top             =   840
               Width           =   3495
            End
            Begin VB.Image IEtatsPostes 
               Appearance      =   0  'Flat
               BorderStyle     =   1  'Fixed Single
               Height          =   315
               Index           =   24
               Left            =   4680
               Picture         =   "FSynoptique.frx":1D9B3
               Stretch         =   -1  'True
               Top             =   840
               Width           =   435
            End
            Begin VB.Label LNomsPostes 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H00C0E0FF&
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
               ForeColor       =   &H80000008&
               Height          =   315
               Index           =   24
               Left            =   180
               TabIndex        =   172
               Top             =   840
               Width           =   615
            End
            Begin VB.Label LNumCharges 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
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
               ForeColor       =   &H80000008&
               Height          =   315
               Index           =   24
               Left            =   4260
               TabIndex        =   171
               Top             =   840
               Width           =   435
            End
            Begin VB.Shape SDecoration 
               BackColor       =   &H00C0E0FF&
               BackStyle       =   1  'Opaque
               Height          =   540
               Index           =   3
               Left            =   5100
               Top             =   6810
               Width           =   675
            End
         End
         Begin Anodisation.ImageMask IMLegende 
            Height          =   360
            Left            =   21360
            TabIndex        =   153
            Top             =   9960
            Width           =   1050
            _ExtentX        =   1852
            _ExtentY        =   635
            Picture         =   "FSynoptique.frx":1DD79
            MaskPicture     =   "FSynoptique.frx":1F1AB
            MaskColor       =   16711935
         End
         Begin VB.Frame FDivers 
            Height          =   1215
            Left            =   21000
            TabIndex        =   103
            Top             =   6900
            Width           =   6795
            Begin VB.CommandButton CBSauvegardeEtatsPostes 
               BackColor       =   &H00C0C0FF&
               Caption         =   "Sauvegarde de l'états des postes"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   495
               Left            =   180
               Style           =   1  'Graphical
               TabIndex        =   105
               Top             =   480
               Width           =   6435
            End
            Begin VB.Label LLibelles 
               Alignment       =   2  'Center
               BackColor       =   &H00FF0000&
               BorderStyle     =   1  'Fixed Single
               Caption         =   "DIVERS"
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
               Height          =   315
               Index           =   6
               Left            =   0
               TabIndex        =   104
               Top             =   0
               Width           =   6795
            End
         End
         Begin VB.Frame FAnodisation 
            Height          =   3555
            Left            =   240
            TabIndex        =   81
            Top             =   7020
            Width           =   10095
            Begin Anodisation.ImageMask IMProgrammateurCycliqueCuves 
               Height          =   360
               Index           =   8
               Left            =   5280
               TabIndex        =   233
               Top             =   840
               Width           =   330
               _ExtentX        =   582
               _ExtentY        =   635
               Picture         =   "FSynoptique.frx":205DD
               MaskPicture     =   "FSynoptique.frx":20C8F
               MaskColor       =   16711935
            End
            Begin Anodisation.ImageMask IMProgrammateurCycliqueCuves 
               Height          =   360
               Index           =   9
               Left            =   5280
               TabIndex        =   234
               Top             =   1260
               Width           =   330
               _ExtentX        =   582
               _ExtentY        =   635
               Picture         =   "FSynoptique.frx":21341
               MaskPicture     =   "FSynoptique.frx":219F3
               MaskColor       =   16711935
            End
            Begin Anodisation.ImageMask IMProgrammateurCycliqueCuves 
               Height          =   360
               Index           =   10
               Left            =   5280
               TabIndex        =   235
               Top             =   1680
               Width           =   330
               _ExtentX        =   582
               _ExtentY        =   635
               Picture         =   "FSynoptique.frx":220A5
               MaskPicture     =   "FSynoptique.frx":22757
               MaskColor       =   16711935
            End
            Begin Anodisation.ImageMask IMProgrammateurCycliqueCuves 
               Height          =   360
               Index           =   11
               Left            =   5280
               TabIndex        =   236
               Top             =   2100
               Visible         =   0   'False
               Width           =   330
               _ExtentX        =   582
               _ExtentY        =   635
               Picture         =   "FSynoptique.frx":22E09
               MaskPicture     =   "FSynoptique.frx":234BB
               MaskColor       =   16711935
            End
            Begin VB.Label LManuAutoRegulation 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H00C0E0FF&
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
               ForeColor       =   &H80000008&
               Height          =   315
               Index           =   11
               Left            =   5760
               TabIndex        =   291
               Top             =   2100
               Visible         =   0   'False
               Width           =   1155
            End
            Begin VB.Label LManuAutoRegulation 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H00C0E0FF&
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
               ForeColor       =   &H80000008&
               Height          =   315
               Index           =   10
               Left            =   5760
               TabIndex        =   290
               Top             =   1680
               Width           =   1155
            End
            Begin VB.Label LManuAutoRegulation 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H00C0E0FF&
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
               ForeColor       =   &H80000008&
               Height          =   315
               Index           =   9
               Left            =   5760
               TabIndex        =   289
               Top             =   1260
               Width           =   1155
            End
            Begin VB.Label LManuAutoRegulation 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H00C0E0FF&
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
               ForeColor       =   &H80000008&
               Height          =   315
               Index           =   8
               Left            =   5760
               TabIndex        =   288
               Top             =   840
               Width           =   1155
            End
            Begin VB.Label LLibelles 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H00FF0000&
               BorderStyle     =   1  'Fixed Single
               Caption         =   "Régulation"
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
               Height          =   315
               Index           =   30
               Left            =   5760
               TabIndex        =   279
               Top             =   420
               Width           =   1155
            End
            Begin VB.Label LLibelles 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H00FF0000&
               BorderStyle     =   1  'Fixed Single
               Caption         =   "Températures"
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
               Height          =   315
               Index           =   26
               Left            =   7020
               TabIndex        =   275
               Top             =   420
               Width           =   1395
            End
            Begin VB.Label LLibelles 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H00FF0000&
               BorderStyle     =   1  'Fixed Single
               Caption         =   "Niveaux"
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
               Height          =   315
               Index           =   24
               Left            =   8520
               TabIndex        =   274
               Top             =   420
               Width           =   1395
            End
            Begin VB.Label LNiveaux 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H00C0E0FF&
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
               ForeColor       =   &H80000008&
               Height          =   315
               Index           =   11
               Left            =   8520
               TabIndex        =   265
               Top             =   2100
               Visible         =   0   'False
               Width           =   1395
            End
            Begin VB.Label LNiveaux 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H00C0E0FF&
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
               ForeColor       =   &H80000008&
               Height          =   315
               Index           =   10
               Left            =   8520
               TabIndex        =   264
               Top             =   1680
               Width           =   1395
            End
            Begin VB.Label LNiveaux 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H00C0E0FF&
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
               ForeColor       =   &H80000008&
               Height          =   315
               Index           =   9
               Left            =   8520
               TabIndex        =   263
               Top             =   1260
               Width           =   1395
            End
            Begin VB.Label LNiveaux 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H00C0E0FF&
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
               ForeColor       =   &H80000008&
               Height          =   315
               Index           =   8
               Left            =   8520
               TabIndex        =   262
               Top             =   840
               Width           =   1395
            End
            Begin VB.Label LTemperatures 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H00C0E0FF&
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
               ForeColor       =   &H80000008&
               Height          =   315
               Index           =   11
               Left            =   7020
               TabIndex        =   240
               Top             =   2100
               Visible         =   0   'False
               Width           =   1395
            End
            Begin VB.Label LTemperatures 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H00C0E0FF&
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
               ForeColor       =   &H80000008&
               Height          =   315
               Index           =   8
               Left            =   7020
               TabIndex        =   239
               Top             =   840
               Width           =   1395
            End
            Begin VB.Label LTemperatures 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H00C0E0FF&
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
               ForeColor       =   &H80000008&
               Height          =   315
               Index           =   9
               Left            =   7020
               TabIndex        =   102
               Top             =   1260
               Width           =   1395
            End
            Begin VB.Label LTemperatures 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H00C0E0FF&
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
               ForeColor       =   &H80000008&
               Height          =   315
               Index           =   10
               Left            =   7020
               TabIndex        =   101
               Top             =   1680
               Width           =   1395
            End
            Begin VB.Label LLibellesPostes 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
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
               ForeColor       =   &H80000008&
               Height          =   315
               Index           =   18
               Left            =   780
               TabIndex        =   100
               Top             =   840
               Width           =   3495
            End
            Begin VB.Label LLibellesPostes 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
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
               ForeColor       =   &H80000008&
               Height          =   315
               Index           =   19
               Left            =   780
               TabIndex        =   99
               Top             =   1260
               Width           =   3495
            End
            Begin VB.Label LLibellesPostes 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
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
               ForeColor       =   &H80000008&
               Height          =   315
               Index           =   20
               Left            =   780
               TabIndex        =   98
               Top             =   1680
               Width           =   3495
            End
            Begin VB.Label LLibellesPostes 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
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
               ForeColor       =   &H80000008&
               Height          =   315
               Index           =   21
               Left            =   780
               TabIndex        =   97
               Top             =   2100
               Width           =   3495
            End
            Begin VB.Label LLibellesPostes 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
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
               ForeColor       =   &H80000008&
               Height          =   315
               Index           =   22
               Left            =   780
               TabIndex        =   96
               Top             =   2520
               Width           =   3495
            End
            Begin VB.Label LLibellesPostes 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
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
               ForeColor       =   &H80000008&
               Height          =   315
               Index           =   23
               Left            =   780
               TabIndex        =   95
               Top             =   2940
               Width           =   3495
            End
            Begin VB.Image IEtatsPostes 
               Appearance      =   0  'Flat
               BorderStyle     =   1  'Fixed Single
               Height          =   315
               Index           =   18
               Left            =   4680
               Picture         =   "FSynoptique.frx":23B6D
               Stretch         =   -1  'True
               Top             =   840
               Width           =   435
            End
            Begin VB.Image IEtatsPostes 
               Appearance      =   0  'Flat
               BorderStyle     =   1  'Fixed Single
               Height          =   315
               Index           =   19
               Left            =   4680
               Picture         =   "FSynoptique.frx":23F33
               Stretch         =   -1  'True
               Top             =   1260
               Width           =   435
            End
            Begin VB.Image IEtatsPostes 
               Appearance      =   0  'Flat
               BorderStyle     =   1  'Fixed Single
               Height          =   315
               Index           =   20
               Left            =   4680
               Picture         =   "FSynoptique.frx":242F9
               Stretch         =   -1  'True
               Top             =   1680
               Width           =   435
            End
            Begin VB.Image IEtatsPostes 
               Appearance      =   0  'Flat
               BorderStyle     =   1  'Fixed Single
               Height          =   315
               Index           =   21
               Left            =   4680
               Picture         =   "FSynoptique.frx":246BF
               Stretch         =   -1  'True
               Top             =   2100
               Width           =   435
            End
            Begin VB.Image IEtatsPostes 
               Appearance      =   0  'Flat
               BorderStyle     =   1  'Fixed Single
               Height          =   315
               Index           =   22
               Left            =   4680
               Picture         =   "FSynoptique.frx":24A85
               Stretch         =   -1  'True
               Top             =   2520
               Width           =   435
            End
            Begin VB.Image IEtatsPostes 
               Appearance      =   0  'Flat
               BorderStyle     =   1  'Fixed Single
               Height          =   315
               Index           =   23
               Left            =   4680
               Picture         =   "FSynoptique.frx":24E4B
               Stretch         =   -1  'True
               Top             =   2940
               Width           =   435
            End
            Begin VB.Label LNomsPostes 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H00C0E0FF&
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
               ForeColor       =   &H80000008&
               Height          =   315
               Index           =   18
               Left            =   180
               TabIndex        =   94
               Top             =   840
               Width           =   615
            End
            Begin VB.Label LNomsPostes 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H00C0E0FF&
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
               ForeColor       =   &H80000008&
               Height          =   315
               Index           =   19
               Left            =   180
               TabIndex        =   93
               Top             =   1260
               Width           =   615
            End
            Begin VB.Label LNomsPostes 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H00C0E0FF&
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
               ForeColor       =   &H80000008&
               Height          =   315
               Index           =   20
               Left            =   180
               TabIndex        =   92
               Top             =   1680
               Width           =   615
            End
            Begin VB.Label LNomsPostes 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H00C0E0FF&
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
               ForeColor       =   &H80000008&
               Height          =   315
               Index           =   21
               Left            =   180
               TabIndex        =   91
               Top             =   2100
               Width           =   615
            End
            Begin VB.Label LNomsPostes 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H00C0E0FF&
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
               ForeColor       =   &H80000008&
               Height          =   315
               Index           =   22
               Left            =   180
               TabIndex        =   90
               Top             =   2520
               Width           =   615
            End
            Begin VB.Label LNomsPostes 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H00C0E0FF&
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
               ForeColor       =   &H80000008&
               Height          =   315
               Index           =   23
               Left            =   180
               TabIndex        =   89
               Top             =   2940
               Width           =   615
            End
            Begin VB.Label LNumCharges 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
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
               ForeColor       =   &H80000008&
               Height          =   315
               Index           =   18
               Left            =   4260
               TabIndex        =   88
               Top             =   840
               Width           =   435
            End
            Begin VB.Label LNumCharges 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
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
               ForeColor       =   &H80000008&
               Height          =   315
               Index           =   19
               Left            =   4260
               TabIndex        =   87
               Top             =   1260
               Width           =   435
            End
            Begin VB.Label LNumCharges 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
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
               ForeColor       =   &H80000008&
               Height          =   315
               Index           =   20
               Left            =   4260
               TabIndex        =   86
               Top             =   1680
               Width           =   435
            End
            Begin VB.Label LNumCharges 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
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
               ForeColor       =   &H80000008&
               Height          =   315
               Index           =   21
               Left            =   4260
               TabIndex        =   85
               Top             =   2100
               Width           =   435
            End
            Begin VB.Label LNumCharges 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
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
               ForeColor       =   &H80000008&
               Height          =   315
               Index           =   22
               Left            =   4260
               TabIndex        =   84
               Top             =   2520
               Width           =   435
            End
            Begin VB.Label LNumCharges 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
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
               ForeColor       =   &H80000008&
               Height          =   315
               Index           =   23
               Left            =   4260
               TabIndex        =   83
               Top             =   2940
               Width           =   435
            End
            Begin VB.Label LLibelles 
               Alignment       =   2  'Center
               BackColor       =   &H00FF0000&
               BorderStyle     =   1  'Fixed Single
               Caption         =   "ANODISATION"
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
               Height          =   315
               Index           =   5
               Left            =   0
               TabIndex        =   82
               Top             =   0
               Width           =   10095
            End
         End
         Begin VB.Frame FAvantAnodisation 
            Height          =   6735
            Left            =   240
            TabIndex        =   28
            Top             =   240
            Width           =   10095
            Begin Anodisation.ImageMask IMProgrammateurCycliqueCuves 
               Height          =   375
               Index           =   1
               Left            =   5280
               TabIndex        =   154
               Top             =   840
               Width           =   375
               _ExtentX        =   582
               _ExtentY        =   635
               Picture         =   "FSynoptique.frx":25211
               MaskPicture     =   "FSynoptique.frx":258C3
               MaskColor       =   16711935
            End
            Begin Anodisation.ImageMask IMProgrammateurCycliqueCuves 
               Height          =   360
               Index           =   2
               Left            =   5280
               TabIndex        =   155
               Top             =   1245
               Width           =   330
               _ExtentX        =   582
               _ExtentY        =   635
               Picture         =   "FSynoptique.frx":25F75
               MaskPicture     =   "FSynoptique.frx":26627
               MaskColor       =   16711935
            End
            Begin Anodisation.ImageMask IMProgrammateurCycliqueCuves 
               Height          =   360
               Index           =   6
               Left            =   5280
               TabIndex        =   169
               Top             =   3840
               Visible         =   0   'False
               Width           =   330
               _ExtentX        =   582
               _ExtentY        =   635
               Picture         =   "FSynoptique.frx":26CD9
               MaskPicture     =   "FSynoptique.frx":2738B
               MaskColor       =   16711935
            End
            Begin Anodisation.ImageMask IMProgrammateurCycliqueCuves 
               Height          =   360
               Index           =   3
               Left            =   5280
               TabIndex        =   229
               Top             =   1680
               Visible         =   0   'False
               Width           =   330
               _ExtentX        =   582
               _ExtentY        =   635
               Picture         =   "FSynoptique.frx":27A3D
               MaskPicture     =   "FSynoptique.frx":280EF
               MaskColor       =   16711935
            End
            Begin Anodisation.ImageMask IMProgrammateurCycliqueCuves 
               Height          =   360
               Index           =   4
               Left            =   5280
               TabIndex        =   230
               Top             =   2520
               Width           =   330
               _ExtentX        =   582
               _ExtentY        =   635
               Picture         =   "FSynoptique.frx":287A1
               MaskPicture     =   "FSynoptique.frx":28E53
               MaskColor       =   16711935
            End
            Begin Anodisation.ImageMask IMProgrammateurCycliqueCuves 
               Height          =   360
               Index           =   5
               Left            =   5280
               TabIndex        =   231
               Top             =   3420
               Visible         =   0   'False
               Width           =   330
               _ExtentX        =   582
               _ExtentY        =   635
               Picture         =   "FSynoptique.frx":29505
               MaskPicture     =   "FSynoptique.frx":29BB7
               MaskColor       =   16711935
            End
            Begin Anodisation.ImageMask IMProgrammateurCycliqueCuves 
               Height          =   360
               Index           =   7
               Left            =   5280
               TabIndex        =   232
               Top             =   4260
               Width           =   330
               _ExtentX        =   582
               _ExtentY        =   635
               Picture         =   "FSynoptique.frx":2A269
               MaskPicture     =   "FSynoptique.frx":2A91B
               MaskColor       =   16711935
            End
            Begin VB.Label LNomsPostes 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H00C0E0FF&
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
               ForeColor       =   &H80000008&
               Height          =   315
               Index           =   4
               Left            =   180
               TabIndex        =   320
               Top             =   2160
               Width           =   615
            End
            Begin VB.Label LNumCharges 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
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
               ForeColor       =   &H80000008&
               Height          =   315
               Index           =   4
               Left            =   4260
               TabIndex        =   319
               Top             =   2160
               Width           =   435
            End
            Begin VB.Image IEtatsPostes 
               Appearance      =   0  'Flat
               BorderStyle     =   1  'Fixed Single
               Height          =   315
               Index           =   4
               Left            =   4680
               Picture         =   "FSynoptique.frx":2AFCD
               Stretch         =   -1  'True
               Top             =   2160
               Width           =   435
            End
            Begin VB.Label LLibellesPostes 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
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
               ForeColor       =   &H80000008&
               Height          =   315
               Index           =   4
               Left            =   780
               TabIndex        =   318
               Top             =   2160
               Width           =   3495
            End
            Begin VB.Line LDecoration 
               BorderColor     =   &H000000FF&
               BorderWidth     =   3
               Index           =   3
               X1              =   5445
               X2              =   5445
               Y1              =   3720
               Y2              =   4560
            End
            Begin VB.Label LManuAutoRegulation 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H00C0E0FF&
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
               ForeColor       =   &H80000008&
               Height          =   315
               Index           =   7
               Left            =   5760
               TabIndex        =   287
               Top             =   4260
               Width           =   1155
            End
            Begin VB.Label LManuAutoRegulation 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H00C0E0FF&
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
               ForeColor       =   &H80000008&
               Height          =   315
               Index           =   6
               Left            =   5760
               TabIndex        =   286
               Top             =   3840
               Visible         =   0   'False
               Width           =   1155
            End
            Begin VB.Label LManuAutoRegulation 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H00C0E0FF&
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
               ForeColor       =   &H80000008&
               Height          =   315
               Index           =   5
               Left            =   5760
               TabIndex        =   285
               Top             =   3420
               Visible         =   0   'False
               Width           =   1155
            End
            Begin VB.Label LManuAutoRegulation 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H00C0E0FF&
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
               ForeColor       =   &H80000008&
               Height          =   315
               Index           =   4
               Left            =   5760
               TabIndex        =   284
               Top             =   2580
               Visible         =   0   'False
               Width           =   1155
            End
            Begin VB.Label LManuAutoRegulation 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H00C0E0FF&
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
               ForeColor       =   &H80000008&
               Height          =   315
               Index           =   3
               Left            =   5760
               TabIndex        =   283
               Top             =   1680
               Visible         =   0   'False
               Width           =   1155
            End
            Begin VB.Label LManuAutoRegulation 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H00C0E0FF&
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
               ForeColor       =   &H80000008&
               Height          =   315
               Index           =   2
               Left            =   5760
               TabIndex        =   282
               Top             =   1260
               Width           =   1155
            End
            Begin VB.Label LManuAutoRegulation 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H00C0E0FF&
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
               ForeColor       =   &H80000008&
               Height          =   315
               Index           =   1
               Left            =   5760
               TabIndex        =   281
               Top             =   840
               Width           =   1155
            End
            Begin VB.Label LLibelles 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H00FF0000&
               BorderStyle     =   1  'Fixed Single
               Caption         =   "Régulation"
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
               Height          =   315
               Index           =   29
               Left            =   5760
               TabIndex        =   278
               Top             =   420
               Width           =   1155
            End
            Begin VB.Label LLibelles 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H00FF0000&
               BorderStyle     =   1  'Fixed Single
               Caption         =   "Niveaux"
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
               Height          =   315
               Index           =   25
               Left            =   8520
               TabIndex        =   273
               Top             =   420
               Width           =   1395
            End
            Begin VB.Label LLibelles 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H00FF0000&
               BorderStyle     =   1  'Fixed Single
               Caption         =   "Températures"
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
               Height          =   315
               Index           =   23
               Left            =   7020
               TabIndex        =   272
               Top             =   420
               Width           =   1395
            End
            Begin VB.Label LNiveaux 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H00C0E0FF&
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
               ForeColor       =   &H80000008&
               Height          =   315
               Index           =   7
               Left            =   8520
               TabIndex        =   261
               Top             =   4260
               Width           =   1395
            End
            Begin VB.Label LNiveaux 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H00C0E0FF&
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
               ForeColor       =   &H80000008&
               Height          =   315
               Index           =   6
               Left            =   8520
               TabIndex        =   260
               Top             =   3840
               Visible         =   0   'False
               Width           =   1395
            End
            Begin VB.Label LNiveaux 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H00C0E0FF&
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
               ForeColor       =   &H80000008&
               Height          =   315
               Index           =   5
               Left            =   8520
               TabIndex        =   259
               Top             =   3420
               Visible         =   0   'False
               Width           =   1395
            End
            Begin VB.Label LNiveaux 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H00C0E0FF&
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
               ForeColor       =   &H80000008&
               Height          =   315
               Index           =   4
               Left            =   8520
               TabIndex        =   258
               Top             =   2580
               Visible         =   0   'False
               Width           =   1395
            End
            Begin VB.Label LNiveaux 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H00C0E0FF&
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
               ForeColor       =   &H80000008&
               Height          =   315
               Index           =   3
               Left            =   8520
               TabIndex        =   257
               Top             =   1680
               Visible         =   0   'False
               Width           =   1395
            End
            Begin VB.Label LNiveaux 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H00C0E0FF&
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
               ForeColor       =   &H80000008&
               Height          =   315
               Index           =   2
               Left            =   8520
               TabIndex        =   256
               Top             =   1260
               Width           =   1395
            End
            Begin VB.Label LNiveaux 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H00C0E0FF&
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
               ForeColor       =   &H80000008&
               Height          =   315
               Index           =   1
               Left            =   8520
               TabIndex        =   255
               Top             =   840
               Width           =   1395
            End
            Begin VB.Label LTemperatures 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H00C0E0FF&
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
               ForeColor       =   &H80000008&
               Height          =   315
               Index           =   7
               Left            =   7020
               TabIndex        =   242
               Top             =   4260
               Width           =   1395
            End
            Begin VB.Label LTemperatures 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H00C0E0FF&
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
               ForeColor       =   &H80000008&
               Height          =   315
               Index           =   6
               Left            =   7020
               TabIndex        =   241
               Top             =   3840
               Visible         =   0   'False
               Width           =   1395
            End
            Begin VB.Label LLibellesPostes 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
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
               ForeColor       =   &H80000008&
               Height          =   315
               Index           =   6
               Left            =   780
               TabIndex        =   80
               Top             =   1260
               Width           =   3495
            End
            Begin VB.Label LLibellesPostes 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
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
               ForeColor       =   &H80000008&
               Height          =   315
               Index           =   5
               Left            =   780
               TabIndex        =   79
               Top             =   840
               Width           =   3495
            End
            Begin VB.Image IEtatsPostes 
               Appearance      =   0  'Flat
               BorderStyle     =   1  'Fixed Single
               Height          =   315
               Index           =   6
               Left            =   4680
               Picture         =   "FSynoptique.frx":2B393
               Stretch         =   -1  'True
               Top             =   1260
               Width           =   435
            End
            Begin VB.Image IEtatsPostes 
               Appearance      =   0  'Flat
               BorderStyle     =   1  'Fixed Single
               Height          =   315
               Index           =   5
               Left            =   4680
               Picture         =   "FSynoptique.frx":2B759
               Stretch         =   -1  'True
               Top             =   840
               Width           =   435
            End
            Begin VB.Label LNumCharges 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
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
               ForeColor       =   &H80000008&
               Height          =   315
               Index           =   6
               Left            =   4260
               TabIndex        =   78
               Top             =   1260
               Width           =   435
            End
            Begin VB.Label LNumCharges 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
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
               ForeColor       =   &H80000008&
               Height          =   315
               Index           =   5
               Left            =   4260
               TabIndex        =   77
               Top             =   840
               Width           =   435
            End
            Begin VB.Label LNomsPostes 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H00C0E0FF&
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
               ForeColor       =   &H80000008&
               Height          =   315
               Index           =   6
               Left            =   180
               TabIndex        =   76
               Top             =   1260
               Width           =   615
            End
            Begin VB.Label LNomsPostes 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H00C0E0FF&
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
               ForeColor       =   &H80000008&
               Height          =   315
               Index           =   5
               Left            =   180
               TabIndex        =   75
               Top             =   840
               Width           =   615
            End
            Begin VB.Label LLibelles 
               Alignment       =   2  'Center
               BackColor       =   &H00FF0000&
               BorderStyle     =   1  'Fixed Single
               Caption         =   "AVANT ANODISATION"
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
               Height          =   315
               Index           =   2
               Left            =   0
               TabIndex        =   67
               Top             =   0
               Width           =   10095
            End
            Begin VB.Label LTemperatures 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H00C0E0FF&
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
               ForeColor       =   &H80000008&
               Height          =   315
               Index           =   1
               Left            =   7020
               TabIndex        =   66
               Top             =   840
               Width           =   1395
            End
            Begin VB.Label LTemperatures 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H00C0E0FF&
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
               ForeColor       =   &H80000008&
               Height          =   315
               Index           =   2
               Left            =   7020
               TabIndex        =   65
               Top             =   1260
               Width           =   1395
            End
            Begin VB.Label LTemperatures 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H00C0E0FF&
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
               ForeColor       =   &H80000008&
               Height          =   315
               Index           =   3
               Left            =   7020
               TabIndex        =   64
               Top             =   1680
               Visible         =   0   'False
               Width           =   1395
            End
            Begin VB.Label LTemperatures 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H00C0E0FF&
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
               ForeColor       =   &H80000008&
               Height          =   315
               Index           =   4
               Left            =   7020
               TabIndex        =   63
               Top             =   2580
               Visible         =   0   'False
               Width           =   1395
            End
            Begin VB.Label LTemperatures 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H00C0E0FF&
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
               ForeColor       =   &H80000008&
               Height          =   315
               Index           =   5
               Left            =   7020
               TabIndex        =   62
               Top             =   3420
               Visible         =   0   'False
               Width           =   1395
            End
            Begin VB.Label LLibellesPostes 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
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
               ForeColor       =   &H80000008&
               Height          =   315
               Index           =   7
               Left            =   780
               TabIndex        =   61
               Top             =   1680
               Width           =   3495
            End
            Begin VB.Label LLibellesPostes 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
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
               ForeColor       =   &H80000008&
               Height          =   315
               Index           =   8
               Left            =   780
               TabIndex        =   60
               Top             =   2580
               Width           =   3495
            End
            Begin VB.Label LLibellesPostes 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
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
               ForeColor       =   &H80000008&
               Height          =   315
               Index           =   9
               Left            =   780
               TabIndex        =   59
               Top             =   3000
               Width           =   3495
            End
            Begin VB.Label LLibellesPostes 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
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
               ForeColor       =   &H80000008&
               Height          =   315
               Index           =   10
               Left            =   780
               TabIndex        =   58
               Top             =   3420
               Width           =   3495
            End
            Begin VB.Label LLibellesPostes 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
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
               ForeColor       =   &H80000008&
               Height          =   315
               Index           =   11
               Left            =   780
               TabIndex        =   57
               Top             =   3840
               Width           =   3495
            End
            Begin VB.Label LLibellesPostes 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
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
               ForeColor       =   &H80000008&
               Height          =   315
               Index           =   12
               Left            =   780
               TabIndex        =   56
               Top             =   4260
               Width           =   3495
            End
            Begin VB.Label LLibellesPostes 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
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
               ForeColor       =   &H80000008&
               Height          =   315
               Index           =   13
               Left            =   780
               TabIndex        =   55
               Top             =   4680
               Width           =   3495
            End
            Begin VB.Label LLibellesPostes 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
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
               ForeColor       =   &H80000008&
               Height          =   315
               Index           =   14
               Left            =   780
               TabIndex        =   54
               Top             =   5100
               Width           =   3495
            End
            Begin VB.Label LLibellesPostes 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
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
               ForeColor       =   &H80000008&
               Height          =   315
               Index           =   15
               Left            =   780
               TabIndex        =   53
               Top             =   5520
               Width           =   3495
            End
            Begin VB.Label LLibellesPostes 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
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
               ForeColor       =   &H80000008&
               Height          =   315
               Index           =   16
               Left            =   780
               TabIndex        =   52
               Top             =   5940
               Width           =   3495
            End
            Begin VB.Label LLibellesPostes 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
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
               ForeColor       =   &H80000008&
               Height          =   315
               Index           =   17
               Left            =   780
               TabIndex        =   51
               Top             =   6360
               Width           =   3495
            End
            Begin VB.Image IEtatsPostes 
               Appearance      =   0  'Flat
               BorderStyle     =   1  'Fixed Single
               Height          =   315
               Index           =   7
               Left            =   4680
               Picture         =   "FSynoptique.frx":2BB1F
               Stretch         =   -1  'True
               Top             =   1680
               Width           =   435
            End
            Begin VB.Image IEtatsPostes 
               Appearance      =   0  'Flat
               BorderStyle     =   1  'Fixed Single
               Height          =   315
               Index           =   8
               Left            =   4680
               Picture         =   "FSynoptique.frx":2BEE5
               Stretch         =   -1  'True
               Top             =   2580
               Width           =   435
            End
            Begin VB.Image IEtatsPostes 
               Appearance      =   0  'Flat
               BorderStyle     =   1  'Fixed Single
               Height          =   315
               Index           =   9
               Left            =   4680
               Picture         =   "FSynoptique.frx":2C2AB
               Stretch         =   -1  'True
               Top             =   3000
               Width           =   435
            End
            Begin VB.Image IEtatsPostes 
               Appearance      =   0  'Flat
               BorderStyle     =   1  'Fixed Single
               Height          =   315
               Index           =   10
               Left            =   4680
               Picture         =   "FSynoptique.frx":2C671
               Stretch         =   -1  'True
               Top             =   3420
               Width           =   435
            End
            Begin VB.Image IEtatsPostes 
               Appearance      =   0  'Flat
               BorderStyle     =   1  'Fixed Single
               Height          =   315
               Index           =   11
               Left            =   4680
               Picture         =   "FSynoptique.frx":2CA37
               Stretch         =   -1  'True
               Top             =   3840
               Width           =   435
            End
            Begin VB.Image IEtatsPostes 
               Appearance      =   0  'Flat
               BorderStyle     =   1  'Fixed Single
               Height          =   315
               Index           =   12
               Left            =   4680
               Picture         =   "FSynoptique.frx":2CDFD
               Stretch         =   -1  'True
               Top             =   4260
               Width           =   435
            End
            Begin VB.Image IEtatsPostes 
               Appearance      =   0  'Flat
               BorderStyle     =   1  'Fixed Single
               Height          =   315
               Index           =   13
               Left            =   4680
               Picture         =   "FSynoptique.frx":2D1C3
               Stretch         =   -1  'True
               Top             =   4680
               Width           =   435
            End
            Begin VB.Image IEtatsPostes 
               Appearance      =   0  'Flat
               BorderStyle     =   1  'Fixed Single
               Height          =   315
               Index           =   14
               Left            =   4680
               Picture         =   "FSynoptique.frx":2D589
               Stretch         =   -1  'True
               Top             =   5100
               Width           =   435
            End
            Begin VB.Image IEtatsPostes 
               Appearance      =   0  'Flat
               BorderStyle     =   1  'Fixed Single
               Height          =   315
               Index           =   15
               Left            =   4680
               Picture         =   "FSynoptique.frx":2D94F
               Stretch         =   -1  'True
               Top             =   5520
               Width           =   435
            End
            Begin VB.Image IEtatsPostes 
               Appearance      =   0  'Flat
               BorderStyle     =   1  'Fixed Single
               Height          =   315
               Index           =   16
               Left            =   4680
               Picture         =   "FSynoptique.frx":2DD15
               Stretch         =   -1  'True
               Top             =   5940
               Width           =   435
            End
            Begin VB.Image IEtatsPostes 
               Appearance      =   0  'Flat
               BorderStyle     =   1  'Fixed Single
               Height          =   315
               Index           =   17
               Left            =   4680
               Picture         =   "FSynoptique.frx":2E0DB
               Stretch         =   -1  'True
               Top             =   6360
               Width           =   435
            End
            Begin VB.Label LNomsPostes 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H00C0E0FF&
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
               ForeColor       =   &H80000008&
               Height          =   315
               Index           =   7
               Left            =   180
               TabIndex        =   50
               Top             =   1680
               Width           =   615
            End
            Begin VB.Label LNomsPostes 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H00C0E0FF&
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
               ForeColor       =   &H80000008&
               Height          =   315
               Index           =   8
               Left            =   180
               TabIndex        =   49
               Top             =   2580
               Width           =   615
            End
            Begin VB.Label LNomsPostes 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H00C0E0FF&
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
               ForeColor       =   &H80000008&
               Height          =   315
               Index           =   9
               Left            =   180
               TabIndex        =   48
               Top             =   3000
               Width           =   615
            End
            Begin VB.Label LNomsPostes 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H00C0E0FF&
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
               ForeColor       =   &H80000008&
               Height          =   315
               Index           =   10
               Left            =   180
               TabIndex        =   47
               Top             =   3420
               Width           =   615
            End
            Begin VB.Label LNomsPostes 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H00C0E0FF&
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
               ForeColor       =   &H80000008&
               Height          =   315
               Index           =   11
               Left            =   180
               TabIndex        =   46
               Top             =   3840
               Width           =   615
            End
            Begin VB.Label LNomsPostes 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H00C0E0FF&
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
               ForeColor       =   &H80000008&
               Height          =   315
               Index           =   12
               Left            =   180
               TabIndex        =   45
               Top             =   4260
               Width           =   615
            End
            Begin VB.Label LNomsPostes 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H00C0E0FF&
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
               ForeColor       =   &H80000008&
               Height          =   315
               Index           =   13
               Left            =   180
               TabIndex        =   44
               Top             =   4680
               Width           =   615
            End
            Begin VB.Label LNomsPostes 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H00C0E0FF&
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
               ForeColor       =   &H80000008&
               Height          =   315
               Index           =   14
               Left            =   180
               TabIndex        =   43
               Top             =   5100
               Width           =   615
            End
            Begin VB.Label LNomsPostes 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H00C0E0FF&
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
               ForeColor       =   &H80000008&
               Height          =   315
               Index           =   15
               Left            =   180
               TabIndex        =   42
               Top             =   5520
               Width           =   615
            End
            Begin VB.Label LNomsPostes 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H00C0E0FF&
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
               ForeColor       =   &H80000008&
               Height          =   315
               Index           =   16
               Left            =   180
               TabIndex        =   41
               Top             =   5940
               Width           =   615
            End
            Begin VB.Label LNomsPostes 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H00C0E0FF&
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
               ForeColor       =   &H80000008&
               Height          =   315
               Index           =   17
               Left            =   180
               TabIndex        =   40
               Top             =   6360
               Width           =   615
            End
            Begin VB.Label LNumCharges 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
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
               ForeColor       =   &H80000008&
               Height          =   315
               Index           =   7
               Left            =   4260
               TabIndex        =   39
               Top             =   1680
               Width           =   435
            End
            Begin VB.Label LNumCharges 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
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
               ForeColor       =   &H80000008&
               Height          =   315
               Index           =   8
               Left            =   4260
               TabIndex        =   38
               Top             =   2580
               Width           =   435
            End
            Begin VB.Label LNumCharges 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
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
               ForeColor       =   &H80000008&
               Height          =   315
               Index           =   9
               Left            =   4260
               TabIndex        =   37
               Top             =   3000
               Width           =   435
            End
            Begin VB.Label LNumCharges 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
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
               ForeColor       =   &H80000008&
               Height          =   315
               Index           =   10
               Left            =   4260
               TabIndex        =   36
               Top             =   3420
               Width           =   435
            End
            Begin VB.Label LNumCharges 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
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
               ForeColor       =   &H80000008&
               Height          =   315
               Index           =   11
               Left            =   4260
               TabIndex        =   35
               Top             =   3840
               Width           =   435
            End
            Begin VB.Label LNumCharges 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
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
               ForeColor       =   &H80000008&
               Height          =   315
               Index           =   12
               Left            =   4260
               TabIndex        =   34
               Top             =   4260
               Width           =   435
            End
            Begin VB.Label LNumCharges 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
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
               ForeColor       =   &H80000008&
               Height          =   315
               Index           =   13
               Left            =   4260
               TabIndex        =   33
               Top             =   4680
               Width           =   435
            End
            Begin VB.Label LNumCharges 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
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
               ForeColor       =   &H80000008&
               Height          =   315
               Index           =   14
               Left            =   4260
               TabIndex        =   32
               Top             =   5100
               Width           =   435
            End
            Begin VB.Label LNumCharges 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
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
               ForeColor       =   &H80000008&
               Height          =   315
               Index           =   15
               Left            =   4260
               TabIndex        =   31
               Top             =   5520
               Width           =   435
            End
            Begin VB.Label LNumCharges 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
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
               ForeColor       =   &H80000008&
               Height          =   315
               Index           =   16
               Left            =   4260
               TabIndex        =   30
               Top             =   5940
               Width           =   435
            End
            Begin VB.Label LNumCharges 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
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
               ForeColor       =   &H80000008&
               Height          =   315
               Index           =   17
               Left            =   4260
               TabIndex        =   29
               Top             =   6360
               Width           =   435
            End
            Begin VB.Line LDecoration 
               BorderColor     =   &H000000FF&
               BorderWidth     =   3
               Index           =   2
               X1              =   5085
               X2              =   5445
               Y1              =   3570
               Y2              =   3570
            End
            Begin VB.Line LDecoration 
               BorderColor     =   &H000000FF&
               BorderWidth     =   3
               Index           =   4
               X1              =   5070
               X2              =   5835
               Y1              =   4440
               Y2              =   4440
            End
         End
         Begin VB.Frame FAnnexes 
            Height          =   6375
            Left            =   21000
            TabIndex        =   16
            Top             =   240
            Width           =   6795
            Begin VB.Label LLibelles 
               Alignment       =   2  'Center
               BackColor       =   &H00FF0000&
               BorderStyle     =   1  'Fixed Single
               Caption         =   "ANNEXES"
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
               Height          =   315
               Index           =   18
               Left            =   0
               TabIndex        =   27
               Top             =   0
               Width           =   6795
            End
            Begin VB.Label LLibellesAnnexes 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
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
               ForeColor       =   &H80000008&
               Height          =   315
               Index           =   0
               Left            =   180
               TabIndex        =   26
               Top             =   540
               Width           =   6015
            End
            Begin VB.Label LLibellesAnnexes 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
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
               ForeColor       =   &H80000008&
               Height          =   315
               Index           =   1
               Left            =   180
               TabIndex        =   25
               Top             =   960
               Width           =   6015
            End
            Begin VB.Label LLibellesAnnexes 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
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
               ForeColor       =   &H80000008&
               Height          =   315
               Index           =   2
               Left            =   180
               TabIndex        =   24
               Top             =   1380
               Width           =   6015
            End
            Begin VB.Label LLibellesAnnexes 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
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
               ForeColor       =   &H80000008&
               Height          =   315
               Index           =   3
               Left            =   180
               TabIndex        =   23
               Top             =   1800
               Width           =   6015
            End
            Begin VB.Label LEtatsAnnexes 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H0000FF00&
               BorderStyle     =   1  'Fixed Single
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
               Height          =   315
               Index           =   0
               Left            =   6180
               TabIndex        =   22
               Top             =   540
               Width           =   435
            End
            Begin VB.Label LEtatsAnnexes 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H0000FF00&
               BorderStyle     =   1  'Fixed Single
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
               Height          =   315
               Index           =   1
               Left            =   6180
               TabIndex        =   21
               Top             =   960
               Width           =   435
            End
            Begin VB.Label LEtatsAnnexes 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H0000FF00&
               BorderStyle     =   1  'Fixed Single
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
               Height          =   315
               Index           =   2
               Left            =   6180
               TabIndex        =   20
               Top             =   1380
               Width           =   435
            End
            Begin VB.Label LEtatsAnnexes 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H0000FF00&
               BorderStyle     =   1  'Fixed Single
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
               Height          =   315
               Index           =   3
               Left            =   6180
               TabIndex        =   19
               Top             =   1800
               Width           =   435
            End
            Begin VB.Label LLibellesAnnexes 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
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
               ForeColor       =   &H80000008&
               Height          =   315
               Index           =   4
               Left            =   180
               TabIndex        =   18
               Top             =   2220
               Width           =   6015
            End
            Begin VB.Label LEtatsAnnexes 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H0000FF00&
               BorderStyle     =   1  'Fixed Single
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
               Height          =   315
               Index           =   4
               Left            =   6180
               TabIndex        =   17
               Top             =   2220
               Width           =   435
            End
         End
         Begin VB.PictureBox PBLegende 
            Appearance      =   0  'Flat
            BackColor       =   &H0000FF00&
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   0
            Left            =   22080
            ScaleHeight     =   225
            ScaleWidth      =   285
            TabIndex        =   15
            Top             =   8700
            Width           =   315
         End
         Begin VB.PictureBox PBLegende 
            Appearance      =   0  'Flat
            BackColor       =   &H0000FF00&
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   1
            Left            =   22080
            Picture         =   "FSynoptique.frx":2E4A1
            ScaleHeight     =   225
            ScaleWidth      =   285
            TabIndex        =   14
            Top             =   9555
            Width           =   315
         End
         Begin VB.PictureBox PBLegende 
            Appearance      =   0  'Flat
            BackColor       =   &H000000FF&
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   2
            Left            =   22080
            ScaleHeight     =   225
            ScaleWidth      =   285
            TabIndex        =   13
            Top             =   9120
            Width           =   315
         End
         Begin VB.Label LLibellesLegende 
            BackStyle       =   0  'Transparent
            Caption         =   "Elément CONDAMNE (BOUTON DROIT de la SOURIS)"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   0
            Left            =   22560
            TabIndex        =   71
            Top             =   9540
            Width           =   5115
         End
         Begin VB.Label LLibellesLegende 
            BackStyle       =   0  'Transparent
            Caption         =   "Elément avec un DEFAUT"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   1
            Left            =   22560
            TabIndex        =   70
            Top             =   9135
            Width           =   2475
         End
         Begin VB.Label LLibellesLegende 
            BackStyle       =   0  'Transparent
            Caption         =   "Elément SANS DEFAUT"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   2
            Left            =   22560
            TabIndex        =   69
            Top             =   8700
            Width           =   2475
         End
         Begin VB.Label LLibellesLegende 
            BackStyle       =   0  'Transparent
            Caption         =   "ARRET, VEILLE, PRODUCTION d'un CHAUFFAGE"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   3
            Left            =   22560
            TabIndex        =   68
            Top             =   10020
            Width           =   4575
         End
         Begin VB.Shape SLegende 
            BackColor       =   &H00C0E0FF&
            BackStyle       =   1  'Opaque
            BorderWidth     =   2
            Height          =   2085
            Left            =   21000
            Shape           =   4  'Rounded Rectangle
            Top             =   8460
            Width           =   6795
         End
      End
      Begin VB.PictureBox PBGeneral 
         BackColor       =   &H00E0E0E0&
         Height          =   4635
         Left            =   420
         ScaleHeight     =   305
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   1877
         TabIndex        =   72
         Top             =   540
         Width           =   28215
         Begin VB.PictureBox PBPoidsSouleve 
            BackColor       =   &H00C0E0FF&
            Height          =   1935
            Index           =   0
            Left            =   24480
            ScaleHeight     =   1875
            ScaleWidth      =   3375
            TabIndex        =   148
            Top             =   2400
            Width           =   3435
            Begin VB.Label LLibelles 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H00FF0000&
               BorderStyle     =   1  'Fixed Single
               Caption         =   "POIDS SOULEVE - PONT 2"
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
               Height          =   315
               Index           =   20
               Left            =   240
               TabIndex        =   152
               Top             =   1020
               Width           =   2895
            End
            Begin VB.Label LPoidsSouleve 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               BorderStyle     =   1  'Fixed Single
               Caption         =   "-"
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
               Height          =   375
               Index           =   2
               Left            =   240
               TabIndex        =   151
               Top             =   1320
               Width           =   2895
            End
            Begin VB.Label LPoidsSouleve 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               BorderStyle     =   1  'Fixed Single
               Caption         =   "-"
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
               Height          =   375
               Index           =   1
               Left            =   240
               TabIndex        =   150
               Top             =   480
               Width           =   2895
            End
            Begin VB.Label LLibelles 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H00FF0000&
               BorderStyle     =   1  'Fixed Single
               Caption         =   "POIDS SOULEVE - PONT 1"
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
               Height          =   315
               Index           =   19
               Left            =   240
               TabIndex        =   149
               Top             =   180
               Width           =   2895
            End
         End
         Begin VB.PictureBox PBRedresseurSpectrocoloration 
            BackColor       =   &H00C0FFC0&
            Height          =   4155
            Left            =   3780
            ScaleHeight     =   4095
            ScaleWidth      =   2115
            TabIndex        =   142
            Top             =   180
            Width           =   2175
            Begin Anodisation.OCXRedresseur OCXRedresseurs 
               Height          =   3240
               Index           =   5
               Left            =   180
               TabIndex        =   144
               Top             =   720
               Width           =   1785
               _ExtentX        =   3149
               _ExtentY        =   5715
               Modele          =   1
            End
            Begin VB.Label LLibelles 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H00FF0000&
               BorderStyle     =   1  'Fixed Single
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
               Height          =   255
               Index           =   14
               Left            =   180
               TabIndex        =   147
               Top             =   480
               Width           =   1785
            End
            Begin VB.Label LLibelles 
               Alignment       =   2  'Center
               BackColor       =   &H00FF0000&
               BorderStyle     =   1  'Fixed Single
               Caption         =   "SPECTROCOLOR."
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
               Height          =   315
               Index           =   12
               Left            =   0
               TabIndex        =   143
               Top             =   0
               Width           =   2115
            End
         End
         Begin VB.PictureBox PBDechargement 
            BackColor       =   &H00C0E0FF&
            Height          =   4155
            Left            =   120
            ScaleHeight     =   4095
            ScaleWidth      =   3495
            TabIndex        =   123
            Top             =   180
            Width           =   3555
            Begin VB.TextBox TextInfo 
               Height          =   975
               Left            =   120
               TabIndex        =   315
               Text            =   "Text2"
               Top             =   2760
               Width           =   3135
            End
            Begin VB.TextBox CoordReelle 
               Height          =   375
               Left            =   120
               TabIndex        =   314
               Text            =   "Text1"
               Top             =   1800
               Width           =   3135
            End
            Begin VB.Label Label2 
               Caption         =   "INFOS:"
               Height          =   255
               Left            =   120
               TabIndex        =   317
               Top             =   2400
               Width           =   2295
            End
            Begin VB.Label Label1 
               Caption         =   "AUTOMATE:"
               Height          =   255
               Left            =   120
               TabIndex        =   316
               Top             =   1440
               Width           =   2175
            End
            Begin VB.Label LLibelles 
               Alignment       =   2  'Center
               BackColor       =   &H00FF0000&
               BorderStyle     =   1  'Fixed Single
               Caption         =   "DECHARGEMENT"
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
               Height          =   315
               Index           =   33
               Left            =   0
               TabIndex        =   306
               Top             =   0
               Width           =   3495
            End
            Begin VB.Label LNumCharges 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
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
               ForeColor       =   &H80000008&
               Height          =   315
               Index           =   41
               Left            =   2280
               TabIndex        =   133
               Top             =   540
               Width           =   435
            End
            Begin VB.Label LNumCharges 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
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
               ForeColor       =   &H80000008&
               Height          =   315
               Index           =   42
               Left            =   2280
               TabIndex        =   132
               Top             =   840
               Width           =   435
            End
            Begin VB.Label LLibellesPostes 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               BorderStyle     =   1  'Fixed Single
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
               ForeColor       =   &H80000008&
               Height          =   315
               Index           =   42
               Left            =   1140
               TabIndex        =   131
               Top             =   840
               Width           =   735
            End
            Begin VB.Label LLibellesPostes 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               BorderStyle     =   1  'Fixed Single
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
               ForeColor       =   &H80000008&
               Height          =   315
               Index           =   41
               Left            =   1140
               TabIndex        =   130
               Top             =   540
               Width           =   735
            End
            Begin VB.Label LNomsPostes 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H00C0E0FF&
               BorderStyle     =   1  'Fixed Single
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
               Height          =   315
               Index           =   42
               Left            =   240
               TabIndex        =   129
               Top             =   840
               Width           =   915
            End
            Begin VB.Label LNomsPostes 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H00C0E0FF&
               BorderStyle     =   1  'Fixed Single
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
               Height          =   315
               Index           =   41
               Left            =   240
               TabIndex        =   128
               Top             =   540
               Width           =   915
            End
            Begin VB.Image IEtatsChariots 
               Appearance      =   0  'Flat
               BorderStyle     =   1  'Fixed Single
               Height          =   315
               Index           =   41
               Left            =   1860
               Picture         =   "FSynoptique.frx":2E867
               Stretch         =   -1  'True
               Top             =   540
               Width           =   435
            End
            Begin VB.Image IEtatsChariots 
               Appearance      =   0  'Flat
               BorderStyle     =   1  'Fixed Single
               Height          =   315
               Index           =   42
               Left            =   1860
               Picture         =   "FSynoptique.frx":2EC2D
               Stretch         =   -1  'True
               Top             =   840
               Width           =   435
            End
            Begin VB.Image IEtatsPostes 
               Appearance      =   0  'Flat
               BorderStyle     =   1  'Fixed Single
               Height          =   315
               Index           =   42
               Left            =   2700
               Picture         =   "FSynoptique.frx":2EFF3
               Stretch         =   -1  'True
               Top             =   840
               Width           =   435
            End
            Begin VB.Image IEtatsPostes 
               Appearance      =   0  'Flat
               BorderStyle     =   1  'Fixed Single
               Height          =   315
               Index           =   41
               Left            =   2700
               Picture         =   "FSynoptique.frx":2F3B9
               Stretch         =   -1  'True
               Top             =   540
               Width           =   435
            End
            Begin VB.Shape SDecoration 
               BackColor       =   &H00E0E0E0&
               BackStyle       =   1  'Opaque
               Height          =   855
               Index           =   1
               Left            =   120
               Top             =   420
               Width           =   3135
            End
         End
         Begin VB.PictureBox PBRedresseursAnodisation 
            BackColor       =   &H00C0E0FF&
            Height          =   4155
            Left            =   6120
            ScaleHeight     =   4095
            ScaleWidth      =   8055
            TabIndex        =   117
            Top             =   180
            Width           =   8115
            Begin Anodisation.OCXRedresseur OCXRedresseurs 
               Height          =   3255
               Index           =   1
               Left            =   6120
               TabIndex        =   118
               Top             =   720
               Width           =   1815
               _ExtentX        =   3149
               _ExtentY        =   5715
               Modele          =   1
            End
            Begin Anodisation.OCXRedresseur OCXRedresseurs 
               Height          =   3240
               Index           =   2
               Left            =   4140
               TabIndex        =   119
               Top             =   720
               Width           =   1785
               _ExtentX        =   3149
               _ExtentY        =   5715
               Modele          =   1
            End
            Begin Anodisation.OCXRedresseur OCXRedresseurs 
               Height          =   3240
               Index           =   3
               Left            =   2160
               TabIndex        =   120
               Top             =   720
               Width           =   1785
               _ExtentX        =   3149
               _ExtentY        =   5715
               Modele          =   1
            End
            Begin Anodisation.OCXRedresseur OCXRedresseurs 
               Height          =   3240
               Index           =   4
               Left            =   180
               TabIndex        =   121
               Top             =   720
               Width           =   1785
               _ExtentX        =   3149
               _ExtentY        =   5715
               Modele          =   1
            End
            Begin VB.Label LLibelles 
               Alignment       =   2  'Center
               BackColor       =   &H00FF0000&
               BorderStyle     =   1  'Fixed Single
               Caption         =   "ANODISATION"
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
               Height          =   315
               Index           =   8
               Left            =   0
               TabIndex        =   141
               Top             =   0
               Width           =   8055
            End
            Begin VB.Label LLibelles 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H00FF0000&
               BorderStyle     =   1  'Fixed Single
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
               Height          =   255
               Index           =   11
               Left            =   180
               TabIndex        =   127
               Top             =   465
               Width           =   1785
            End
            Begin VB.Label LLibelles 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H00FF0000&
               BorderStyle     =   1  'Fixed Single
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
               Height          =   255
               Index           =   10
               Left            =   2160
               TabIndex        =   126
               Top             =   465
               Width           =   1785
            End
            Begin VB.Label LLibelles 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H00FF0000&
               BorderStyle     =   1  'Fixed Single
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
               Height          =   255
               Index           =   9
               Left            =   4140
               TabIndex        =   125
               Top             =   465
               Width           =   1785
            End
            Begin VB.Label LLibelles 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H00FF0000&
               BorderStyle     =   1  'Fixed Single
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
               Height          =   255
               Index           =   7
               Left            =   6120
               TabIndex        =   122
               Top             =   465
               Width           =   1785
            End
         End
         Begin VB.PictureBox PBChargement 
            BackColor       =   &H00C0E0FF&
            Height          =   2055
            Left            =   24480
            ScaleHeight     =   1995
            ScaleWidth      =   3375
            TabIndex        =   124
            Top             =   180
            Width           =   3435
            Begin VB.Label LLibelles 
               Alignment       =   2  'Center
               BackColor       =   &H00FF0000&
               BorderStyle     =   1  'Fixed Single
               Caption         =   "CHARGEMENT"
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
               Height          =   315
               Index           =   4
               Left            =   0
               TabIndex        =   140
               Top             =   0
               Width           =   3375
            End
            Begin VB.Label LLibellesPostes 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               BorderStyle     =   1  'Fixed Single
               Caption         =   "-"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   315
               Index           =   2
               Left            =   1140
               TabIndex        =   139
               Top             =   840
               Width           =   735
            End
            Begin VB.Label LLibellesPostes 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               BorderStyle     =   1  'Fixed Single
               Caption         =   "-"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   315
               Index           =   1
               Left            =   1140
               TabIndex        =   138
               Top             =   540
               Width           =   735
            End
            Begin VB.Image IEtatsChariots 
               Appearance      =   0  'Flat
               BorderStyle     =   1  'Fixed Single
               Height          =   315
               Index           =   2
               Left            =   1860
               Picture         =   "FSynoptique.frx":2F77F
               Stretch         =   -1  'True
               Top             =   840
               Width           =   435
            End
            Begin VB.Image IEtatsChariots 
               Appearance      =   0  'Flat
               BorderStyle     =   1  'Fixed Single
               Height          =   315
               Index           =   1
               Left            =   1860
               Picture         =   "FSynoptique.frx":2FB45
               Stretch         =   -1  'True
               Top             =   540
               Width           =   435
            End
            Begin VB.Image IEtatsPostes 
               Appearance      =   0  'Flat
               BorderStyle     =   1  'Fixed Single
               Height          =   315
               Index           =   2
               Left            =   2700
               Picture         =   "FSynoptique.frx":2FF0B
               Stretch         =   -1  'True
               Top             =   840
               Width           =   435
            End
            Begin VB.Image IEtatsPostes 
               Appearance      =   0  'Flat
               BorderStyle     =   1  'Fixed Single
               Height          =   315
               Index           =   1
               Left            =   2700
               Picture         =   "FSynoptique.frx":302D1
               Stretch         =   -1  'True
               Top             =   540
               Width           =   435
            End
            Begin VB.Label LNumCharges 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
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
               ForeColor       =   &H80000008&
               Height          =   315
               Index           =   2
               Left            =   2280
               TabIndex        =   137
               Top             =   840
               Width           =   435
            End
            Begin VB.Label LNumCharges 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
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
               ForeColor       =   &H80000008&
               Height          =   315
               Index           =   1
               Left            =   2280
               TabIndex        =   136
               Top             =   540
               Width           =   435
            End
            Begin VB.Label LNomsPostes 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H00C0E0FF&
               BorderStyle     =   1  'Fixed Single
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
               Height          =   315
               Index           =   2
               Left            =   240
               TabIndex        =   135
               Top             =   840
               Width           =   915
            End
            Begin VB.Label LNomsPostes 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H00C0E0FF&
               BorderStyle     =   1  'Fixed Single
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
               Height          =   315
               Index           =   1
               Left            =   240
               TabIndex        =   134
               Top             =   540
               Width           =   915
            End
            Begin VB.Shape SDecoration 
               BackColor       =   &H00E0E0E0&
               BackStyle       =   1  'Opaque
               Height          =   1455
               Index           =   2
               Left            =   120
               Top             =   420
               Width           =   3135
            End
         End
         Begin VB.PictureBox PBParametres 
            BackColor       =   &H00C0E0FF&
            Height          =   4155
            Left            =   14400
            ScaleHeight     =   4095
            ScaleWidth      =   9855
            TabIndex        =   145
            Top             =   180
            Width           =   9915
            Begin VB.Frame FChoixBD 
               BackColor       =   &H00C0E0FF&
               Caption         =   " Base de données "
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   1155
               Left            =   3900
               TabIndex        =   302
               Top             =   420
               Width           =   2235
               Begin VB.OptionButton OBChoixBaseDeDonnees 
                  BackColor       =   &H00C0E0FF&
                  Caption         =   "SAGE"
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   240
                  Index           =   1
                  Left            =   240
                  TabIndex        =   304
                  Top             =   720
                  Width           =   1275
               End
               Begin VB.OptionButton OBChoixBaseDeDonnees 
                  BackColor       =   &H00C0E0FF&
                  Caption         =   "CLIPPER"
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   240
                  Index           =   0
                  Left            =   240
                  TabIndex        =   303
                  Top             =   360
                  Value           =   -1  'True
                  Width           =   1275
               End
            End
            Begin VB.CommandButton CBEntretienGraphesProduction 
               BackColor       =   &H00C0FFFF&
               Caption         =   "Entretien des graphes de production"
               DownPicture     =   "FSynoptique.frx":30697
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   1035
               Left            =   6300
               MaskColor       =   &H00FF00FF&
               Picture         =   "FSynoptique.frx":30DE1
               Style           =   1  'Graphical
               TabIndex        =   300
               Top             =   480
               UseMaskColor    =   -1  'True
               Visible         =   0   'False
               Width           =   3435
            End
            Begin VB.CommandButton CBEssais2 
               Caption         =   "Essais 2"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   495
               Left            =   3900
               TabIndex        =   299
               Top             =   2340
               Visible         =   0   'False
               Width           =   1635
            End
            Begin VB.Frame FModeEntreeCharges 
               BackColor       =   &H00C0E0FF&
               Caption         =   " Modes d'entrées des charges "
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   1755
               Left            =   180
               TabIndex        =   251
               Top             =   2040
               Width           =   3555
               Begin VB.CommandButton CBEntreeAutomatiqueCharges 
                  BackColor       =   &H00FFFFC0&
                  Caption         =   "CHOIX du MODE des ENTREES"
                  BeginProperty Font 
                     Name            =   "Arial"
                     Size            =   9
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   375
                  Left            =   360
                  Style           =   1  'Graphical
                  TabIndex        =   252
                  Top             =   420
                  Width           =   2835
               End
               Begin VB.Label LEntreeAutomatiqueCharges 
                  Alignment       =   2  'Center
                  Appearance      =   0  'Flat
                  BackColor       =   &H000000FF&
                  BorderStyle     =   1  'Fixed Single
                  Caption         =   "MODE MANUEL - ENTREE IMMEDIATE DES CHARGES"
                  BeginProperty Font 
                     Name            =   "Arial"
                     Size            =   9
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00FFFFFF&
                  Height          =   495
                  Left            =   360
                  TabIndex        =   301
                  Top             =   960
                  Width           =   2835
               End
            End
            Begin VB.CommandButton CBFinDeJournee 
               BackColor       =   &H00C0FFFF&
               Caption         =   "Fin de journée"
               DownPicture     =   "FSynoptique.frx":3152B
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   915
               Left            =   6300
               MaskColor       =   &H00FF00FF&
               Picture         =   "FSynoptique.frx":3318D
               Style           =   1  'Graphical
               TabIndex        =   250
               Top             =   1680
               UseMaskColor    =   -1  'True
               Visible         =   0   'False
               Width           =   3435
            End
            Begin VB.CommandButton CBEssais 
               Caption         =   "Essais"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   495
               Left            =   3900
               TabIndex        =   249
               Top             =   1740
               Visible         =   0   'False
               Width           =   1635
            End
            Begin VB.Frame FModeAffichageSynoptique 
               BackColor       =   &H00C0E0FF&
               Caption         =   " Mode d'affichage "
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   1515
               Left            =   180
               TabIndex        =   166
               Top             =   420
               Width           =   3555
               Begin VB.OptionButton OBModeAffichageSynoptique 
                  BackColor       =   &H00C0E0FF&
                  Caption         =   " Par les CHARGES COLOREES"
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   9.75
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   2
                  Left            =   240
                  TabIndex        =   305
                  Top             =   1080
                  Width           =   3195
               End
               Begin VB.OptionButton OBModeAffichageSynoptique 
                  BackColor       =   &H00C0E0FF&
                  Caption         =   " Par le n° de BARRES"
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   9.75
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   1
                  Left            =   240
                  TabIndex        =   168
                  Top             =   720
                  Width           =   2595
               End
               Begin VB.OptionButton OBModeAffichageSynoptique 
                  BackColor       =   &H00C0E0FF&
                  Caption         =   " Par le n° de CHARGES"
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   9.75
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   0
                  Left            =   240
                  TabIndex        =   167
                  Top             =   360
                  Value           =   -1  'True
                  Width           =   2595
               End
            End
            Begin VB.Label LLibelles 
               Alignment       =   2  'Center
               BackColor       =   &H00FF0000&
               BorderStyle     =   1  'Fixed Single
               Caption         =   "PARAMETRES / FONCTIONS DIVERSES"
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
               Height          =   315
               Index           =   13
               Left            =   0
               TabIndex        =   146
               Top             =   0
               Width           =   9855
            End
         End
      End
      Begin VB.PictureBox PBBarreAgrandissement 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   16200
         Index           =   0
         Left            =   0
         Picture         =   "FSynoptique.frx":34DEF
         ScaleHeight     =   16200
         ScaleWidth      =   435
         TabIndex        =   3
         Top             =   0
         Width           =   435
         Begin VB.CommandButton CBAgrandirRestaurerZonesFenetre 
            BackColor       =   &H00FFFFFF&
            DownPicture     =   "FSynoptique.frx":4C171
            Height          =   255
            Index           =   0
            Left            =   60
            MaskColor       =   &H00FF00FF&
            Picture         =   "FSynoptique.frx":4C31B
            Style           =   1  'Graphical
            TabIndex        =   321
            ToolTipText     =   " Agrandir / Réduire "
            Top             =   120
            UseMaskColor    =   -1  'True
            Width           =   255
         End
         Begin VB.Label LAppelZones 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "F2"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000FFFF&
            Height          =   240
            Index           =   3
            Left            =   60
            TabIndex        =   73
            Top             =   660
            Width           =   270
         End
         Begin VB.Label LAppelZones 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Maj"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000FFFF&
            Height          =   225
            Index           =   0
            Left            =   45
            TabIndex        =   6
            Top             =   420
            Width           =   300
         End
      End
      Begin VB.PictureBox PBImageLigne 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   435
         Left            =   -600
         ScaleHeight     =   29
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   53
         TabIndex        =   2
         Top             =   0
         Width           =   795
      End
   End
   Begin MSComctlLib.ImageList ILOutilsDivers 
      Left            =   9720
      Top             =   13080
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   12
      ImageHeight     =   10
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   11
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FSynoptique.frx":4C4C5
            Key             =   "agrandir en haut"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FSynoptique.frx":4C677
            Key             =   "restaurer taille en haut"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FSynoptique.frx":4C829
            Key             =   "rectangle blanc"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FSynoptique.frx":4CC05
            Key             =   "rectangle vert"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FSynoptique.frx":4CFDD
            Key             =   "rectangle orange"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FSynoptique.frx":4D3B5
            Key             =   "rectangle rouge"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FSynoptique.frx":4D78D
            Key             =   "chariot present"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FSynoptique.frx":4DB6D
            Key             =   "chariot present verrouille"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FSynoptique.frx":4DF45
            Key             =   "chariot present verrouille charge"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FSynoptique.frx":4E325
            Key             =   "croix de condamnation 1"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FSynoptique.frx":4E6FD
            Key             =   "croix de condamnation 2"
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox PBDialogues 
      BackColor       =   &H00E0E0E0&
      Height          =   2595
      Left            =   60
      ScaleHeight     =   169
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   632
      TabIndex        =   0
      Top             =   13080
      Width           =   9540
      Begin VB.PictureBox PBBarreAgrandissement 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   16200
         Index           =   1
         Left            =   0
         Picture         =   "FSynoptique.frx":4EAD5
         ScaleHeight     =   16200
         ScaleWidth      =   435
         TabIndex        =   4
         Top             =   0
         Width           =   435
         Begin VB.CommandButton CBAgrandirRestaurerZonesFenetre 
            BackColor       =   &H00FFFFFF&
            DownPicture     =   "FSynoptique.frx":65E57
            Height          =   255
            Index           =   1
            Left            =   60
            MaskColor       =   &H00FF00FF&
            Picture         =   "FSynoptique.frx":66001
            Style           =   1  'Graphical
            TabIndex        =   5
            ToolTipText     =   " Agrandir / Réduire "
            Top             =   60
            UseMaskColor    =   -1  'True
            Width           =   255
         End
         Begin VB.Label LAppelZones 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Maj"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000FFFF&
            Height          =   225
            Index           =   2
            Left            =   60
            TabIndex        =   74
            Top             =   420
            Width           =   300
         End
         Begin VB.Label LAppelZones 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "F3"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000FFFF&
            Height          =   240
            Index           =   1
            Left            =   60
            TabIndex        =   7
            Top             =   660
            Width           =   270
         End
      End
      Begin RichTextLib.RichTextBox RTBDialogues 
         Height          =   855
         Left            =   420
         TabIndex        =   8
         Top             =   420
         Width           =   2475
         _ExtentX        =   4366
         _ExtentY        =   1508
         _Version        =   393217
         BorderStyle     =   0
         Enabled         =   -1  'True
         ScrollBars      =   3
         AutoVerbMenu    =   -1  'True
         TextRTF         =   $"FSynoptique.frx":661AB
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin ComCtl3.CoolBar COBConteneurOutilsDialogues 
         Height          =   435
         Left            =   420
         TabIndex        =   9
         Top             =   0
         Width           =   11055
         _ExtentX        =   19500
         _ExtentY        =   767
         FixedOrder      =   -1  'True
         _CBWidth        =   11055
         _CBHeight       =   435
         _Version        =   "6.7.9816"
         Child1          =   "TOBEffacementDialogues"
         MinWidth1       =   23
         MinHeight1      =   22
         Width1          =   23
         FixedBackground1=   0   'False
         NewRow1         =   0   'False
         Child2          =   "TOBOutilsDialogues"
         MinWidth2       =   455
         MinHeight2      =   25
         Width2          =   455
         NewRow2         =   0   'False
         Child3          =   "TOBOutilsIntranet"
         MinHeight3      =   22
         Width3          =   463
         NewRow3         =   0   'False
         Visible3        =   0   'False
         Begin MSComctlLib.Toolbar TOBOutilsDialogues 
            Height          =   375
            Left            =   600
            TabIndex        =   11
            Top             =   30
            Width           =   6825
            _ExtentX        =   12039
            _ExtentY        =   661
            ButtonWidth     =   4657
            ButtonHeight    =   661
            AllowCustomize  =   0   'False
            Wrappable       =   0   'False
            Style           =   1
            ImageList       =   "ILOutilsDialogues2"
            HotImageList    =   "ILOutilsDialogues2"
            _Version        =   393216
            BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
               NumButtons      =   8
               BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "renseignements"
                  Object.ToolTipText     =   " Passage en mode renseignements "
                  ImageIndex      =   2
               EndProperty
               BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Style           =   3
               EndProperty
               BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "questions reponses"
                  Object.ToolTipText     =   " Passage en mode questions / réponses "
                  ImageIndex      =   3
               EndProperty
               BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Style           =   3
               EndProperty
               BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "previsionnel"
                  Object.ToolTipText     =   " Passage en mode d'informations sur le prévisionnel "
                  ImageIndex      =   5
               EndProperty
               BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Style           =   3
               EndProperty
               BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "entrees charges"
                  Object.ToolTipText     =   " Passage en mode d'informations sur les entrées "
                  ImageIndex      =   7
               EndProperty
               BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Style           =   3
               EndProperty
            EndProperty
         End
         Begin MSComctlLib.Toolbar TOBEffacementDialogues 
            Height          =   330
            Left            =   30
            TabIndex        =   10
            Top             =   45
            Width           =   345
            _ExtentX        =   609
            _ExtentY        =   582
            ButtonWidth     =   609
            ButtonHeight    =   582
            AllowCustomize  =   0   'False
            Wrappable       =   0   'False
            Style           =   1
            ImageList       =   "ILOutilsDialogues1"
            HotImageList    =   "ILOutilsDialogues1"
            _Version        =   393216
            BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
               NumButtons      =   1
               BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "effacer"
                  Object.ToolTipText     =   " Effacement des dialogues "
                  ImageIndex      =   1
               EndProperty
            EndProperty
         End
      End
   End
   Begin MSComctlLib.ImageList ILOutilsDialogues1 
      Left            =   10380
      Top             =   13080
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   16711935
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   1
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FSynoptique.frx":6622F
            Key             =   "effacer ecran"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ILOutilsIntranet 
      Left            =   11040
      Top             =   13080
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   16711935
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   5
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FSynoptique.frx":66583
            Key             =   "precedente"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FSynoptique.frx":668D7
            Key             =   "suivante"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FSynoptique.frx":66C2B
            Key             =   "arreter"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FSynoptique.frx":66F7F
            Key             =   "actualiser"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FSynoptique.frx":672D3
            Key             =   "demarrage"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ILProgrammateurCycliqueCuves 
      Left            =   11040
      Top             =   13725
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   22
      ImageHeight     =   24
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   3
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FSynoptique.frx":67627
            Key             =   "chronometre fond blanc"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FSynoptique.frx":67CD9
            Key             =   "chronometre fond orange"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FSynoptique.frx":6838B
            Key             =   "chronometre fond cyan"
         EndProperty
      EndProperty
   End
   Begin VB.Image IEtatsPostes 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Height          =   315
      Index           =   0
      Left            =   0
      Picture         =   "FSynoptique.frx":68A3D
      Stretch         =   -1  'True
      Top             =   0
      Width           =   435
   End
   Begin VB.Label LNomsPostes 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0E0FF&
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
      ForeColor       =   &H80000008&
      Height          =   315
      Index           =   0
      Left            =   0
      TabIndex        =   307
      Top             =   0
      Width           =   615
   End
End
Attribute VB_Name = "FSynoptique"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Rôle                    : Fenêtre affichant le synoptique de la ligne
' Nom                    : FSynoptique.frm
' Date de création : 01/10/2010
' Détails                :
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

'--- déclarations obligatoires ---
Option Explicit

'--- options générales ---
Option Base 1
DefVar A-Z


'--- constantes privées ---
Private Const LONGUEUR_BASE_ETATS_LIGNE As Integer = 191

Private Const NBR_LIGNES_DETAILS_EN_COURS As Integer = CHARGES.C_NUM_MAXI
Private Const NBR_COLONNES_DETAILS_EN_COURS  As Integer = 7

Private Const TITRE_FENETRE As String = "SYNOPTIQUE"
Private Const TITRE_MESSAGES As String = TITRE_FENETRE

Public mlID As Long

'--- énumérations privées ---

'--- les zones de la fenêtre ---
Private Enum ZONES_FENETRE
    Z_SYNOPTIQUE = 0
    Z_DIALOGUES = 1
End Enum

'--- noms des animations ---
Private Enum NOMS_ANIMATIONS
    N_PONT = 0                                                  'pont de la ligne
    N_PONT_FICTIF = 1                                      'pont fictif
    N_PALONNIER = 2                                        'palonnier
    N_ACCROCHE = 3                                         'accroche d'une charge
    N_CHARGE_PONT = 4                                  'charge sur le pont
    N_CHARGE_POSTE = 5                                'charge dans un poste
    N_COUVERCLES = 6                                     'couvercles
    N_CHARIOTS = 8                                          'chariots
    N_LIBELLES = 9                                           'libellés
End Enum

'--- positions d'origines pour les éléments attachés aux autres (exemples bacs anti-égouttures) ---
Private Enum POSITIONS_ORIGINES

    PO_Y_PONT = 22                                                          'la position Y du pont ne bouge pas
    PO_Y_PONT_FICTIF = 22                                              'la position Y du pont ne bouge pas
    
    AJOUT_X_PALONNIER = 18                                          'la position X du palonnier ne bouge pas
    AJOUT_Y_PALONNIER_BAS_PONT = 53
    
    AJOUT_X_ACCROCHE = 24                                           'la position X de l'accroche ne bouge pas

    PO_Y_CHARIOT = 99                                                    'la position Y d'un chariot ne bouge pas
    
    PO_Y_CHARGE_POSTE = 99                                        'la position Y d'une charge au poste ne bouge pas
    
    PO_Y_CHARGE_PONT = 24
    PO_Y_CHARGE_HAUT_PONT = 46                               'la position Y d'une charge au niveau haut du pont
    PO_Y_CHARGE_BAS_PONT = 99                                 'la position Y d'une charge au niveau bas du pont

    PO_X_PONT_1_HORS_LIGNE = 1200                          'position X du pont 1 hors ligne
    PO_X_PONT_2_HORS_LIGNE = 1                                'position X du pont 2 hors ligne

End Enum

'--- ajouts pour les niveaux ---
Private Enum AJOUTS_NIVEAUX
    A_Y_NIVEAU_HAUT = 0
    A_Y_NIVEAU_EGOUTTAGE = 25
    A_Y_NIVEAU_INTERMEDIAIRE = 40
    A_Y_NIVEAU_BAS = 58
End Enum

'--- pour signaler les differents états des ponts ---
Private Enum TYPES_AFFICHAGES_PONTS
    T_NORMAL = 0
    T_DEFAUT_BLANC = 1
    T_DEFAUT_ROUGE = 2
    T_CONDAMNATION = 3
    T_FICTIF_P1 = 4                             'fictif pont 1
    T_FICTIF_P2 = 5                             'fictif pont 2
End Enum

Private Enum TYPES_AFFICHAGES_PALONNIERS
    T_NORMAL = 0
    T_DEFAUT_BLANC = 1
    T_DEFAUT_ROUGE = 2
    T_CONDAMNATION = 3
End Enum

Private Enum TYPES_AFFICHAGES_ACCROCHES
    T_NORMAL = 0
    T_DEFAUT_BLANC = 1
    T_DEFAUT_ROUGE = 2
    T_CONDAMNATION = 3
End Enum

'--- pour signaler dans les libellés des postes les différents états ---
Private Enum TYPES_AFFICHAGES_LIBELLES
    E_NORMAL = 0
    E_DEFAUT = 1
    E_CONDAMNATION = 2
End Enum

'--- énumérations publiques ---

'--- les formes possibles de la fenêtre ---
Public Enum FORMES_FENETRE
    F_STANDARD = 0                      'synoptique + dialogues + états de la ligne
    F_SYNOPTIQUE = 1                  'synoptique seul
    F_DIALOGUES = 2                     'dialogues seul
End Enum

'--- les onglets des zones de dialogues ---
Public Enum ONGLETS_DIALOGUES
    O_CHARGEMENT = 0
    O_PREVISIONNEL = 1
    O_DIALOGUES = 2
End Enum

'--- modes des dialogues ---
Public Enum MODES_DIALOGUES
    M_RENSEIGNEMENTS = 0
    M_QUESTIONS_REPONSES = 1
    M_PREVISIONNEL = 2
    M_ENTREE_CHARGES = 3
End Enum

'--- colonnes des détails des en cours ---
Private Enum COLONNES_DETAILS_EN_COURS
    C_NUM_LIGNES = 0
    C_NUM_COMMANDE_INTERNE = 1
    C_POSTE = 2
    C_NUM_BARRE = 3
    C_CODE_CLIENT = 4
    C_NBR_PIECES = 5
    C_DESIGNATION = 6
    C_MATIERE = 7
End Enum

'--- pour la vitesse des animations à l'écran (par la mesure de différence de temps) ---
Private Type VitesseAnimations
    PremierPassage As Boolean
    DateDernierPassage As Date
End Type

'--- variables privées ---
Private PremiereActivation As Boolean                                                'première activation de la fenêtre
Private ClignotantPourSynoptique As Boolean                                     'constitue un clignotant pour le synoptique

Private FormeFenetre As Integer                                                           'indique la forme actuelle de la fenêtre

'--- tableau privées ---

'--- variables et tableaux privées DIRECTX 7.0 ---
Private ObjDX As New DirectX7                                                            'objet DirectX
Private ObjDD As DirectDraw7                                                              'objet DirectDraw
        
Private ObjDDSEcran As DirectDrawSurface7                                       'objet de la surface de l'écran
Private DDSDEcran As DDSURFACEDESC2                                          'description de la surface de l'écran

Private ObjDDClip As DirectDrawClipper                                               'objet du clipper

Private ObjDDSImageLigne As DirectDrawSurface7                             'objet de la surface de l'image de la ligne
Private DDSDImageLigne As DDSURFACEDESC2                                 'description de la surface de l'image de la ligne
Private RImageLigne As RECT                                                               'coordonnées du rectangle de l'image de la ligne

'--- tableaux privées ---
Private TObjDDSEnsemblePonts As DirectDrawSurface7                     'objets des surfaces des ponts
Private TDDSDEnsemblePonts As DDSURFACEDESC2                        'description des surfaces des ponts

Private TXPontsFictifs(PONTS.P_1 To PONTS.P_2) As Long                 'tableau des X des ponts fictifs
Private TDerniersXPontsFictifs(PONTS.P_1 To PONTS.P_2) As Long   'tableau des derniers X des ponts fictifs
Private TYPontsFictifs(PONTS.P_1 To PONTS.P_2) As Long                 'tableau des Y des ponts fictifs
Private TDerniersYPontsFictifs(PONTS.P_1 To PONTS.P_2) As Long   'tableau desderniers Y des ponts  fictifs

Private TObjDDSEnsemblePalonniers As DirectDrawSurface7              'objets des surfaces des palonniers
Private TDDSDEnsemblePalonniers As DDSURFACEDESC2                 'description des surfaces des palonniers
Private TXPalonniers(PONTS.P_1 To PONTS.P_2) As Single                 'X des palonniers
Private TDerniersXPalonniers(PONTS.P_1 To PONTS.P_2) As Single   'derniers X des palonniers
Private TYPalonniers(PONTS.P_1 To PONTS.P_2) As Single                 'Y des palonniers
Private TDerniersYPalonniers(PONTS.P_1 To PONTS.P_2) As Single   'derniers Y des palonniers

Private TObjDDSEnsembleAccroches As DirectDrawSurface7               'objets des surfaces des accroches d'une charge
Private TDDSDEnsembleAccroches As DDSURFACEDESC2                  'description des surfaces des accroches d'une charge
Private TXAccroches(PONTS.P_1 To PONTS.P_2) As Long                     'tableau des X des accroches d'une charge
Private TDerniersXAccroches(PONTS.P_1 To PONTS.P_2) As Long       'tableau des derniers X des accroches d'une charge
Private TYAccroches(PONTS.P_1 To PONTS.P_2) As Long                     'tableau des Y des accroches d'une charge
Private TDerniersYAccroches(PONTS.P_1 To PONTS.P_2)  As Long      'tableau des derniers Y des accroches d'une charge

Private TObjDDSEnsembleCharges As DirectDrawSurface7                   'objet surface de l'ensemble des charges
Private TDDSDEnsembleCharges As DDSURFACEDESC2                      'description de la surface de l'ensemble des charges
Private TXChargesPonts(PONTS.P_1 To PONTS.P_2) As Long               'X d'une charge d'une charge sur un pont
Private TDerniersXChargesPonts(PONTS.P_1 To PONTS.P_2)  As Long 'derniers X d'une charge sur un pont
Private TYChargesPonts(PONTS.P_1 To PONTS.P_2)  As Long              'Y d'une charge d'une charge sur un pont
Private TDerniersYChargesPonts(PONTS.P_1 To PONTS.P_2)  As Long 'derniers Y d'une charge sur un pont

Private TXChargePoste As Long                                                              'X d'une charge dans un poste
Private TDernierXChargePoste As Long                                                  'derniers X d'une charge dans un poste
Private TYChargePoste As Long                                                              'Y d'une charge dans un poste
Private TDernierYChargePoste As Long                                                  'derniers Y d'une charge dans un poste

Private TObjDDSEnsembleCouvercles As DirectDrawSurface7             'objets des surfaces de l'ensemble des couvercles
Private TDDSDEnsembleCouvercles As DDSURFACEDESC2                'description des surfaces de l'ensemble des couvercles
Private TXCouvercles As Long                                                                'X des couvercles
Private TDernierXCouvercles As Long                                                    'derniers X des couvercles
Private TYCouvercles As Long                                                                'Y des couvercles
Private TDernierYCouvercles As Long                                                    'derniers Y des couvercles

Private TObjDDSChariot(ETATS_CHARIOTS.E_PRESENT To ETATS_CHARIOTS.E_PRESENT_VERROUILLE) As DirectDrawSurface7    'objets des surfaces d'un chariots
Private TDDSDChariot As DDSURFACEDESC2                                      'description des surfaces d'un chariots
Private TXChariot As Long                                                                      'X d'un chariots
Private TDernierXChariot As Long                                                          'derniers X d'un chariots
Private TYChariot As Long                                                                      'Y d'un chariots
Private TDernierYChariot As Long                                                          'derniers Y d'un chariots

Private TObjDDSEnsembleLibelles As DirectDrawSurface7                  'objets des surfaces de l'ensemble des libellés
Private TDDSDEnsembleLibelles As DDSURFACEDESC2                     'description des surfaces de l'ensemble des libellés
Private TXLibelle As Long                                                                       'X des libellés
Private TDernierXLibelle As Long                                                           'derniers X des libellés
Private TYLibelle As Long                                                                       'Y des libellés
Private TDernierYLibelle As Long                                                           'derniers Y des libellés

'--- variables publiques ---
Public ArretTachesRapides As Boolean                                                 'TRUE = arrêt du noyau des taches rapides
Public ModeDialoguesEnCours As Integer                                             'mode des dialogues en cours
Public NumFenetre As Long                                                                    'numéro de la fenêtre lorsqu'elle devient active

Private NbPosteClicked As Integer



Private Sub CBAgrandirRestaurerZonesFenetre_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next
    
    If Button = vbLeftButton Then
        
        '--- rafraichissement ---
        CBAgrandirRestaurerZonesFenetre(Index).Refresh
        
        '--- changement de la forme de la fenêtre ---
        If FormeFenetre <> FORMES_FENETRE.F_STANDARD Then
            ChangeFormeFenetre FORMES_FENETRE.F_STANDARD
        Else
            Select Case Index
                Case ZONES_FENETRE.Z_SYNOPTIQUE: ChangeFormeFenetre F_SYNOPTIQUE
                Case ZONES_FENETRE.Z_DIALOGUES: ChangeFormeFenetre F_DIALOGUES
                Case Else
            End Select
        End If
    
    End If

End Sub

Private Sub CBEntretienGraphesProduction_Click()
    On Error Resume Next
    AppelFenetre F_NETTOYAGE_GRAPHES_PRODUCTION
End Sub

Private Sub CBEssais_Click()
    
    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next

    '--- déclaration ---
    Dim a As Integer
    Dim NumCharge As Integer
    Static Cpt As Integer
    
    '--- affectation ---
    NumCharge = 7
    
    
    For a = PREMIER_BAIN To POSTES.P_C12
        TEtatsPostes(a).NumCharge = a
        If a <= POSTES.P_C07 Then
            TEtatsCharges(TEtatsPostes(a).NumCharge).TGammesAnodisation.PassageAnodisation = True
        End If
    Next a
    
    Select Case Cpt
        Case 0
            TEtatsCharges(POSTES.P_C00).TGammesAnodisation.PassageSpectro = True
        Case 1
            TEtatsCharges(POSTES.P_DEC).TGammesAnodisation.PassageOr = True
            TEtatsCharges(POSTES.P_DEC).TGammesAnodisation.PassageSpectro = True
        Case 2
            TEtatsCharges(POSTES.P_SAT).TGammesAnodisation.PassageNoir = True
            TEtatsCharges(POSTES.P_SAT).TGammesAnodisation.PassageSpectro = True
        Case 3
            TEtatsCharges(POSTES.P_C03).TGammesAnodisation.PassageOr = True
        Case 4
            TEtatsCharges(POSTES.P_C04).TGammesAnodisation.PassageNoir = True
        Case Else
    End Select
    Inc Cpt
    
    Exit Sub
    
    '--- affectation ---
    NumCharge = 7

    With TEtatsCharges(NumCharge)
        
        .DateEntreeEnLigne = Now
        .DateArriveeAuDechargement = DateAdd("h", 2, Now)
        
        .NumBarre = 10
        
        With .TDetailsCharges(1)
            .NumCommandeInterne = 1
            .CodeClient = "CODE CLIENT 1"
            .NbrPieces = 3
            .Designation = "PLAQUES"
            .Matiere = "ALU"
        End With
        
        With .TDetailsCharges(2)
            .NumCommandeInterne = 2
            .CodeClient = "CODE CLIENT 2"
            .NbrPieces = 6
            .Designation = "PLAQUES 2"
            .Matiere = "ALU 2"
        End With

    End With

    '--- enregistrement des bains pour CLIPPER ---
    EnregistrementBainsPourCLIPPER NumCharge

End Sub

Private Sub CBEssais2_Click()
    
'    '--- aiguillage en cas d'erreurs ---
'    On Error Resume Next
'
'    '--- déclaration ---
'    Dim a  As Integer                                            'pour les boucles FOR...NEXT
'
'    '--- charges sur le pont ---
'    'For a = PONTS.P_1 To PONTS.P_2
'    '    With TEtatsPonts(a)
'    '        .NumCharge = a
'    '    End With
'    'Next a
'
'    '--- appel de la fenêtre ---
'    'AppelFenetre F_ESSAIS
'
'    'TEtatsRedresseurs(1).EtatRedresseur = ER_EXCLUSION
'
'    With TEtatsRedresseurs(REDRESSEURS.R_C13)
'
'        .NumCharge = 0
'
'    End With
'
'    With TEtatsRedresseurs(REDRESSEURS.R_C14)
'
'        .NumCharge = 0
'
'    End With
'
'    With TEtatsRedresseurs(REDRESSEURS.R_C15)
'
'        .NumCharge = 0
'
'    End With
'
'    'AppelFenetre F_VISUALISATION_GRAPHES_PRODUCTION
'
'
'
'    '--- aiguillage en cas d'erreurs ---
'    On Error Resume Next
'
'    '--- déclaration ---
'    Dim a  As Integer                                            'pour les boucles FOR...NEXT
'
'    '--- charges sur le pont ---
'    For a = PONTS.P_1 To PONTS.P_2
'        With TEtatsPonts(a)
'            .NumCharge = a
'        End With
'    Next a
'
'    '--- charges dans les postes ---
'    For a = POSTES.P_C00 To POSTES.P_C35
'        With TEtatsPostes(a)
'            .NumCharge = a
'        End With
'    Next a
'
'
'
'
'    '--- appel de la fenêtre ---
'    'AppelFenetre F_ESSAIS
'
'    'TEtatsRedresseurs(1).EtatRedresseur = ER_EXCLUSION
'
'    'With TEtatsRedresseurs(REDRESSEURS.R_C13)
'
'     '   .NumCharge = 1
'
'   ' End With
'
'    'With TEtatsRedresseurs(REDRESSEURS.R_C14)
'
'    '    .NumCharge = 8
'
'    'End With
'
'    'With TEtatsRedresseurs(REDRESSEURS.R_C15)
'
'     '   .NumCharge = 15
'
'    'End With
'
'    'AppelFenetre F_VISUALISATION_GRAPHES_PRODUCTION


End Sub

Private Sub CBFinDeJournee_Click()
    
    '--- aiguillage en cas d'erreurs ---
    'On Error Resume Next

    '--- fin de journée ---
    'AppelFenetre FENETRES.F_FIN_DE_JOURNEE

End Sub

Private Sub CBMonteeDescenteAccrochesPCP1_Click(Index As Integer)

    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next
    
    '--- constantes privées ---
    Const NOM_GROUPE = "MANUEL_PONTS"    'nom du groupe

    '--- déclaration ---
    Dim a  As Integer                                            'pour les boucles FOR...NEXT
    Dim ValeurRetourneeAPI As Long                  'valeur retournée par une fonction concernant le dialogue avec l'automate
    Dim NomVariable As String                            'nom de la variable OPC
        
    '--- transfert des valeurs ---
    If PROGRAMME_AVEC_AUTOMATE = True Then
    
        '--- transfert des valeurs ---
        If Index = 0 Then
            ValeurRetourneeAPI = APIEcritureVariableNommee(NOM_GROUPE, "M_ManuPCMontAccrocP1", "1")
        Else
            ValeurRetourneeAPI = APIEcritureVariableNommee(NOM_GROUPE, "M_ManuPCDescAccrocP1", "1")
        End If
    
    End If

End Sub

Private Sub CBMonteeDescenteAccrochesPCP2_Click(Index As Integer)
    
    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next
    
    '--- constantes privées ---
    Const NOM_GROUPE = "MANUEL_PONTS"    'nom du groupe

    '--- déclaration ---
    Dim a  As Integer                                            'pour les boucles FOR...NEXT
    Dim ValeurRetourneeAPI As Long                  'valeur retournée par une fonction concernant le dialogue avec l'automate
    Dim NomVariable As String                            'nom de la variable OPC
        
    '--- transfert des valeurs ---
    If PROGRAMME_AVEC_AUTOMATE = True Then
    
        '--- transfert des valeurs ---
        If Index = 0 Then
            ValeurRetourneeAPI = APIEcritureVariableNommee(NOM_GROUPE, "M_ManuPCMontAccrocP2", "1")
        Else
            ValeurRetourneeAPI = APIEcritureVariableNommee(NOM_GROUPE, "M_ManuPCDescAccrocP2", "1")
        End If
    
    End If

End Sub

Private Sub CBSauvegardeEtatsPostes_Click()
    
    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next

    '--- sauvegarde de l'états des postes ---
    SauveEtatsPostes

End Sub

Private Sub Form_Activate()

    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next
    NbPosteClicked = 1
    '--- renseigne la fenêtre principale ---
    RenseigneFPrincipale
   
    '--- placement du focus ---
    If PremiereActivation = False Then
        
        '--- affectation ---
        PremiereActivation = True
        
        '--- construction du cadre 3D ---
        ConstructionCadre3D
        
        '--- visualise les libellés de tous les états de la ligne ---
        VisualisationLibellesEtatsLigne
        
    End If

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error Resume Next
    GestionTouches KeyCode, Shift
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    
    '--- aiguillage en cas d'erreur ---
    On Error Resume Next

    '--- affectation ---
    ArretTachesRapides = True
    With TimerSynoptique
        .Interval = 0
        .Enabled = False
    End With
    
End Sub

Public Sub CopyComplete( _
           ByRef vRet As Variant)
    'Call Log("FIN insertionClipperPointage")
   
     
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    
    '--- aiguillage en cas d'erreur ---
    On Error Resume Next
    
    '--- affectation ---
    ArretTachesRapides = True
    With TimerSynoptique
        .Interval = 0
        .Enabled = False
    End With
    PremiereActivation = False

    '--- curseur souris par défaut ---
    Screen.MousePointer = vbDefault

End Sub

'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Rôle      : Effectue le paramètrage de la fenêtre
' Entrées :
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Sub ParametrageFenetre()

    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next

End Sub


'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Rôle      : Initialise la fenêtre (chargement ou en vue de la rendre visible)
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Sub InitialisationFenetre()
    
    '--- aiguillage en cas d'erreurs ---
    On Error GoTo GestionErreurs
    
    '--- déclaration ---
    
    '--- affectation ---

    '--- divers sur la fenêtre ---
    With Me
        .Caption = TITRE_FENETRE
        .Width = OccFPrincipale.ScaleWidth
        .Height = OccFPrincipale.ScaleHeight
    End With
    
    '--- images des fonds ---
    PBGeneral.Picture = ImgFondGris2
    PBEtatsLigne.Picture = ImgFondVert2
    PBEtatsPrincipaux.Picture = ImgFondOrange1
    
    '--- changement de la forme de la fenêtre ---
    ChangeFormeFenetre F_STANDARD
            
    '--- préparation de l'animation de la ligne ---
    InitialisationDirectX                          'initialisation de DirectX
    InitialisationSurfaces                        'Initialisation des surfaces
    PremieresPositionsAnimations        'premières positions des animations
    
    '--- gestion des en cours ---
    GestionEnCours GG_INITIALISATION
    GestionEnCours GG_AFFICHAGE
    
    '--- visualisation une première fois des états de la ligne ---
    VisualisationEtatsLigne
    
    '--- lancement du timer pour l'affichage synoptique ---
    TimerEtatsLigne.Enabled = True
    
    '--- lancement du timer pour l'affichage synoptique ---
    TimerSynoptique.Enabled = True
    


    
    Exit Sub

'--- gestion des erreurs ---
GestionErreurs:
      
    '--- affichage du message d'erreur ---
    MessageErreur TITRE_MESSAGES, Err.Description, Err.Number
    
End Sub

Private Sub IMProgrammateurCycliqueCuves_Click(Index As Integer)
    On Error Resume Next
    AppelFenetre FENETRES.F_PROGRAMMATEUR_CYCLIQUE
End Sub

Private Sub LEtatsAnnexes_Click(Index As Integer)
    On Error Resume Next
    AppelFenetre FENETRES.F_ANNEXES
End Sub

Private Sub LLibellesAnnexes_Click(Index As Integer)
    On Error Resume Next
    AppelFenetre FENETRES.F_ANNEXES
End Sub

Private Sub LLibellesPostes_Click(Index As Integer)
   On Error Resume Next
    Dim NumCuve As Integer
    NumCuve = CorrespondancePostesCuvesAPI(Index)
    If NumCuve > 0 Then AppelFenetre FENETRES.F_GESTION_CUVES, NumCuve
End Sub

Private Sub LNumCharges_Click(Index As Integer)
    On Error Resume Next
    AppelFenetre FENETRES.F_CHARGES_EN_LIGNE, TEtatsPostes(Index).NumCharge
End Sub

Private Sub LTemperatures_Click(Index As Integer)
    On Error Resume Next
    AppelFenetre FENETRES.F_GESTION_REGULATION, Index
End Sub

Private Sub OBChoixBaseDeDonnees_Click(Index As Integer)
    
    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next

    '--- affectation du type de base de données ---
    TypeBD = Index

End Sub

Private Sub OBModeAffichageSynoptique_Click(Index As Integer)

    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next

    '--- affectation du mode d'affichage ---
    ModeAffichageSynoptique = Index

End Sub

Private Sub OCXRedresseurs_Click(Index As Integer)
    On Error Resume Next
    AppelFenetre FENETRES.F_GESTION_REDRESSEURS, Index
End Sub

Private Sub PBBarreAgrandissement_DblClick(Index As Integer)
    
    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next
    
    '--- rafraichissement ---
    CBAgrandirRestaurerZonesFenetre(Index).Refresh
        
    '--- changement de la forme de la fenetre ---
    If FormeFenetre <> FORMES_FENETRE.F_STANDARD Then
        ChangeFormeFenetre FORMES_FENETRE.F_STANDARD
    Else
        Select Case Index
            Case ZONES_FENETRE.Z_SYNOPTIQUE: ChangeFormeFenetre F_SYNOPTIQUE
            Case ZONES_FENETRE.Z_DIALOGUES: ChangeFormeFenetre F_DIALOGUES
            Case Else
        End Select
    End If

End Sub

Private Sub IEtatsPonts_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next

    '--- condamnation du pont par un clic droit de la souris ---
    If Button = vbRightButton Then
        CondamnationPont Index, TITRE_MESSAGES
    End If

End Sub

Private Sub IEtatsPostes_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next

    '--- déclaration ---
    Dim NumCuve As Integer
    
    If Button = vbLeftButton Then
        
        '--- appel de la fenêtre gérant les cuves ---
        NumCuve = CorrespondancePostesCuvesAPI(Index)
        If NumCuve > 0 Then AppelFenetre FENETRES.F_GESTION_CUVES, NumCuve
    
    Else
        
        '--- condamnation du poste par un clic droit de la souris ---
        CondamnationPoste Index, TITRE_MESSAGES
    
    End If

End Sub


Private Sub calculCoordEnBase(X As Single)

    Dim pixLibL, pixLibR As Integer
    Dim pixPosteL, pixPosteR As Integer
    
    pixLibL = X - 15
    pixLibR = X + 15
    pixPosteL = X - 21
    pixPosteR = X + 21
    
    If (NbPosteClicked < 5) Then
        pixLibL = X - 15
        pixLibR = X + 15
    End If
    
    If (NbPosteClicked = 42 Or NbPosteClicked = 43) Then
        pixLibL = X - 13
        pixLibR = X + 13
    End If
    
    Dim ConnexionBDAnodisationSQL As New ADODB.Connection
     With ConnexionBDAnodisationSQL
        .ConnectionString = PARAMETRES_CONNEXION_BD_ANODISATION_SQL
        .CursorLocation = adUseServer
        .Mode = adModeReadWrite
        .ConnectionTimeout = 2     'X secondes d'attente de connexion avant de lancer un message d'erreur
        .Open
    End With
    Dim Enregistrement As New ADODB.Recordset
    Dim Requete As String
    
    With Enregistrement
              
              '--- lancement d'une requête ---
            .CursorLocation = adUseServer
            .CursorType = adOpenKeyset
            .LockType = adLockOptimistic
            Requete = "UPDATE POSTES SET  XAxePosteSynoptique=" & X & "," & _
                    "  XInferieurPosteSynoptique=" & pixPosteL & "," & _
                    "  XSuperieurPosteSynoptique=" & pixPosteR & "," & _
                    "  XInferieurLibellePosteSynoptique=" & pixLibL & "," & _
                    "  XSuperieurLibellePosteSynoptique=" & pixLibR & _
                    "WHERE Numposte=" & NbPosteClicked
            .Open Requete, PARAMETRES_CONNEXION_BD_ANODISATION_SQL, , adCmdText
        
    End With
    
    
    NbPosteClicked = NbPosteClicked + 1
    
        
        
End Sub
Private Sub PBImageLigne_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next
    ' moyen pas top d'adapter les coordonnées
    ' a l'avenir refaire un synoptique avec 2*20 pixel de large pour chacune des 45 cuves => 1800 px /1877 dispo. x left première cuve=36px (77/2)
    'calculCoordEnBase (X)
   
    '--- déclaration ---
    Dim a As Integer
    Dim NumPontClique, _
           NumPosteClique As Integer, _
           NumLibellePosteClique As Integer, _
           NumCuve As Integer, _
           NumCharge As Integer
    
    '--- recherche de la partie du synoptique cliqué pour les ponts ---
    For a = PONTS.P_1 To PONTS.P_2
        If X >= TXPonts(a) And Y >= TYPonts(a) And X <= (TXPonts(a) + DIMENSIONS_ANIMATIONS.D_LONG_PONT) And Y <= (TYPonts(a) + DIMENSIONS_ANIMATIONS.D_HAUT_PONT) Then
            NumPontClique = a
            Exit For
        End If
    Next a
    
    '--- recherche de la partie du synoptique cliqué pour les postes ---
    For a = POSTES.P_CHGT_1 To DERNIER_POSTE
        With TEtatsPostes(a).DefinitionPoste
            
            '--- recherche si poste cliqué ---
            If X >= .XInferieurPosteSynoptique And Y >= .YInferieurPosteSynoptique And X <= .XSuperieurPosteSynoptique And Y <= .YSuperieurPosteSynoptique Then
                NumPosteClique = a
                Exit For
            End If
            
            '--- recherche si libellé du poste cliqué ---
            If X >= .XInferieurLibellePosteSynoptique And Y >= .YInferieurLibellePosteSynoptique And X <= .XSuperieurLibellePosteSynoptique And Y <= .YSuperieurLibellePosteSynoptique Then
                NumLibellePosteClique = a
                Exit For
            End If
        
        End With
    Next a
    
    If Button = vbLeftButton Then
        
        '******************************************************************************************************
        '*                                                  ANALYSE SUR LE PONT CLIQUE
        '******************************************************************************************************
        If NumPontClique >= PONTS.P_1 And NumPontClique <= PONTS.P_2 Then
        
            With TEtatsPonts(NumPontClique)
        
                '--- affectation du numéro de charge ---
                If .NumCharge >= CHARGES.C_NUM_MINI And .NumCharge <= CHARGES.C_NUM_MAXI Then
                    NumCharge = .NumCharge
                    AppelFenetre FENETRES.F_CHARGES_EN_LIGNE, NumCharge
                End If
        
            End With
        
        End If
        
        '******************************************************************************************************
        '*                                                 ANALYSE SUR LE POSTE CLIQUE
        '******************************************************************************************************
        If NumPosteClique >= POSTES.P_CHGT_1 And NumPosteClique <= DERNIER_POSTE Then
        
            With TEtatsPostes(NumPosteClique)
        
                '--- affectation du numéro de charge ---
                If .NumCharge >= CHARGES.C_NUM_MINI And .NumCharge <= CHARGES.C_NUM_MAXI Then
                    NumCharge = .NumCharge
                End If

                '--- postes de chargement ---
                If NumPosteClique >= POSTES.P_CHGT_1 And NumPosteClique <= POSTES.P_CHGT_2 Then
                    AppelFenetre FENETRES.F_CHARGEMENT_PREVISIONNEL
                End If
        
                
        
                '--- postes de déchargement ---
                If NumPosteClique >= POSTES.P_D1 And NumPosteClique <= POSTES.P_D2 And NumCharge > 0 Then
                    AppelFenetre FENETRES.F_CHARGES_EN_LIGNE, NumCharge
                Else
                    '--- postes de traitement ---
                    If NumPosteClique >= PREMIER_BAIN And NumPosteClique <= DERNIER_POSTE And NumCharge > 0 Then
                        AppelFenetre FENETRES.F_CHARGES_EN_LIGNE, NumCharge
                    End If
                End If
        
            End With
        
        End If
        
        '******************************************************************************************************
        '*                                                 ANALYSE SUR LE LIBELLE CLIQUE
        '******************************************************************************************************
        If NumLibellePosteClique >= POSTES.P_CHGT_1 And NumLibellePosteClique <= DERNIER_POSTE Then
    
            With TEtatsPostes(NumLibellePosteClique)
            
                '--- affectation du numéro de charge ---
                If .NumCharge >= CHARGES.C_NUM_MINI And .NumCharge <= CHARGES.C_NUM_MAXI Then
                    NumCharge = .NumCharge
                End If
                
                '--- postes de chargement ---
                If NumLibellePosteClique >= POSTES.P_CHGT_1 And NumLibellePosteClique <= POSTES.P_CHGT_2 Then
                    AppelFenetre FENETRES.F_CHARGEMENT_PREVISIONNEL
                End If
              
                
                '--- postes de déchargement ---
                If NumLibellePosteClique >= POSTES.P_D1 And NumLibellePosteClique <= POSTES.P_D2 And NumCharge > 0 Then
                    AppelFenetre FENETRES.F_CHARGES_EN_LIGNE, NumCharge
                Else
                    'cuve
                    If NumLibellePosteClique >= PREMIER_BAIN And NumLibellePosteClique <= DERNIER_POSTE Then
                        NumCuve = CorrespondancePostesCuvesAPI(NumLibellePosteClique)
                        'If NumCuve > 0 Then AppelFenetre FENETRES.F_GESTION_CUVES, NumCuve
                    End If
                End If
            
            End With
        
        End If

    Else

        '--- condamnation du poste par un clic droit de la souris ---
        If NumLibellePosteClique >= POSTES.P_CHGT_1 And NumLibellePosteClique <= DERNIER_POSTE Then
            CondamnationPoste NumLibellePosteClique, TITRE_MESSAGES
        End If

    End If

End Sub

Private Sub RTBDialogues_KeyPress(KeyAscii As Integer)
    
    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next
    
    '--- analyse de la touche frappée ---
    If ModeDialoguesEnCours = MODES_DIALOGUES.M_QUESTIONS_REPONSES Then
        Me.RTBDialogues.SelColor = COULEURS.VERT_5
        Select Case KeyAscii
            Case vbKeyReturn
                KeyAscii = 0
            Case Else
                FiltreToucheASCII KeyAscii, DONNEES.D_GENERALE_MAJUSCULES
        End Select
    Else
        KeyAscii = 0
    End If

End Sub

Private Sub RTBDialogues_KeyUp(KeyCode As Integer, Shift As Integer)
    
    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next
    
    '--- analyse de la touche frappée ---
    If ModeDialoguesEnCours = MODES_DIALOGUES.M_QUESTIONS_REPONSES Then
        Select Case KeyCode
            Case vbKeyReturn
                ExtractionQuestion
                KeyCode = 0
            Case Else
        End Select
    End If

End Sub

'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Rôle      : Visualise tous les états du synoptique
' Détails  :
' Retours :
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Sub VisualisationEtatsSynoptique()
    
    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next

    '--- déclaration ---
    Dim RestaurerSynoptique As Boolean
    
    Dim a As Integer                                                                                                         'pour les boucles FOR...NEXT
    Dim b As Integer                                                                                                         'pour les boucles FOR...NEXT
    
    Dim PosteReference As Integer                                                                                 'représente un poste de référence pour les calculs
    
    Dim TRapportLaserTrlSynoptique(PONTS.P_1 To PONTS.P_2) As Long                  'rapport entre la valeur laser de la translation et les points du synoptique
    Dim TRapportCodeurLevPontSynoptique(PONTS.P_1 To PONTS.P_2) As Long       'rapport entre la valeur codeur du levage et les points du synoptique
    Dim RapportCodeurElevateurSynoptique As Long                                                    'rapport entre la valeur codeur de l'élévateur et les points du synoptique
    Dim XAxePosteLigne As Long                                                                                    'axe de poste laser sur la ligne
    Dim XAxePosteSynoptique As Long                                                                           'axe de poste laser sur la ligne
    Static TMemPositionsSynoptique(PONTS.P_1 To PONTS.P_2) As Long                   'mémoire de position sur le synoptique
    Dim PositionActuelleLaserTrlPont As Long                                                                'position actuel du laser de la translation des ponts
    
    Dim RapportLaserSynoptique As Double                                                                   'rapport entre la valeur laser et les points du synoptique
    
    Dim TCopieEtatsPonts() As EtatsPonts
    Dim TCopieEtatsPostes() As EtatsPostes
    
    '-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    
    '--- copie des états des ponts ---
    TCopieEtatsPonts() = TEtatsPonts()
    
    '--- copie des états des postes ---
    TCopieEtatsPostes() = TEtatsPostes()
    
    '-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    
    '--- teste de la coopération avec Windows ---
    RestaurerSynoptique = False
    Do Until TesteNiveauCooperation
        DoEvents
        RestaurerSynoptique = True
    Loop
    
    '-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    
    '--- restauration complète des surfaces ---
    If RestaurerSynoptique = True Then
        RestaurerSynoptique = False
        ObjDD.RestoreAllSurfaces
        InitialisationSurfaces
    End If
    
    '-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    
    '--- effacement par ponts ---
    For a = PONTS.P_1 To PONTS.P_2
        
        '--- effacement des ponts ---
        AfficheAnimations N_CHARGE_PONT, False, a
        AfficheAnimations N_ACCROCHE, False, a
        AfficheAnimations N_PALONNIER, False, a
        AfficheAnimations N_PONT, False, a
        
        '--- effacement des ponts fictifs ---
        AfficheAnimations N_PONT_FICTIF, False, a, CHARGES.PAS_DE_CHARGE
    
    Next a
    
    '-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    
    '--- effacements par poste et reconstruction immédiat dans l'ordre de création ---
    For a = POSTES.P_CHGT_1 To DERNIER_POSTE
        
        '--- effacement ---
        VisualisationLibelles TCopieEtatsPostes(a), False
        VisualisationCharges TCopieEtatsPostes(a), False
        VisualisationChariots TCopieEtatsPostes(a), False
        VisualisationCouvercles TCopieEtatsPostes(a), False
        
        '--- reconstruction ---
        VisualisationLibelles TCopieEtatsPostes(a), True
        VisualisationChariots TCopieEtatsPostes(a), True
        VisualisationCharges TCopieEtatsPostes(a), True
        VisualisationCouvercles TCopieEtatsPostes(a), True
    
    Next a
    
    '-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    
    '--- analyse complète des ponts ---
    For a = PONTS.P_1 To PONTS.P_2
        
        '--- analyse complète pour le pont indexé ---
        With TCopieEtatsPonts(a)
        
            '-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
            
            '--- affectation de la position du laser gauche ---
            PositionActuelleLaserTrlPont = .PositionActuelleLaserTrlPont
            
            '-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
            
            '----------------------------------------------------------------------
            '--- coordonnées et affichage du pont fictif ---
            '----------------------------------------------------------------------
            If .ModePont <> MODES_PONTS.M_MAINTENANCE Then
                If .PosteDestination >= POSTES.P_CHGT_1 And .PosteDestination <= DERNIER_POSTE Then
                    TXPontsFictifs(a) = TEtatsPostes(.PosteDestination).DefinitionPoste.XAxePosteSynoptique - DIMENSIONS_ANIMATIONS.D_AXE_PONT
                    TDerniersXPontsFictifs(a) = TXPontsFictifs(a)
                    TDerniersYPontsFictifs(a) = TYPontsFictifs(a)
                    If a = PONTS.P_1 Then
                        AfficheAnimations N_PONT_FICTIF, True, a, TYPES_AFFICHAGES_PONTS.T_FICTIF_P1            'transfert de l'image concerné dans l'image tampon
                    Else
                        AfficheAnimations N_PONT_FICTIF, True, a, TYPES_AFFICHAGES_PONTS.T_FICTIF_P2            'transfert de l'image concerné dans l'image tampon
                    End If
                End If
            End If

            '-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
            
            '------------------------------------------
            '--- coordonnées du pont ---
            '------------------------------------------
            If .PosteActuel > 0 Then
    
                CoordReelle = PositionActuelleLaserTrlPont
                Select Case .SensX
    
                    Case SENS_X.S_AU_POSTE
                        '--- pont au poste ---
                        TXPonts(a) = TEtatsPostes(.PosteActuel).DefinitionPoste.XAxePosteSynoptique - DIMENSIONS_ANIMATIONS.D_AXE_PONT
                        TMemPositionsSynoptique(a) = TXPonts(a)
    
                    Case SENS_X.S_AVANT
                        '--- sens avant ---
                        '--- calcul du rapport entre le laser et le synoptique ---
                        If .PosteActuel + 1 < .PosteDestination Then
                            PosteReference = .PosteActuel + 1
                        Else
                            PosteReference = .PosteDestination
                        End If
                        XAxePosteLigne = TEtatsPostes(PosteReference).DefinitionPoste.XAxePosteLigne
                        XAxePosteSynoptique = TEtatsPostes(PosteReference).DefinitionPoste.XAxePosteSynoptique
                        
                        '--- calcul du rapport ---
                        TRapportLaserTrlSynoptique(a) = (PositionActuelleLaserTrlPont * XAxePosteSynoptique) / XAxePosteLigne - DIMENSIONS_ANIMATIONS.D_AXE_PONT
                        
                        '--- affectation des points graphiques ---
                        If TRapportLaserTrlSynoptique(a) < TMemPositionsSynoptique(a) Then
                            TXPonts(a) = TMemPositionsSynoptique(a) - 1
                            If (.ModePont = MODES_PONTS.M_SEMI_AUTOMATIQUE Or .ModePont = MODES_PONTS.M_AUTOMATIQUE) And .PosteDestination > 0 Then             'compensation pour les à coups
                                If TXPonts(a) <= TEtatsPostes(.PosteDestination).DefinitionPoste.XAxePosteSynoptique - DIMENSIONS_ANIMATIONS.D_AXE_PONT Then
                                    TXPonts(a) = TEtatsPostes(.PosteDestination).DefinitionPoste.XAxePosteSynoptique - DIMENSIONS_ANIMATIONS.D_AXE_PONT
                                End If
                            End If
                            TMemPositionsSynoptique(a) = TXPonts(a)
                            
                            'TextInfo = " Pos synoptique: " & TMemPositionsSynoptique(a)
                        'Else
                        '    TextInfo = " RIEN !!!: "
                        End If
    
                    Case SENS_X.S_ARRIERE
                        '--- sens arrière ---
                        '--- calcul du rapport en le laser er le synoptique ---
                        If .PosteActuel - 1 > .PosteDestination Then
                            PosteReference = .PosteActuel - 1
                        Else
                            PosteReference = .PosteDestination
                        End If
                        XAxePosteLigne = TEtatsPostes(PosteReference).DefinitionPoste.XAxePosteLigne
                        XAxePosteSynoptique = TEtatsPostes(PosteReference).DefinitionPoste.XAxePosteSynoptique
                        If XAxePosteLigne = 0 Then XAxePosteLigne = 1   'pour éviter la division par 0 (cas normallement impossible en marche avant)
                        
                        '--- calcul du rapport ---
                        TRapportLaserTrlSynoptique(a) = (PositionActuelleLaserTrlPont * XAxePosteSynoptique) / XAxePosteLigne - DIMENSIONS_ANIMATIONS.D_AXE_PONT
                        
                        '--- affectation des points graphiques ---
                        If TRapportLaserTrlSynoptique(a) > TMemPositionsSynoptique(a) Then
                           TXPonts(a) = TMemPositionsSynoptique(a) + 1
                           If (.ModePont = MODES_PONTS.M_SEMI_AUTOMATIQUE Or .ModePont = MODES_PONTS.M_AUTOMATIQUE) And .PosteDestination > 0 Then              'compensation pour les à coups
                               If TXPonts(a) >= TEtatsPostes(.PosteDestination).DefinitionPoste.XAxePosteSynoptique - DIMENSIONS_ANIMATIONS.D_AXE_PONT Then
                                   TXPonts(a) = TEtatsPostes(.PosteDestination).DefinitionPoste.XAxePosteSynoptique - DIMENSIONS_ANIMATIONS.D_AXE_PONT
                               End If
                           End If
                           TMemPositionsSynoptique(a) = TXPonts(a)
                           
                           'TextInfo = " Pos synoptique: " & TMemPositionsSynoptique(a)
                        'Else
                        '   TextInfo = " RIEN !!!: "
                        End If
                    
                    Case Else
    
                End Select
                
            End If
            
            '-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
            
            If .ModePont <> MODES_PONTS.M_MAINTENANCE Then
        
                '---------------------------------------------------------------------
                '--- coordonnées d'une charge sur le pont ---
                '---------------------------------------------------------------------
                 
                '--- position sur l'axe des x ---
                TXChargesPonts(a) = TXPonts(a) + DIMENSIONS_ANIMATIONS.D_AXE_PONT - DIMENSIONS_ANIMATIONS.D_AXE_CHARGE
                 
                '-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
                 
                '--- position sur l'axe des y pour le NIVEAU BAS ---
                If .TEntreesAPI.E_NiveauBas = True Or .PositionActuelleCodeurLevPont <= 0 Then
                     
                    '--- coordonnées du palonnier ---
                    TXPalonniers(a) = TXPonts(a) + POSITIONS_ORIGINES.AJOUT_X_PALONNIER
                    TYPalonniers(a) = TYPonts(a) + POSITIONS_ORIGINES.AJOUT_Y_PALONNIER_BAS_PONT
                     
                    '--- coordonnées de l'accroche ---
                    TXAccroches(a) = TXPonts(a) + POSITIONS_ORIGINES.AJOUT_X_ACCROCHE
                    TYAccroches(a) = TYPalonniers(a) + DIMENSIONS_ANIMATIONS.D_HAUT_PALONNIER
                     
                    '--- sur capteur niveau bas ---
                    TYChargesPonts(a) = POSITIONS_ORIGINES.PO_Y_CHARGE_BAS_PONT
            
                '-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
                 
                '--- position sur l'axe des y pour le NIVEAU HAUT ---
                ElseIf .TEntreesAPI.E_NiveauHaut = True Or .PositionActuelleCodeurLevPont >= VALEUR_CODEUR_NIVEAU_HAUT_PONTS Then
                     
                    '--- coordonnées du palonnier ---
                    TXPalonniers(a) = TXPonts(a) + POSITIONS_ORIGINES.AJOUT_X_PALONNIER
                    TYPalonniers(a) = TYPonts(a)
                     
                    '--- coordonnées de l'accroche ---
                    TXAccroches(a) = TXPonts(a) + POSITIONS_ORIGINES.AJOUT_X_ACCROCHE
                    TYAccroches(a) = TYPalonniers(a) + DIMENSIONS_ANIMATIONS.D_HAUT_PALONNIER
                     
                    '--- sur capteur niveau haut ---
                    TYChargesPonts(a) = POSITIONS_ORIGINES.PO_Y_CHARGE_HAUT_PONT
                     
                '-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
                 
                '--- position sur l'axe des y avec le CODEUR ---
                Else
        
                    '--- calcul du rapport entre la valeur codeur du levage et les points du synoptique ---
                    TRapportCodeurLevPontSynoptique(a) = .PositionActuelleCodeurLevPont * (POSITIONS_ORIGINES.PO_Y_CHARGE_BAS_PONT - POSITIONS_ORIGINES.PO_Y_CHARGE_HAUT_PONT) / VALEUR_CODEUR_NIVEAU_HAUT_PONTS
                    TYChargesPonts(a) = POSITIONS_ORIGINES.PO_Y_CHARGE_BAS_PONT - TRapportCodeurLevPontSynoptique(a)
                    
                    '--- niveau sur codeur (calcul de la position y avec limites mini/maxi) ---
                    If TYChargesPonts(a) > POSITIONS_ORIGINES.PO_Y_CHARGE_BAS_PONT Then
                        TYChargesPonts(a) = POSITIONS_ORIGINES.PO_Y_CHARGE_BAS_PONT
                    End If
                    If TYChargesPonts(a) < POSITIONS_ORIGINES.PO_Y_CHARGE_HAUT_PONT Then
                        TYChargesPonts(a) = POSITIONS_ORIGINES.PO_Y_CHARGE_HAUT_PONT
                    End If
                    
                    '--- coordonnées l'accroche ---
                    TXAccroches(a) = TXPonts(a) + POSITIONS_ORIGINES.AJOUT_X_ACCROCHE
                    TYAccroches(a) = TYChargesPonts(a) - DIMENSIONS_ANIMATIONS.D_HAUT_ACCROCHE + 5     '+5 pour emboiter l'accroche
                    
                    '--- coordonnées du palonnier ---
                    TXPalonniers(a) = TXPonts(a) + POSITIONS_ORIGINES.AJOUT_X_PALONNIER
                    TYPalonniers(a) = TYAccroches(a) - DIMENSIONS_ANIMATIONS.D_HAUT_PALONNIER
                 
                 End If
        
                '-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
                
                '----------------------------------------------------
                '--- affichage du pont concerné ---
                '----------------------------------------------------
                TDerniersXPonts(a) = TXPonts(a)
                TDerniersYPonts(a) = TYPonts(a)
                If .Condamnation = True Then
                    AfficheAnimations N_PONT, True, a, TYPES_AFFICHAGES_PONTS.T_CONDAMNATION
                Else
                    If .UnDefautAuMoinsSignale = True Then
                        If ClignotantPourSynoptique = False Then
                            AfficheAnimations N_PONT, True, a, TYPES_AFFICHAGES_PONTS.T_DEFAUT_BLANC
                        Else
                            AfficheAnimations N_PONT, True, a, TYPES_AFFICHAGES_PONTS.T_DEFAUT_ROUGE
                        End If
                    Else
                        AfficheAnimations N_PONT, True, a, TYPES_AFFICHAGES_PONTS.T_NORMAL
                    End If
                End If
       
                '-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
                
                '------------------------------------------------------------
                '--- affichage du palonnier concerné ---
                '------------------------------------------------------------
                TDerniersXPalonniers(a) = TXPalonniers(a)
                TDerniersYPalonniers(a) = TYPalonniers(a)
                If .Condamnation = True Then
                    AfficheAnimations NOMS_ANIMATIONS.N_PALONNIER, True, a, TYPES_AFFICHAGES_PALONNIERS.T_CONDAMNATION
                Else
                    If .UnDefautAuMoinsSignale = True Then
                        If ClignotantPourSynoptique = False Then
                            AfficheAnimations NOMS_ANIMATIONS.N_PALONNIER, True, a, TYPES_AFFICHAGES_PALONNIERS.T_DEFAUT_BLANC
                        Else
                            AfficheAnimations NOMS_ANIMATIONS.N_PALONNIER, True, a, TYPES_AFFICHAGES_PALONNIERS.T_DEFAUT_ROUGE
                        End If
                    Else
                        AfficheAnimations NOMS_ANIMATIONS.N_PALONNIER, True, a, TYPES_AFFICHAGES_PALONNIERS.T_NORMAL
                    End If
                End If
                
                '-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
                
                '-------------------------------------------------------------
                '--- affichage de l'accroche concerné ---
                '-------------------------------------------------------------
                TDerniersXAccroches(a) = TXAccroches(a)
                TDerniersYAccroches(a) = TYAccroches(a)
                If .Condamnation = True Then
                    AfficheAnimations NOMS_ANIMATIONS.N_ACCROCHE, True, a, TYPES_AFFICHAGES_ACCROCHES.T_CONDAMNATION
                Else
                    If .UnDefautAuMoinsSignale = True Then
                        If ClignotantPourSynoptique = False Then
                            AfficheAnimations NOMS_ANIMATIONS.N_ACCROCHE, True, a, TYPES_AFFICHAGES_ACCROCHES.T_DEFAUT_BLANC
                        Else
                            AfficheAnimations NOMS_ANIMATIONS.N_ACCROCHE, True, a, TYPES_AFFICHAGES_ACCROCHES.T_DEFAUT_ROUGE
                        End If
                    Else
                        If .EtatsAccrochesCharge = ETATS_ACCROCHES.E_ACCROCHES_EN_BAS Then
                            AfficheAnimations NOMS_ANIMATIONS.N_ACCROCHE, True, a, TYPES_AFFICHAGES_ACCROCHES.T_NORMAL
                        End If
                    End If
                End If
                
                '-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
                
                '--------------------------------------------------------------
                '--- affichage de la charge concernée ---
                '--------------------------------------------------------------
                If .NumCharge > 0 Then
                    TDerniersXChargesPonts(a) = TXChargesPonts(a)
                    TDerniersYChargesPonts(a) = TYChargesPonts(a)
                    If .EtatsAccrochesCharge = ETATS_ACCROCHES.E_ACCROCHES_EN_BAS Then
                        If .NumCharge > 0 Then
                            AfficheAnimations N_CHARGE_PONT, True, a, .NumCharge        'numéro de la charge
                        End If
                    End If
                End If
        
            Else
                    
                '--- positions des ponts en hors ligne ---
                If a = PONTS.P_1 Then
                    TXPonts(a) = POSITIONS_ORIGINES.PO_X_PONT_1_HORS_LIGNE
                Else
                    TXPonts(a) = POSITIONS_ORIGINES.PO_X_PONT_2_HORS_LIGNE
                End If
                TDerniersXPonts(a) = TXPonts(a)
                TDerniersYPonts(a) = TYPonts(a)
                
                '--- pont en mode hors ligne ---
                AfficheAnimations N_PONT, True, a, TYPES_AFFICHAGES_PONTS.T_CONDAMNATION
        
            End If
        
        End With
        
    Next a
    
    '--- affichage de l'image tampon à l'écran ---
    GestionImageTampon True
        
End Sub

'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Rôle      : Visualisation d'une charge
' Détails  :
' Retours :
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Private Sub VisualisationCharges(ByRef TEtatsUnPoste As EtatsPostes, _
                                                       ByVal EffacerOuActualiser As Boolean)
    
    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next

    '--- déclaration ---
    Dim NumPoste As Integer
    Static TCopieEtatsUnPoste As EtatsPostes
    
    With TEtatsUnPoste
        
        Select Case .DefinitionPoste.NumPoste
            
            Case POSTES.P_CHGT_1 To DERNIER_POSTE
                '--- postes de chargement, ligne alu, postes de déchargement ---
                TXChargePoste = .DefinitionPoste.XAxePosteSynoptique - DIMENSIONS_ANIMATIONS.D_AXE_CHARGE
                TYChargePoste = POSITIONS_ORIGINES.PO_Y_CHARGE_POSTE
                TDernierXChargePoste = TXChargePoste
                TDernierYChargePoste = TYChargePoste

                '--- effacer ou actualiser ---
                If EffacerOuActualiser = False Or (EffacerOuActualiser = True And .NumCharge > 0) Then
                    AfficheAnimations N_CHARGE_POSTE, EffacerOuActualiser, , .NumCharge
                End If
                
            Case Else

        End Select

    End With

End Sub

'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Rôle      : Visualisation des chariots
' Détails  :
' Retours :
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Private Sub VisualisationChariots(ByRef TEtatsUnPoste As EtatsPostes, _
                                                       ByVal EffacerOuActualiser As Boolean)
    
    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next

    '--- constantes privées ---
    Const AJOUT_DE_POSITION As Integer = 4

    '--- déclaration ---
    Dim NumPoste As Integer
    
    With TEtatsUnPoste
        
        Select Case .DefinitionPoste.NumPoste
            
            Case POSTES.P_CHGT_1 To POSTES.P_CHGT_2, POSTES.P_D1 To POSTES.P_D2
                '--- postes de chargement et déchargement ---
                TXChariot = .DefinitionPoste.XAxePosteSynoptique - DIMENSIONS_ANIMATIONS.D_AXE_CHARIOT
                TYChariot = POSITIONS_ORIGINES.PO_Y_CHARIOT
                TDernierXChariot = TXChariot
                TDernierYChariot = TYChariot

                '--- effacer ou actualiser ---
                If EffacerOuActualiser = False Or _
                   ((.EtatsChariots = ETATS_CHARIOTS.E_PRESENT_VERROUILLE Or .EtatsChariots = ETATS_CHARIOTS.E_PRESENT) And EffacerOuActualiser = True) Then
                    AfficheAnimations N_CHARIOTS, EffacerOuActualiser, .EtatsChariots
                End If
                
            Case Else

        End Select

    End With

End Sub

'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Rôle      : Visualisation des couvercles
' Détails  :
' Retours :
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Private Sub VisualisationCouvercles(ByRef EtatsUnPoste As EtatsPostes, _
                                                            ByVal EffacerOuActualiser As Boolean)
    
    '--- aiguillage en cas d'erreurs ---
'    On Error Resume Next
'
'    '--- constantes privées ---
'    Const DELAI_ENTRE_IMAGES_COUVERCLES_EN_OUVERTURE As Long = 1
'    Const DELAI_ENTRE_IMAGES_COUVERCLES_EN_FERMETURE As Long = 2
'
'    '--- déclaration ---
'    Dim ModifierImage As Boolean
'    Static TIndexImagesCouvercles(POSTES.P_CHGT_1 To POSTES.P_D2) As Integer
'    Static TVitessesAnimations(POSTES.P_CHGT_1 To POSTES.P_D2) As VitesseAnimations
'
'    With EtatsUnPoste
'
'        Select Case .DefinitionPoste.NumPoste
'
'            Case POSTES.P_A1, POSTES.P_A2, POSTES.P_C13, POSTES.P_C14, POSTES.P_A17, POSTES.P_A18, _
'                     POSTES.P_B1, POSTES.P_B4, POSTES.P_B5, POSTES.P_C15
'
'                '--- affectation ---
'                TXCouvercles = .DefinitionPoste.XAxePosteSynoptique - DIMENSIONS_ANIMATIONS.D_AXE_COUVERCLES
'                Select Case .DefinitionPoste.NumPoste
'                    Case POSTES.P_B1, POSTES.P_B4, POSTES.P_B5, POSTES.P_C15
'                        TYCouvercles = AJOUTS_COUVERCLES.A_Y_LIGNE_ALU
'                    Case Else
'                        TYCouvercles = AJOUTS_COUVERCLES.A_Y_LIGNE_ACIER
'                End Select
'                TDernierXCouvercles = TXCouvercles
'                TDernierYCouvercles = TYCouvercles
'
'                '--- détermination de l'image ---
'                Select Case .EtatsCouvercles
'
'                    Case ETATS_COUVERCLES.E_COUVERCLES_OUVERTS
'                        '--- couvercles ouverts ---
'                        AfficheAnimations N_COUVERCLES, EffacerOuActualiser, , 4
'                        TIndexImagesCouvercles(.DefinitionPoste.NumPoste) = 4
'
'                    Case ETATS_COUVERCLES.E_COUVERCLES_FERMES
'                        '--- couvercles fermés ---
'                        AfficheAnimations N_COUVERCLES, EffacerOuActualiser, , 0
'                       TIndexImagesCouvercles(.DefinitionPoste.NumPoste) = 0
'
'                    Case ETATS_COUVERCLES.E_COUVERCLES_EN_FERMETURE
'                        '--- couvercles en fermeture ---
'                        ModifierImage = (DateDiff("s", TVitessesAnimations(.DefinitionPoste.NumPoste).DateDernierPassage, Now) >= DELAI_ENTRE_IMAGES_COUVERCLES_EN_FERMETURE) Or _
'                                                     TVitessesAnimations(.DefinitionPoste.NumPoste).PremierPassage = False
'                        If ModifierImage = True Then
'                            Dec TIndexImagesCouvercles(.DefinitionPoste.NumPoste)
'                            TVitessesAnimations(.DefinitionPoste.NumPoste).DateDernierPassage = Now
'                        End If
'                        If TIndexImagesCouvercles(.DefinitionPoste.NumPoste) < 1 Then
'                            TIndexImagesCouvercles(.DefinitionPoste.NumPoste) = 1
'                        End If
'                        AfficheAnimations N_COUVERCLES, EffacerOuActualiser, , TIndexImagesCouvercles(.DefinitionPoste.NumPoste)
'
'                    Case ETATS_COUVERCLES.E_COUVERCLES_EN_OUVERTURE
'                        '--- couvercles en ouverture ---
'                        ModifierImage = (DateDiff("s", TVitessesAnimations(.DefinitionPoste.NumPoste).DateDernierPassage, Now) >= DELAI_ENTRE_IMAGES_COUVERCLES_EN_OUVERTURE) Or _
'                                                    TVitessesAnimations(.DefinitionPoste.NumPoste).PremierPassage = False
'                        If ModifierImage = True Then
'                            Inc TIndexImagesCouvercles(.DefinitionPoste.NumPoste)
'                            TVitessesAnimations(.DefinitionPoste.NumPoste).DateDernierPassage = Now
'                        End If
'                        If TIndexImagesCouvercles(.DefinitionPoste.NumPoste) > 3 Then
'                            TIndexImagesCouvercles(.DefinitionPoste.NumPoste) = 3
'                        End If
'                        AfficheAnimations N_COUVERCLES, EffacerOuActualiser, , TIndexImagesCouvercles(.DefinitionPoste.NumPoste)
'
'                    Case Else
'
'                End Select
'
'                '--- valeur au premier passage ---
'                With TVitessesAnimations(.DefinitionPoste.NumPoste)
'                    If .PremierPassage = False Then
'                        .DateDernierPassage = Now
'                        .PremierPassage = True
'                    End If
'                End With
'
'            Case Else
'
'        End Select
'
'    End With

End Sub

'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Rôle      : Visualisation des libellés
' Détails  :
' Retours :
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Private Sub VisualisationLibelles(ByRef EtatsUnPoste As EtatsPostes, _
                                                       ByVal EffacerOuActualiser As Boolean)
    
    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next

    '--- constantes privées ---
    
    '--- déclaration ---
    Dim NumCuve As Integer
    
    With EtatsUnPoste
        
        '--- affectation ---
        TXLibelle = .DefinitionPoste.XInferieurLibellePosteSynoptique
        TYLibelle = .DefinitionPoste.YInferieurLibellePosteSynoptique
        TDernierXLibelle = TXLibelle
        TDernierYLibelle = TYLibelle
        
        '--- effacer ou actualiser ---
        If EffacerOuActualiser = False Then
            AfficheAnimations N_LIBELLES, EffacerOuActualiser, .DefinitionPoste.NumPoste
        Else
            If .Condamnation = True Then
                AfficheAnimations N_LIBELLES, EffacerOuActualiser, .DefinitionPoste.NumPoste, TYPES_AFFICHAGES_LIBELLES.E_CONDAMNATION
            Else
                NumCuve = CorrespondancePostesCuvesAPI(.DefinitionPoste.NumPoste)
                If NumCuve > 0 Then
                    If TEtatsCuves(NumCuve).UnDefautAuMoinsSignale = True Then
                        AfficheAnimations N_LIBELLES, EffacerOuActualiser, .DefinitionPoste.NumPoste, TYPES_AFFICHAGES_LIBELLES.E_DEFAUT
                    End If
                End If
            End If
        End If

    End With
            
End Sub

'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Rôle      : Initialisation de DirectX
' Détails  :
' Retours :
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Private Sub InitialisationDirectX()
        
    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next
    
    '--- déclaration ---
    
    '--- création de l'objet direct draw ---
    Set ObjDD = ObjDX.DirectDrawCreate("")
    
    '--- niveau de coopération avec l'écran ---
    Call ObjDD.SetCooperativeLevel(Me.hwnd, DDSCL_NORMAL)
    
    '--- description de la surface de l'écran ---
    With DDSDEcran
        .lFlags = DDSD_CAPS
        .ddsCaps.lCaps = DDSCAPS_PRIMARYSURFACE
    End With
    
    '--- création de la surface ---
    Set ObjDDSEcran = ObjDD.CreateSurface(DDSDEcran)
    
    '--- création de l'objet clipper pour utiliser que certaines régions de l'écran ---
    Set ObjDDClip = ObjDD.CreateClipper(0)
    
    '--- association de l'image à l'objet clipper ---
    ObjDDClip.SetHWnd PBImageLigne.hwnd
    
    '--- attachement du clipping à l'écran ---
    ObjDDSEcran.SetClipper ObjDDClip
    
    '--- description de l'image tampon (surface invisible dans la mémoire système) ---
    With DDSDImageTampon
        .lFlags = DDSD_CAPS Or DDSD_HEIGHT Or DDSD_WIDTH   'Indicate that we want to specify the ddscaps height and width The format of the surface (bits per pixel) will be the same as the primary
        .ddsCaps.lCaps = DDSCAPS_OFFSCREENPLAIN Or DDSCAPS_SYSTEMMEMORY 'Indicate that we want a surface that is not visible and that we want it in system memory wich is plentiful as opposed to video memory
        .lWidth = PBImageLigne.Width        'Specify the height and width of the surface to be the same as the picture box (note unit are in pixels)
        .lHeight = PBImageLigne.Height
    End With
    
    '--- création de l'image tampon (surface invisible dans la mémoire système) ---
    Set ObjDDSImageTampon = ObjDD.CreateSurface(DDSDImageTampon)
   
    '--- coordonnées du rectangle de l'image tampon ---
    With RImageTampon
        .Left = 0
        .Top = 0
        .Right = DDSDImageTampon.lWidth
        .Bottom = DDSDImageTampon.lHeight
    End With
    
End Sub

'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Rôle      : Gére l'image tampon
' Détails  : ModeChoisi -> FALSE = Reconstruit l'image tampon dans la mémoire (il n'y a pas d'affichage)
'                                         TRUE  = Affichage de l'image tampon à l'écran
' Retours :
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Sub GestionImageTampon(ByVal ModeChoisi As Boolean)
    
    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next
    
    '--- déclaration ---
    Dim RDestination As RECT  'coordonnées du rectangle de destination
    
    If ModeChoisi = False Then
    
        '--- reconstruction ---
        Call ObjDDSImageTampon.BltFast(0, 0, ObjDDSImageLigne, RImageLigne, DDBLTFAST_WAIT)
    
    Else
    
        '--- récupération des coordonnées écran de l'image de la ligne ---
        Call ObjDX.GetWindowRect(PBImageLigne.hwnd, RDestination)
    
        '--- transfert de l'image tampon à l'écran ---
        Call ObjDDSEcran.Blt(RDestination, ObjDDSImageTampon, RImageTampon, DDBLT_WAIT)
    
    End If
    
End Sub

'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Rôle      : Initialisation des surfaces
' Détails  :
' Retours :
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Private Sub InitialisationSurfaces()
    
    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next
    
    '--- déclaration ---
    Dim a As Integer
    Dim CouleurCle As DDCOLORKEY
    Dim DDFormatEnPixels As DDPIXELFORMAT
    
    '-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    
    '--- description de l'image de la ligne ---
    With DDSDImageLigne
        .lFlags = DDSD_CAPS Or DDSD_WIDTH Or DDSD_HEIGHT
        .ddsCaps.lCaps = DDSCAPS_OFFSCREENPLAIN
        .lWidth = PBImageLigne.Width
        .lHeight = PBImageLigne.Height
    End With
    
    '--- création de la surface et chargement de l'image de la ligne ---
    Set ObjDDSImageLigne = ObjDD.CreateSurfaceFromFile(RepImagesAnodisation & "Synoptique.bmp", DDSDImageLigne)
    
    '--- coordonnées du rectangle du synoptique ---
    With RImageLigne
        .Left = 0
        .Top = 0
        .Right = DDSDImageLigne.lWidth
        .Bottom = DDSDImageLigne.lHeight
    End With
                                                                        
    '--- reconstruction de l'image tampon en mémoire ---
    GestionImageTampon False
        
    '-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    
    '***********
    '* PONTS *
    '***********
    
    '--- description du pont concerné ---
    With TDDSDEnsemblePonts
        .lFlags = DDSD_CAPS Or DDSD_WIDTH Or DDSD_HEIGHT
        .ddsCaps.lCaps = DDSCAPS_OFFSCREENPLAIN
        .lWidth = DIMENSIONS_ANIMATIONS.D_LONG_ENSEMBLE_PONTS
        .lHeight = DIMENSIONS_ANIMATIONS.D_HAUT_PONT
    End With

    '--- création de la surface et chargement de l'image du pont concerné ---
    Set TObjDDSEnsemblePonts = ObjDD.CreateSurfaceFromFile(RepImagesAnodisation & "Ensemble des ponts.bmp", TDDSDEnsemblePonts)
        
    '-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    
    '******************
    '* PALONNIERS *
    '******************

    '--- description du palonnier concerné ---
    With TDDSDEnsemblePalonniers
        .lFlags = DDSD_CAPS Or DDSD_WIDTH Or DDSD_HEIGHT
        .ddsCaps.lCaps = DDSCAPS_OFFSCREENPLAIN
        .lWidth = DIMENSIONS_ANIMATIONS.D_LONG_ENSEMBLE_PALONNIERS
        .lHeight = DIMENSIONS_ANIMATIONS.D_HAUT_PALONNIER
    End With
    
    '--- création de la surface et chargement de l'image du palonnier concerné ---
    Set TObjDDSEnsemblePalonniers = ObjDD.CreateSurfaceFromFile(RepImagesAnodisation & "Ensemble des palonniers.bmp", TDDSDEnsemblePalonniers)
    
    '-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    
    '*****************
    '* ACCROCHES *
    '*****************

    '--- description de l'accroche concernée ---
    With TDDSDEnsembleAccroches
        .lFlags = DDSD_CAPS Or DDSD_WIDTH Or DDSD_HEIGHT
        .ddsCaps.lCaps = DDSCAPS_OFFSCREENPLAIN
        .lWidth = DIMENSIONS_ANIMATIONS.D_LONG_ENSEMBLE_ACCROCHES
        .lHeight = DIMENSIONS_ANIMATIONS.D_HAUT_ACCROCHE
    End With

    '--- création de la surface et chargement de l'image des accroches concerné ---
    Set TObjDDSEnsembleAccroches = ObjDD.CreateSurfaceFromFile(RepImagesAnodisation & "Ensemble des accroches.bmp", TDDSDEnsembleAccroches)
        
    '-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    
    '**************
    '* CHARGES *
    '**************
    
    '--- description de l'ensemble des charges ---
    With TDDSDEnsembleCharges
        .lFlags = DDSD_CAPS Or DDSD_WIDTH Or DDSD_HEIGHT
        .ddsCaps.lCaps = DDSCAPS_OFFSCREENPLAIN
        .lWidth = DIMENSIONS_ANIMATIONS.D_LONG_ENSEMBLE_CHARGES
        .lHeight = DIMENSIONS_ANIMATIONS.D_HAUT_ENSEMBLE_CHARGES
    End With
    
    '--- création de la surface et chargement de l'image de l'ensemble des charges ---
    Set TObjDDSEnsembleCharges = ObjDD.CreateSurfaceFromFile(RepImagesAnodisation & "Ensemble des charges.bmp", TDDSDEnsembleCharges)
    
    '-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    
    '*******************
    '* COUVERCLES *
    '*******************
    
    '--- description de l'ensemble des couvercles ---
    With TDDSDEnsembleCouvercles
        .lFlags = DDSD_CAPS Or DDSD_WIDTH Or DDSD_HEIGHT
        .ddsCaps.lCaps = DDSCAPS_OFFSCREENPLAIN
        .lWidth = DIMENSIONS_ANIMATIONS.D_LONG_ENSEMBLE_COUVERCLES
        .lHeight = DIMENSIONS_ANIMATIONS.D_HAUT_ENSEMBLE_COUVERCLES
    End With
    
    '--- création de la surface et chargement de l'image des couvercles ---
    Set TObjDDSEnsembleCouvercles = ObjDD.CreateSurfaceFromFile(RepImagesAnodisation & "Ensemble des couvercles.bmp", TDDSDEnsembleCouvercles)
    
    '-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    
    '***************
    '* CHARIOTS *
    '***************
    
    '--- description d'un chariot ---
    With TDDSDChariot
        .lFlags = DDSD_CAPS Or DDSD_WIDTH Or DDSD_HEIGHT
        .ddsCaps.lCaps = DDSCAPS_OFFSCREENPLAIN
        .lWidth = DIMENSIONS_ANIMATIONS.D_LONG_CHARIOT
        .lHeight = DIMENSIONS_ANIMATIONS.D_HAUT_CHARIOT
    End With
    
    '--- création de la surface et chargement de l'image du chariot ---
    Set TObjDDSChariot(ETATS_CHARIOTS.E_PRESENT) = ObjDD.CreateSurfaceFromFile(RepImagesAnodisation & "Chariot présent.bmp", TDDSDChariot)
    Set TObjDDSChariot(ETATS_CHARIOTS.E_PRESENT_VERROUILLE) = ObjDD.CreateSurfaceFromFile(RepImagesAnodisation & "Chariot présent verrouillé.bmp", TDDSDChariot)
    
    '-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    
    '**************
    '* LIBELLES *
    '**************
    
    '--- description de l'ensemble des libellés ---
    With TDDSDEnsembleLibelles
        .lFlags = DDSD_CAPS Or DDSD_WIDTH Or DDSD_HEIGHT
        .ddsCaps.lCaps = DDSCAPS_OFFSCREENPLAIN
        .lWidth = DIMENSIONS_ANIMATIONS.D_LONG_ENSEMBLE_LIBELLES
        .lHeight = DIMENSIONS_ANIMATIONS.D_HAUT_ENSEMBLE_LIBELLES
    End With

    '--- création de la surface et chargement de l'image des libellés ---
    Set TObjDDSEnsembleLibelles = ObjDD.CreateSurfaceFromFile(RepImagesAnodisation & "Ensemble des libellés.bmp", TDDSDEnsembleLibelles)

    '-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    
    '************************************************************
    '* COULEUR TRANSPARENTE POUR LES ANIMATIONS *
    '************************************************************
    
    '--- construction de la couleur transparente pour les animations des ponts ---
    TObjDDSEnsemblePonts.GetPixelFormat DDFormatEnPixels
    With CouleurCle
        .low = DDFormatEnPixels.lRBitMask + DDFormatEnPixels.lBBitMask
        .high = DDFormatEnPixels.lRBitMask + DDFormatEnPixels.lBBitMask
    End With
    
    '--- assignation de la couleur transparente aux animations des ponts ---
    TObjDDSEnsemblePonts.SetColorKey DDCKEY_SRCBLT, CouleurCle
    
    '-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    
    '--- construction de la couleur transparente pour les animations des palonniers ---
    TObjDDSEnsemblePalonniers.GetPixelFormat DDFormatEnPixels
    With CouleurCle
        .low = DDFormatEnPixels.lRBitMask + DDFormatEnPixels.lBBitMask
        .high = DDFormatEnPixels.lRBitMask + DDFormatEnPixels.lBBitMask
    End With
    
    '--- assignation de la couleur transparente aux animations des palonniers ---
    TObjDDSEnsemblePalonniers.SetColorKey DDCKEY_SRCBLT, CouleurCle
    
    '-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    
    '--- construction de la couleur transparente pour les animations des accroches ---
    TObjDDSEnsembleAccroches.GetPixelFormat DDFormatEnPixels
    With CouleurCle
        .low = DDFormatEnPixels.lRBitMask + DDFormatEnPixels.lBBitMask
        .high = DDFormatEnPixels.lRBitMask + DDFormatEnPixels.lBBitMask
    End With
    
    '--- assignation de la couleur transparente aux animations des accroches ---
    TObjDDSEnsembleAccroches.SetColorKey DDCKEY_SRCBLT, CouleurCle
    
    '-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    
    '--- construction de la couleur transparente pour l'ensemble des charges ---
    TObjDDSEnsembleCharges.GetPixelFormat DDFormatEnPixels
    With CouleurCle
        .low = DDFormatEnPixels.lRBitMask + DDFormatEnPixels.lBBitMask
        .high = DDFormatEnPixels.lRBitMask + DDFormatEnPixels.lBBitMask
    End With
    
    '--- assignation de la couleur transparente pour l'ensemble des charges ---
    TObjDDSEnsembleCharges.SetColorKey DDCKEY_SRCBLT, CouleurCle
    
    '-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    
    '--- construction de la couleur transparente pour l'ensemble des couvercles ---
    TObjDDSEnsembleCouvercles.GetPixelFormat DDFormatEnPixels
    With CouleurCle
        .low = DDFormatEnPixels.lRBitMask + DDFormatEnPixels.lBBitMask
        .high = DDFormatEnPixels.lRBitMask + DDFormatEnPixels.lBBitMask
    End With
    
    '--- assignation de la couleur transparente pour l'ensemble des couvercles ---
    TObjDDSEnsembleCouvercles.SetColorKey DDCKEY_SRCBLT, CouleurCle
    
    '-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    
    '--- construction de la couleur transparente pour les chariots ---
    TObjDDSChariot(ETATS_CHARIOTS.E_PRESENT).GetPixelFormat DDFormatEnPixels
    With CouleurCle
        .low = DDFormatEnPixels.lRBitMask + DDFormatEnPixels.lBBitMask
        .high = DDFormatEnPixels.lRBitMask + DDFormatEnPixels.lBBitMask
    End With
    
    '--- assignation de la couleur transparente pour les chariots ---
    For a = ETATS_CHARIOTS.E_PRESENT To ETATS_CHARIOTS.E_PRESENT_VERROUILLE
        TObjDDSChariot(a).SetColorKey DDCKEY_SRCBLT, CouleurCle
    Next
    
    '-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    
    '--- construction de la couleur transparente pour l'ensemble des libellés ---
    TObjDDSEnsembleLibelles.GetPixelFormat DDFormatEnPixels
    With CouleurCle
        .low = DDFormatEnPixels.lRBitMask + DDFormatEnPixels.lBBitMask
        .high = DDFormatEnPixels.lRBitMask + DDFormatEnPixels.lBBitMask
    End With
    
    '--- assignation de la couleur transparente pour l'ensemble des libellés ---
    TObjDDSEnsembleLibelles.SetColorKey DDCKEY_SRCBLT, CouleurCle
    
    '-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    
    '--- reconstruction de l'image tampon en mémoire ---
    GestionImageTampon False
        
End Sub

'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Rôle      : Teste le niveau de coopération entre DirectDraw7 et windows
' Détails  :
' Entrées :
' Retours : TesteNiveauCooperation -> FALSE = une erreur s'est produite
'                                                               TRUE =  test de coopération correcte
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Private Function TesteNiveauCooperation() As Boolean
    
    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next
    
    '--- déclaration ---
    Dim ResultatTestCooperation As Long
    
    '--- lancement du test ---
    ResultatTestCooperation = ObjDD.TestCooperativeLevel
    
    '--- valeur de retour ---
    If (ResultatTestCooperation = DD_OK) Then
        TesteNiveauCooperation = True
    Else
        TesteNiveauCooperation = False
    End If
    
End Function

'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Rôle      : Fixe les premières positions des animations (sprites) dans le synoptique
' Entrées :
' Retours :
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Private Sub PremieresPositionsAnimations()
    
    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next

    '--- décaration ---
    Dim a As Integer
    Dim RTemp As RECT                   'coordonnées du rectangle temporaire

    For a = PONTS.P_1 To PONTS.P_2
    
        '-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
        
        '--- affectation pour les ponts ---
        Select Case a
            Case PONTS.P_1: TXPonts(a) = TEtatsPostes(POSTES.P_CHGT_1).DefinitionPoste.XAxePosteSynoptique - DIMENSIONS_ANIMATIONS.D_AXE_PONT
            Case PONTS.P_2: TXPonts(a) = TEtatsPostes(DERNIER_POSTE).DefinitionPoste.XAxePosteSynoptique - DIMENSIONS_ANIMATIONS.D_AXE_PONT
            Case Else
       End Select
       TYPonts(a) = POSITIONS_ORIGINES.PO_Y_PONT                                                     'cette valeur ne bougent pas dans le programme
       TDerniersXPonts(a) = TXPonts(a)
       TDerniersYPonts(a) = TYPonts(a)
        
        '-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
        
        '--- affectation pour les ponts fictifs ---
        Select Case a
            Case PONTS.P_1: TXPontsFictifs(a) = TEtatsPostes(POSTES.P_CHGT_1 + 1).DefinitionPoste.XAxePosteSynoptique - DIMENSIONS_ANIMATIONS.D_AXE_PONT
            Case PONTS.P_2: TXPontsFictifs(a) = TEtatsPostes(DERNIER_POSTE - 1).DefinitionPoste.XAxePosteSynoptique - DIMENSIONS_ANIMATIONS.D_AXE_PONT
            Case Else
        End Select
        TYPontsFictifs(a) = POSITIONS_ORIGINES.PO_Y_PONT_FICTIF                                 'cette valeur ne bougent pas dans le programme
        TDerniersXPontsFictifs(a) = TXPontsFictifs(a)
        TDerniersYPontsFictifs(a) = TYPontsFictifs(a)
        
        '-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
        
        '--- affectation pour les derniers X et Y ---
        TYPonts(a) = POSITIONS_ORIGINES.PO_Y_PONT
        TDerniersXPonts(a) = TXPonts(a)
        TDerniersYPonts(a) = TYPonts(a)
    
        '-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
        
        '--- palonniers 1 et 2 ---
        TXPalonniers(a) = TXPonts(a) + POSITIONS_ORIGINES.AJOUT_X_PALONNIER
        TYPalonniers(a) = TYPonts(a)
        TDerniersXPalonniers(a) = TXPalonniers(a)
        TDerniersYPalonniers(a) = TYPalonniers(a)
        
        '-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
        
        '--- accroches 1 et 2 ---
        TXAccroches(a) = TXPonts(a) + AJOUT_X_ACCROCHE
        TYAccroches(a) = TYPonts(a) + DIMENSIONS_ANIMATIONS.D_HAUT_PALONNIER
        TDerniersXAccroches(a) = TXAccroches(a)
        TDerniersYAccroches(a) = TYAccroches(a)
        
    Next a
    
End Sub

'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Rôle      : Affiche les animations (sprites) dans la mémoire tampon
' Entrées :        NomAnimations -> Nom de l'animation selon l'énumération NOMS_ANIMATIONS
'                 EffacerOuActualiser -> FALSE=Affacer, TRUE=Actualiser
'                        IndexAnimation -> Index de l'animation si il y a lieu
'                               NumImage -> Numéro de l'image pour les zones multiples d'images
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Private Sub AfficheAnimations(ByVal NomAnimations As NOMS_ANIMATIONS, _
                                                 ByVal EffacerOuActualiser As Boolean, _
                                                 Optional ByVal IndexAnimation As Variant, _
                                                 Optional ByVal NumImage As Variant)

    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next

    '--- constantes privées ---
    Const AJOUT_POUR_NUM_BARRE As Integer = 60              'ajout pour afficher le numéro de barre
    Const AJOUT_POUR_ANODISATION As Integer = 120           'ajout pour afficher les charges ne passant que par l'anodisation
    Const AJOUT_POUR_SPECTRO As Integer = 180                  'ajout pour afficher les charges passant dans la spectrocoloration
    Const AJOUT_POUR_OR As Integer = 240                             'ajout pour afficher les charges passant dans le bain d'or
    Const AJOUT_POUR_NOIR As Integer = 300                          'ajout pour afficher les charges passant dans le bain de noir
    Const AJOUT_POUR_RESTE As Integer = 360                       'ajout pour afficher le reste des charges
    Const AJOUT_POUR_SPECTRO_OR As Integer = 420           'ajout pour afficher les charges passant dans la spectrocoloration et l'or
    Const AJOUT_POUR_SPECTRO_NOIR As Integer = 480        'ajout pour afficher les charges passant dans la spectrocoloration et le noir
    
    '--- décaration ---
    Dim NumPoste As Integer
    Dim NumBarre As Integer                                                       'représente un numéro de barre
    
    Dim NumLigne As Long, _
            NumColonne As Long
    
    Dim RTemp As RECT                                                              'coordonnées du rectangle temporaire
    Dim REffacement As RECT                                                     'coordonnées du rectangle d'effacement d'une animation

    '-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    
    '--- réaffectation de l'index de l'animation pour les numéros de barres ---
    If (NomAnimations = NOMS_ANIMATIONS.N_CHARGE_PONT Or NomAnimations = NOMS_ANIMATIONS.N_CHARGE_POSTE) And _
       IsMissing(NumImage) = False And _
       ModeAffichageSynoptique = MA_NUM_BARRES Then
        
        '--- ATTENTION le n° de l'image correspond au n° de charge ---
        If NumImage >= CHARGES.C_NUM_MINI And NumImage <= CHARGES.C_NUM_MAXI Then
        
            '--- affectation du numéro de barre ---
            NumBarre = TEtatsCharges(NumImage).NumBarre
    
            If NumBarre >= BARRES.B_NUM_MINI And NumBarre <= BARRES.B_NUM_MAXI Then
    
                '--- réaffectation de l'index de l'animation ---
                NumImage = NumBarre + AJOUT_POUR_NUM_BARRE
    
            End If
        
        End If
    
    End If
    
    '-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    
    '--- réaffectation de l'index de l'animation pour les colorations ---
    If (NomAnimations = NOMS_ANIMATIONS.N_CHARGE_PONT Or NomAnimations = NOMS_ANIMATIONS.N_CHARGE_POSTE) And _
       IsMissing(NumImage) = False And _
       ModeAffichageSynoptique = MA_COLORATIONS Then
        
        '--- ATTENTION le n° de l'image correspond au n° de charge ---
        If NumImage >= CHARGES.C_NUM_MINI And NumImage <= CHARGES.C_NUM_MAXI Then
        
            '--- réaffectation de l'index de l'animation ---
            If TEtatsCharges(NumImage).TGammesAnodisation.PassageNoir = True Then
                If TEtatsCharges(NumImage).TGammesAnodisation.PassageSpectro = True Then
                    NumImage = NumImage + AJOUT_POUR_SPECTRO_NOIR
                Else
                    NumImage = NumImage + AJOUT_POUR_NOIR
                End If
            Else
                If TEtatsCharges(NumImage).TGammesAnodisation.PassageOr = True Then
                    If TEtatsCharges(NumImage).TGammesAnodisation.PassageSpectro = True Then
                        NumImage = NumImage + AJOUT_POUR_SPECTRO_OR
                    Else
                        NumImage = NumImage + AJOUT_POUR_OR
                    End If
                Else
                    If TEtatsCharges(NumImage).TGammesAnodisation.PassageSpectro = True Then
                        NumImage = NumImage + AJOUT_POUR_SPECTRO
                    Else
                        If TEtatsCharges(NumImage).TGammesAnodisation.PassageAnodisation = True Then
                            NumImage = NumImage + AJOUT_POUR_ANODISATION
                        Else
                            NumImage = NumImage + AJOUT_POUR_RESTE
                        End If
                    End If
                End If
            End If
        
        End If
    
    End If
    
    '-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    
    '--- affichage des animations ---
    Select Case NomAnimations
    
        '-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
        
        Case NOMS_ANIMATIONS.N_PONT
            '--- pont ---
            If EffacerOuActualiser = False Then
                
                '--- effacement de l'ancienne image du pont concerné ---
                With REffacement
                    .Left = TDerniersXPonts(IndexAnimation)
                    .Top = TDerniersYPonts(IndexAnimation)
                    .Right = .Left + DIMENSIONS_ANIMATIONS.D_LONG_PONT
                    .Bottom = .Top + DIMENSIONS_ANIMATIONS.D_HAUT_PONT
                End With
                Call ObjDDSImageTampon.BltFast(TDerniersXPonts(IndexAnimation), TDerniersYPonts(IndexAnimation), ObjDDSImageLigne, REffacement, DDBLTFAST_WAIT)
                            
            Else
                
                '--- transfert de l'image du pont concerné dans l'image tampon ---
                With RTemp
                    .Left = NumImage * DIMENSIONS_ANIMATIONS.D_LONG_PONT
                    .Top = 0
                    .Right = .Left + DIMENSIONS_ANIMATIONS.D_LONG_PONT
                    .Bottom = .Top + DIMENSIONS_ANIMATIONS.D_HAUT_PONT
                End With
                Call ObjDDSImageTampon.BltFast(TXPonts(IndexAnimation), TYPonts(IndexAnimation), TObjDDSEnsemblePonts, RTemp, DDBLTFAST_SRCCOLORKEY Or DDBLTFAST_WAIT)
            
            End If
        
        '-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
        
        Case NOMS_ANIMATIONS.N_PONT_FICTIF
            '--- pont fictif ---
            If EffacerOuActualiser = False Then
                
                '--- effacement de l'ancienne image du pont concerné ---
                With REffacement
                    .Left = TDerniersXPontsFictifs(IndexAnimation)
                    .Top = TDerniersYPontsFictifs(IndexAnimation)
                    .Right = .Left + DIMENSIONS_ANIMATIONS.D_LONG_PONT
                    .Bottom = .Top + DIMENSIONS_ANIMATIONS.D_HAUT_PONT
                End With
                Call ObjDDSImageTampon.BltFast(TDerniersXPontsFictifs(IndexAnimation), TDerniersYPontsFictifs(IndexAnimation), ObjDDSImageLigne, REffacement, DDBLTFAST_WAIT)
                            
            Else
                
                '--- transfert de l'image du pont concerné dans l'image tampon ---
                With RTemp
                    .Left = NumImage * DIMENSIONS_ANIMATIONS.D_LONG_PONT
                    .Top = 0
                    .Right = .Left + DIMENSIONS_ANIMATIONS.D_LONG_PONT
                    .Bottom = .Top + DIMENSIONS_ANIMATIONS.D_HAUT_PONT
                End With
                Call ObjDDSImageTampon.BltFast(TXPontsFictifs(IndexAnimation), TYPontsFictifs(IndexAnimation), TObjDDSEnsemblePonts, RTemp, DDBLTFAST_SRCCOLORKEY Or DDBLTFAST_WAIT)
            
            End If
        
        '-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
        
        Case NOMS_ANIMATIONS.N_PALONNIER
            '--- palonnier ---
            If EffacerOuActualiser = False Then
            
                '--- effacement de l'ancienne image du palonnier concerné ---
                With REffacement
                    .Left = TDerniersXPalonniers(IndexAnimation)
                    .Top = TDerniersYPalonniers(IndexAnimation)
                    .Right = .Left + DIMENSIONS_ANIMATIONS.D_LONG_PALONNIER
                    .Bottom = .Top + DIMENSIONS_ANIMATIONS.D_HAUT_PALONNIER
                End With
                Call ObjDDSImageTampon.BltFast(TDerniersXPalonniers(IndexAnimation), TDerniersYPalonniers(IndexAnimation), ObjDDSImageLigne, REffacement, DDBLTFAST_WAIT)
                
            Else
                
                '--- transfert de l'image du palonnier concerné dans l'image tampon ---
                With RTemp
                    .Left = NumImage * DIMENSIONS_ANIMATIONS.D_LONG_PALONNIER
                    .Top = 0
                    .Right = .Left + DIMENSIONS_ANIMATIONS.D_LONG_PALONNIER
                    .Bottom = .Top + DIMENSIONS_ANIMATIONS.D_HAUT_PALONNIER
                End With
                Call ObjDDSImageTampon.BltFast(TXPalonniers(IndexAnimation), TYPalonniers(IndexAnimation), TObjDDSEnsemblePalonniers, RTemp, DDBLTFAST_SRCCOLORKEY Or DDBLTFAST_WAIT)
            
            End If
        
        '-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
        
        Case NOMS_ANIMATIONS.N_ACCROCHE
            '--- accroche ---
            If EffacerOuActualiser = False Then
            
                '--- effacement de l'ancienne image du accroche concerné ---
                With REffacement
                    .Left = TDerniersXAccroches(IndexAnimation)
                    .Top = TDerniersYAccroches(IndexAnimation)
                    .Right = .Left + DIMENSIONS_ANIMATIONS.D_LONG_ACCROCHE
                    .Bottom = .Top + DIMENSIONS_ANIMATIONS.D_HAUT_ACCROCHE
                End With
                Call ObjDDSImageTampon.BltFast(TDerniersXAccroches(IndexAnimation), TDerniersYAccroches(IndexAnimation), ObjDDSImageLigne, REffacement, DDBLTFAST_WAIT)
                
            Else
        
                '--- transfert de l'image du accroche concerné dans l'image tampon ---
                With RTemp
                    .Left = NumImage * DIMENSIONS_ANIMATIONS.D_LONG_ACCROCHE
                    .Top = 0
                    .Right = .Left + DIMENSIONS_ANIMATIONS.D_LONG_ACCROCHE
                    .Bottom = .Top + DIMENSIONS_ANIMATIONS.D_HAUT_ACCROCHE
                End With
                Call ObjDDSImageTampon.BltFast(TXAccroches(IndexAnimation), TYAccroches(IndexAnimation), TObjDDSEnsembleAccroches, RTemp, DDBLTFAST_SRCCOLORKEY Or DDBLTFAST_WAIT)
        
            End If
        
        '-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
        
        Case NOMS_ANIMATIONS.N_CHARGE_PONT
            '--- charge sur le pont ---
            If EffacerOuActualiser = False Then
                
                '--- effacement de l'ancienne image d'une charge ---
                With REffacement
                    .Left = TDerniersXChargesPonts(IndexAnimation)
                    .Top = TDerniersYChargesPonts(IndexAnimation)
                    .Right = .Left + DIMENSIONS_ANIMATIONS.D_LONG_CHARGE
                    .Bottom = .Top + DIMENSIONS_ANIMATIONS.D_HAUT_CHARGE
                End With
                Call ObjDDSImageTampon.BltFast(TDerniersXChargesPonts(IndexAnimation), TDerniersYChargesPonts(IndexAnimation), ObjDDSImageLigne, REffacement, DDBLTFAST_WAIT)
            
            Else
                
                If NumImage > 0 Then
                
                    '--- calcul du numéro de ligne et colonne ---
                    NumLigne = Int(NumImage / DIMENSIONS_ANIMATIONS.D_NBR_COLONNES_ENSEMBLE_CHARGES)
                    NumColonne = NumImage Mod DIMENSIONS_ANIMATIONS.D_NBR_COLONNES_ENSEMBLE_CHARGES
                
                    '--- coordonnées de base ---
                    RTemp.Left = NumColonne * DIMENSIONS_ANIMATIONS.D_LONG_CHARGE
                    RTemp.Top = NumLigne * DIMENSIONS_ANIMATIONS.D_HAUT_CHARGE
                
                    '--- complément des coordonnées ---
                    RTemp.Right = RTemp.Left + DIMENSIONS_ANIMATIONS.D_LONG_CHARGE
                    RTemp.Bottom = RTemp.Top + DIMENSIONS_ANIMATIONS.D_HAUT_CHARGE
                    Call ObjDDSImageTampon.BltFast(TXChargesPonts(IndexAnimation), TYChargesPonts(IndexAnimation), TObjDDSEnsembleCharges, RTemp, DDBLTFAST_SRCCOLORKEY Or DDBLTFAST_WAIT)
            
                End If
            
            End If
        
        '-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
        
        Case NOMS_ANIMATIONS.N_CHARGE_POSTE
            '--- charge dans un poste ---
            If EffacerOuActualiser = False Then
                
                '--- effacement de l'ancienne image d'une charge ---
                With REffacement
                    .Left = TDernierXChargePoste
                    .Top = TDernierYChargePoste
                    .Right = .Left + DIMENSIONS_ANIMATIONS.D_LONG_CHARGE
                    .Bottom = .Top + DIMENSIONS_ANIMATIONS.D_HAUT_CHARGE
                End With
                Call ObjDDSImageTampon.BltFast(TDernierXChargePoste, TDernierYChargePoste, ObjDDSImageLigne, REffacement, DDBLTFAST_WAIT)
            
            Else
                
                If NumImage > 0 Then
                
                    '--- calcul du numéro de ligne et colonne ---
                    NumLigne = Int(NumImage / DIMENSIONS_ANIMATIONS.D_NBR_COLONNES_ENSEMBLE_CHARGES)
                    NumColonne = NumImage Mod DIMENSIONS_ANIMATIONS.D_NBR_COLONNES_ENSEMBLE_CHARGES
                
                    '--- coordonnées de base ---
                    RTemp.Left = NumColonne * DIMENSIONS_ANIMATIONS.D_LONG_CHARGE
                    RTemp.Top = NumLigne * DIMENSIONS_ANIMATIONS.D_HAUT_CHARGE
                
                    '--- complément des coordonnées ---
                    RTemp.Right = RTemp.Left + DIMENSIONS_ANIMATIONS.D_LONG_CHARGE
                    RTemp.Bottom = RTemp.Top + DIMENSIONS_ANIMATIONS.D_HAUT_CHARGE
                    Call ObjDDSImageTampon.BltFast(TXChargePoste, TYChargePoste, TObjDDSEnsembleCharges, RTemp, DDBLTFAST_SRCCOLORKEY Or DDBLTFAST_WAIT)
            
                End If
            
            End If
        
        '-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
        
        'Case NOMS_ANIMATIONS.N_COUVERCLES
            '--- couvercles ---
            'If EffacerOuActualiser = False Then
                
                '--- effacement de l'ancienne image des couvercles ---
                'With REffacement
                '    .Left = TDernierXCouvercles
                '    .Top = TDernierYCouvercles
                '    .Right = .Left + DIMENSIONS_ANIMATIONS.D_LONG_COUVERCLES
                '    .Bottom = .Top + DIMENSIONS_ANIMATIONS.D_HAUT_COUVERCLES
                'End With
                'Call ObjDDSImageTampon.BltFast(TDernierXCouvercles, TDernierYCouvercles, ObjDDSImageLigne, REffacement, DDBLTFAST_WAIT)
            
            'Else
                
                '--- transfert de l'image des couvercles dans l'image tampon ---
                'With RTemp
                '    .Left = 0
                '    .Top = DIMENSIONS_ANIMATIONS.D_HAUT_COUVERCLES * NumImage
                '    .Right = DIMENSIONS_ANIMATIONS.D_LONG_COUVERCLES
                '    .Bottom = .Top + DIMENSIONS_ANIMATIONS.D_HAUT_COUVERCLES
                'End With
                'Call ObjDDSImageTampon.BltFast(TXCouvercles, TYCouvercles, TObjDDSEnsembleCouvercles, RTemp, DDBLTFAST_SRCCOLORKEY Or DDBLTFAST_WAIT)
                
            'End If
        
        '-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
        
        Case NOMS_ANIMATIONS.N_CHARIOTS
            '--- chariots ---
            If EffacerOuActualiser = False Then
                
                '--- effacement de l'ancienne image du chariot ---
                With REffacement
                    .Left = TDernierXChariot
                    .Top = TDernierYChariot
                    .Right = .Left + DIMENSIONS_ANIMATIONS.D_LONG_CHARIOT
                    .Bottom = .Top + DIMENSIONS_ANIMATIONS.D_HAUT_CHARIOT
                End With
                Call ObjDDSImageTampon.BltFast(TDernierXChariot, TDernierYChariot, ObjDDSImageLigne, REffacement, DDBLTFAST_WAIT)
            
            Else
                
                '--- transfert de l'image du chariot ---
                With RTemp
                    .Left = 0
                    .Top = 0
                    .Right = DIMENSIONS_ANIMATIONS.D_LONG_CHARIOT
                    .Bottom = DIMENSIONS_ANIMATIONS.D_HAUT_CHARIOT
                End With
                Call ObjDDSImageTampon.BltFast(TXChariot, TYChariot, TObjDDSChariot(IndexAnimation), RTemp, DDBLTFAST_SRCCOLORKEY Or DDBLTFAST_WAIT)
                
            End If
        
        '-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
        
        Case NOMS_ANIMATIONS.N_LIBELLES
            '--- libellés ---
            ' dans ce cas IndexAnimation = le numéro de poste ou ce trouve le libellé
            '                             NumImage = 1 libellé en cas de défaut
            '                                                  2 libellé en cas de condamnation
            NumPoste = IndexAnimation
            
            If EffacerOuActualiser = False Then
                
                '--- effacement de l'ancienne image d'un libellé ---
                With REffacement
                    .Left = TDernierXLibelle
                    .Top = TDernierYLibelle
                    Select Case NumPoste
                        Case POSTES.P_CHGT_1 To POSTES.P_CHGT_2
                            .Right = .Left + DIMENSIONS_ANIMATIONS.D_LONG_1_LIBELLE
                        'Case PREMIER_BAIN To POSTES.P_C35
                        Case POSTES.P_C02 To POSTES.P_C35
                            .Right = .Left + DIMENSIONS_ANIMATIONS.D_LONG_2_LIBELLE
                        Case POSTES.P_D1 To POSTES.P_D2
                            .Right = .Left + DIMENSIONS_ANIMATIONS.D_LONG_3_LIBELLE
                        Case POSTES.P_C37 To POSTES.P_C38
                            .Right = .Left + DIMENSIONS_ANIMATIONS.D_LONG_2_LIBELLE
                        Case Else
                    End Select
                    .Bottom = .Top + DIMENSIONS_ANIMATIONS.D_HAUT_LIBELLE
                End With
                Call ObjDDSImageTampon.BltFast(TDernierXLibelle, TDernierYLibelle, ObjDDSImageLigne, REffacement, DDBLTFAST_WAIT)
            
            Else
            
                If NumPoste >= POSTES.P_CHGT_1 And NumPoste <= DERNIER_POSTE And _
                   NumImage <= DIMENSIONS_ANIMATIONS.D_NBR_COLONNES_ENSEMBLE_LIBELLES Then

                    '--- coordonnées de base ---
                    RTemp.Left = NumImage * DIMENSIONS_ANIMATIONS.D_LONG_1_LIBELLE
                    RTemp.Top = Pred(NumPoste) * DIMENSIONS_ANIMATIONS.D_HAUT_LIBELLE
                    RTemp.Right = RTemp.Left + DIMENSIONS_ANIMATIONS.D_LONG_1_LIBELLE
                    RTemp.Bottom = RTemp.Top + DIMENSIONS_ANIMATIONS.D_HAUT_LIBELLE
                    Call ObjDDSImageTampon.BltFast(TXLibelle, TYLibelle, TObjDDSEnsembleLibelles, RTemp, DDBLTFAST_SRCCOLORKEY Or DDBLTFAST_WAIT)

                End If
            
            End If
        
        Case Else

    End Select

End Sub

'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Rôle      : Change la forme de la fenêtre (positions des divers éléments qui la compose)
' Entrées : FormeSouhaite -> Fonction de l'énumération FORMES_fenetre
' Retours :
' Détails  :
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Sub ChangeFormeFenetre(ByVal FormeSouhaite As FORMES_FENETRE)

    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next

    '--- constantes privées ---
    Const AGRANDIR_EN_HAUT As String = "agrandir en haut"
    Const RESTAURER_TAILLE_EN_HAUT As String = "restaurer taille en haut"

    
    '--- déclaration ---
    Dim TwipsParPixelX As Single, _
            TwipsParPixelY As Single
    
    '--- affectation ---
    TwipsParPixelX = Screen.TwipsPerPixelX
    TwipsParPixelY = Screen.TwipsPerPixelY

    '--- formes de la fenêtre ---
    Select Case FormeSouhaite

        Case FORMES_FENETRE.F_STANDARD
            '--- forme standard ---
            'synoptique en haut visible
            'états de la ligne invisible
            'dialogues en bas à gauche visible
            
            '--- effacement par défaut des élements lents à l'affichage ---
            RTBDialogues.Visible = False
            PBGeneral.Visible = True
            PBEtatsLigne.Visible = False
            
            '--- position et dimensions du synoptique ---
            Set CBAgrandirRestaurerZonesFenetre(ZONES_FENETRE.Z_SYNOPTIQUE).Picture = Me.ILOutilsDivers.ListImages(AGRANDIR_EN_HAUT).Picture
            With PBSynoptique
                .Left = 3
                .Top = 3
                .Width = Me.ScaleWidth - 6
                .Height = HAUTEUR_IMAGE_LIGNE + PBGeneral.Height - 1
                .Visible = True
            End With
            With PBImageLigne
                .Left = 29
                .Top = 0
                .Width = LONGUEUR_IMAGE_LIGNE
                .Height = HAUTEUR_IMAGE_LIGNE
            End With
            With PBGeneral
                .Left = PBImageLigne.Left
                .Top = HAUTEUR_IMAGE_LIGNE + 2
                .Width = LONGUEUR_IMAGE_LIGNE
            End With
            
            '--- position et dimensions de l'ensemble des dialogues ---
            Set CBAgrandirRestaurerZonesFenetre(ZONES_FENETRE.Z_DIALOGUES).Picture = Me.ILOutilsDivers.ListImages(AGRANDIR_EN_HAUT).Picture
            With PBDialogues
                .Left = 3
                .Top = PBGeneral.Top + PBGeneral.Height + 1
                .Width = PBSynoptique.Width - PBEtatsPrincipaux.Width
                .Height = Me.ScaleHeight - PBSynoptique.Height - 6
                .Visible = True
            End With
            
            '--- dialogues et intranet ---
            COBConteneurOutilsDialogues.Width = PBDialogues.ScaleWidth - PBBarreAgrandissement(ZONES_FENETRE.Z_DIALOGUES).Width
            With RTBDialogues
                .Width = COBConteneurOutilsDialogues.Width
                .Height = PBDialogues.ScaleHeight - COBConteneurOutilsDialogues.Height + 3
            End With
           
            '--- les états principaux ---
            With PBEtatsPrincipaux
                .Left = PBDialogues.Left + PBDialogues.Width
                .Top = PBDialogues.Top
                .Height = PBDialogues.Height
                .Visible = True
            End With
           
            '--- réaffichage des éléments invisibles ---
            RTBDialogues.Visible = True
            
            '--- focus par défaut ---
            If PremiereActivation = True Then
                Select Case ModeDialoguesEnCours
                    Case MODES_DIALOGUES.M_RENSEIGNEMENTS
                        RTBDialogues.SetFocus
                    Case MODES_DIALOGUES.M_QUESTIONS_REPONSES
                        RTBDialogues.SetFocus
                    Case Else
                        PBBarreAgrandissement(ZONES_FENETRE.Z_SYNOPTIQUE).SetFocus
                End Select
            End If
   
        Case FORMES_FENETRE.F_SYNOPTIQUE
            '--- synoptique plein écran ---
            'synoptique en haut plein écran visible
            'dialogues en bas à gauche invisible
            'états de la ligne visible
            
            '--- effacement des autres zones ---
            PBGeneral.Visible = False
            PBEtatsLigne.Visible = True
            PBDialogues.Visible = False
            PBEtatsPrincipaux.Visible = False
            
            '--- position et dimensions du synoptique ---
            Set CBAgrandirRestaurerZonesFenetre(ZONES_FENETRE.Z_SYNOPTIQUE).Picture = Me.ILOutilsDivers.ListImages(RESTAURER_TAILLE_EN_HAUT).Picture
            With PBSynoptique
                .Left = 3
                .Top = 3
                .Width = Me.ScaleWidth - 6
                .Height = Me.ScaleHeight - 6
                .Visible = True
            End With
            With PBEtatsLigne
                .Left = PBGeneral.Left
                .Top = PBGeneral.Top
                .Width = PBGeneral.Width
                .Height = PBSynoptique.ScaleHeight - PBImageLigne.Height - 2
            End With
            
            '--- focus par défaut ---
            If PremiereActivation = True Then
                PBBarreAgrandissement(ZONES_FENETRE.Z_SYNOPTIQUE).SetFocus
            End If
        
        Case FORMES_FENETRE.F_DIALOGUES
            '--- dialogues plein écran ---
            'synoptique en haut invisible
            'dialogues en haut plein écran visible

            '--- effacement des autres zones ---
            PBSynoptique.Visible = False
            PBGeneral.Visible = False
            
            '--- effacement par défaut des élements lents à l'affichage ---
            RTBDialogues.Visible = False
            
            '--- position et dimensions de l'ensemble des dialogues ---
            Set CBAgrandirRestaurerZonesFenetre(ZONES_FENETRE.Z_DIALOGUES).Picture = Me.ILOutilsDivers.ListImages(RESTAURER_TAILLE_EN_HAUT).Picture
            With PBDialogues
                .Left = 3
                .Top = 3
                .Width = Me.ScaleWidth - PBEtatsPrincipaux.Width - 6
                .Height = Me.ScaleHeight - 6
                .Visible = True
            End With
            
            '--- dialogues et intranet ---
            COBConteneurOutilsDialogues.Width = PBDialogues.ScaleWidth - PBBarreAgrandissement(ZONES_FENETRE.Z_DIALOGUES).Width
            With RTBDialogues
                .Width = COBConteneurOutilsDialogues.Width
                .Height = PBDialogues.ScaleHeight - COBConteneurOutilsDialogues.Height + 3
            End With
            
            '--- les états principaux ---
            With PBEtatsPrincipaux
                .Left = PBDialogues.Left + PBDialogues.Width
                .Top = PBDialogues.Top
                .Height = PBDialogues.Height
                .Visible = True
            End With
            
            '--- réaffichage des éléments invisibles ---
            RTBDialogues.Visible = True
            
            '--- focus par défaut ---
            If PremiereActivation = True Then
                Select Case ModeDialoguesEnCours
                    Case MODES_DIALOGUES.M_RENSEIGNEMENTS
                        RTBDialogues.SetFocus
                    Case MODES_DIALOGUES.M_QUESTIONS_REPONSES
                        RTBDialogues.SetFocus
                    Case Else
                        PBBarreAgrandissement(ZONES_FENETRE.Z_DIALOGUES).SetFocus
                End Select
            End If
            
        Case Else

    End Select
            
    '--- affectation ---
    FormeFenetre = FormeSouhaite

End Sub

Private Sub Text1_Change()

End Sub

Private Sub TimerEtatsLigne_Timer()
    
    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next
    
    '--- rafraichissement des états de la ligne ---
    TimerEtatsLigne.Enabled = False
    If ArretTachesRapides = False Then
        VisualisationEtatsLigne
        ClignotantPourSynoptique = Not (ClignotantPourSynoptique)
        TimerEtatsLigne.Enabled = True
    End If

End Sub

Private Sub TimerSynoptique_Timer()
    
    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next
    
    '--- rafraichissement du synoptique ---
    TimerSynoptique.Enabled = False
    If ArretTachesRapides = False Then
        VisualisationEtatsSynoptique
        TimerSynoptique.Enabled = True
    End If

End Sub

Private Sub TOBEffacementDialogues_ButtonClick(ByVal Button As MSComctlLib.Button)
    
    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next

    '--- sélection en fonction de l'outil cliqué ---
    Select Case Button.Key
        
        Case "effacer"
            '--- effacement des dialogues ---
            If AppelFenetre(F_MESSAGE, TITRE_MESSAGES, _
                                    vbCrLf & vbCrLf & _
                                    "c|Vous êtes sur le point d'effacer tous les dialogues." & vbCrLf & _
                                    vbCrLf & vbCrLf & vbCrLf & _
                                    "cs|Voulez-vous réellement effectuer cette opération ?", _
                                     1, 0, 1) = vbYes Then
                With RTBDialogues
                    .Text = ""
                    .Refresh
                    If .Visible = True Then .SetFocus
                End With
            End If
        
        Case Else

    End Select

End Sub

'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Rôle      : Sélectionne les modes des dialogues
' Entrées : ModeSouhaitee -> Mode souhaité fonction de l'énumération MODES_DIALOGUES
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Sub SelectionneModesDialogues(ByVal ModeSouhaite As MODES_DIALOGUES)

    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next
    
    '--- constantes privées ---
    Const NUM_BANDE_EFFACEMENT_DIALOGUES As Integer = 1
    Const NUM_BANDE_INTRANET As Integer = 3
    
    '--- déclaration ---
    Dim MemModeDialoguesEnCours As Integer
    
    '--- mémorisation de l'ancien mode des dialogues en cours ---
    MemModeDialoguesEnCours = ModeDialoguesEnCours
    
    '--- désélection de tous les modes ---
    TOBOutilsDialogues.buttons("renseignements").Image = ILOutilsDialogues2.ListImages("renseignements").Index
    TOBOutilsDialogues.buttons("questions reponses").Image = ILOutilsDialogues2.ListImages("questions reponses").Index
    TOBOutilsDialogues.buttons("previsionnel").Image = ILOutilsDialogues2.ListImages("previsionnel").Index
    TOBOutilsDialogues.buttons("entrees charges").Image = ILOutilsDialogues2.ListImages("entrees charges").Index
    
    '--- chargement du mode ---
    Select Case ModeSouhaite
        
        '-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
        
        Case MODES_DIALOGUES.M_RENSEIGNEMENTS
            '--- mise en sélection de l'outil renseignements ---
            TOBOutilsDialogues.buttons("renseignements").Image = ILOutilsDialogues2.ListImages("renseignements en selection").Index
            COBConteneurOutilsDialogues.Bands(NUM_BANDE_EFFACEMENT_DIALOGUES).Visible = True
            
            '--- affichage dans la zone des dialogues ---
            If ModeDialoguesEnCours <> MODES_DIALOGUES.M_RENSEIGNEMENTS Then
                AfficheDialogues COULEURS.ROUGE_3, "Mode renseignements sur la ligne " & DateMessages & vbCrLf
            End If
            
            '--- affectation ---
            ModeDialoguesEnCours = MODES_DIALOGUES.M_RENSEIGNEMENTS
                
            '--- focus ---
            RTBDialogues.SetFocus
        
        '-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
        
        Case MODES_DIALOGUES.M_QUESTIONS_REPONSES
            '--- mise en sélection de l'outil question réponses ---
            TOBOutilsDialogues.buttons("questions reponses").Image = ILOutilsDialogues2.ListImages("questions reponses en selection").Index
            COBConteneurOutilsDialogues.Bands(NUM_BANDE_EFFACEMENT_DIALOGUES).Visible = True
            
            '--- affichage dans la zone des dialogues ---
            If ModeDialoguesEnCours <> MODES_DIALOGUES.M_QUESTIONS_REPONSES Then
                AfficheDialogues COULEURS.ROUGE_3, "Mode questions / réponses " & DateMessages & vbCrLf
            End If
            
            '--- affectation ---
            ModeDialoguesEnCours = MODES_DIALOGUES.M_QUESTIONS_REPONSES
            
            '--- focus ---
            RTBDialogues.SetFocus
        
        '-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
        
        Case MODES_DIALOGUES.M_PREVISIONNEL
            '--- mise en sélection de l'outil prévisionnel ---
            TOBOutilsDialogues.buttons("previsionnel").Image = ILOutilsDialogues2.ListImages("previsionnel en selection").Index
            COBConteneurOutilsDialogues.Bands(NUM_BANDE_EFFACEMENT_DIALOGUES).Visible = True
            
            '--- affichage dans la zone des dialogues ---
            If ModeDialoguesEnCours <> MODES_DIALOGUES.M_PREVISIONNEL Then
                AfficheDialogues COULEURS.ROUGE_3, "Mode prévisionnel " & DateMessages & vbCrLf
            End If
            
            '--- affectation ---
            ModeDialoguesEnCours = MODES_DIALOGUES.M_PREVISIONNEL
            
            '--- focus ---
            RTBDialogues.SetFocus
        
        '-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
        
        Case MODES_DIALOGUES.M_ENTREE_CHARGES
            '--- mise en sélection de l'outil de renseignements de l'entrée des charges ---
            TOBOutilsDialogues.buttons("entrees charges").Image = ILOutilsDialogues2.ListImages("entrees charges en selection").Index
            COBConteneurOutilsDialogues.Bands(NUM_BANDE_EFFACEMENT_DIALOGUES).Visible = True
            
            '--- affichage dans la zone des dialogues ---
            If ModeDialoguesEnCours <> MODES_DIALOGUES.M_ENTREE_CHARGES Then
                AfficheDialogues COULEURS.ROUGE_3, "Mode d'informations sur les entrées des charges " & DateMessages & vbCrLf
            End If
            
            '--- affectation ---
            ModeDialoguesEnCours = MODES_DIALOGUES.M_ENTREE_CHARGES
                
            '--- focus ---
            RTBDialogues.SetFocus
        
        Case Else
    End Select
    
    '--- rafraichissement ---
    TOBOutilsDialogues.Refresh
    
End Sub

Private Sub TOBOutilsDialogues_ButtonClick(ByVal Button As MSComctlLib.Button)
    
    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next

    '--- sélection en fonction de l'outil cliqué ---
    Select Case Button.Key
        
        Case "renseignements"
            '--- mode renseignements ---
            SelectionneModesDialogues MODES_DIALOGUES.M_RENSEIGNEMENTS
        
        Case "questions reponses"
            '--- mode questions réponses ---
            SelectionneModesDialogues MODES_DIALOGUES.M_QUESTIONS_REPONSES
        
        Case "previsionnel"
            '--- prévisionnel ---
            SelectionneModesDialogues MODES_DIALOGUES.M_PREVISIONNEL
        
        Case "entrees charges"
            '--- mode entrée des charges ---
            SelectionneModesDialogues MODES_DIALOGUES.M_ENTREE_CHARGES
        
        Case Else

    End Select

End Sub

'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Rôle      : Gère l'appui des touches du clavier
' Entrées :
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Sub GestionTouches(KeyCode As Integer, Shift As Integer)
    
    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next
    
    '--- action en fonction des touches ---
    Select Case KeyCode
        
        Case vbKeyF1
            '--- touche F1 (synoptique de base) ---
            ChangeFormeFenetre F_STANDARD
            KeyCode = 0
        
        Case vbKeyF2
            '--- touche F2 (zone synoptique) ---
            If Shift = 0 Then
                AppelFenetre FENETRES.F_MODE_CYCLIQUE
            ElseIf Shift = vbShiftMask Then
                If FormeFenetre = FORMES_FENETRE.F_SYNOPTIQUE Then
                    ChangeFormeFenetre F_STANDARD
                Else
                    ChangeFormeFenetre F_SYNOPTIQUE
                End If
            End If
            KeyCode = 0
        
        Case vbKeyF3
            '--- touche F3 (zone des dialogues) ---
            If Shift = 0 Then
                AppelFenetre FENETRES.F_GAMMES_ANODISATION
            ElseIf FormeFenetre = FORMES_FENETRE.F_DIALOGUES Then
                ChangeFormeFenetre F_STANDARD
            Else
                ChangeFormeFenetre F_DIALOGUES
            End If
            KeyCode = 0
        
        Case vbKeyF4
            '--- touche F4 (traçabilité) ---
            AppelFenetre FENETRES.F_TRACABILITE_PRODUCTION
            KeyCode = 0
        
        Case vbKeyF5
            '--- touche F5 (charges en ligne) ---
            AppelFenetre FENETRES.F_CHARGES_EN_LIGNE
            KeyCode = 0
        
        Case vbKeyF6
            '--- touche F6 (cycles des ponts) ---
            AppelFenetre FENETRES.F_CYCLES_PONTS
            KeyCode = 0
        
        Case vbKeyF7
            '--- touche F7 (chargement / prévisionnel) ---
            AppelFenetre FENETRES.F_CHARGEMENT_PREVISIONNEL
            KeyCode = 0
        
        Case vbKeyF8
            '--- touche F8 (redresseurs) ---
            AppelFenetre FENETRES.F_GESTION_REDRESSEURS
            KeyCode = 0
        
        Case vbKeyF9
            '--- touche F9 (cuves) ---
            AppelFenetre FENETRES.F_GESTION_CUVES
            KeyCode = 0
        
        Case vbKeyF10
            '--- touche F10 (régulation) ---
            AppelFenetre FENETRES.F_GESTION_REGULATION
            KeyCode = 0
        
        Case vbKeyF11
            '--- touche F11 (programmateur cyclique) ---
            AppelFenetre FENETRES.F_PROGRAMMATEUR_CYCLIQUE
            KeyCode = 0
        
        Case vbKeyF12
            '--- touche F12 (annexes) ---
            AppelFenetre FENETRES.F_ANNEXES
            'AppelFenetre FENETRES.F_ESSAIS
            KeyCode = 0
       
        
        Case Else

    End Select

End Sub

'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Rôle      : Afficher les dialogues entre l'homme et la machine
' Entrées : CouleurTexteDuDialogue -> Couleur du texte du dialogue
'                             TexteDuDialogue -> Texte à afficher
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Sub AfficheDialogues(ByVal CouleurTexteDuDialogue As COULEURS, _
                                              ByVal TexteDuDialogue As String)
    
    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next
    
    With Me.RTBDialogues
        
        '--- contrôle de la longueur ---
        If Len(.Text) > 5000 Then
            .Text = ""
        End If
        
        '--- affichage du texte ---
        .SelColor = CouleurTexteDuDialogue
        .SelText = TexteDuDialogue
        .SelStart = Len(.Text)
 
        '.Refresh
    
        '--- forcer la couleur verte par défaut après affichage ---
        .SelColor = COULEURS.VERT_5
    
    End With

End Sub

Private Sub CBEntreeAutomatiqueCharges_Click()
    
    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next

    '--- affectation de la variable entrée automatique des charges ---
    EntreeAutomatiqueCharges = Not (EntreeAutomatiqueCharges)

    '--- changement de couleur et de texte du bouton ---
    With LEntreeAutomatiqueCharges

        If EntreeAutomatiqueCharges = True Then
            .Caption = "MODE AUTOMATIQUE - ENTREE OPTIMISEE DES CHARGES"
            .BackColor = COULEURS.VERT_3
            .ForeColor = COULEURS.NOIR
        Else
            .Caption = "MODE MANUEL - ENTREE IMMEDIATE DES CHARGES"
            .BackColor = COULEURS.ROUGE_3
            .ForeColor = COULEURS.BLANC
        End If
        .Refresh
    
    End With

End Sub

'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Rôle      : Gestion des en cours
' Entrées : EtatSouhaite -> Fonction de l'énumération GESTION_GRILLES
' Retours :
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Private Sub GestionEnCours(ByVal EtatSouhaite As GESTION_GRILLES)
    
    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next

    '--- constantes privées ---

    '--- déclaration ---
    Dim a As Integer                                                   'pour les boucles FOR...NEXT
    Dim MemLigne As Integer                                    'mémoire d'un numéro de ligne
    Dim MemColonne As Integer                                'mémoire d'un numéro de colonne
    Dim NumPoste As Integer                                    'représente un numéro de poste quelconque
    
    Dim CouleurLignes As Long                                 'représente une couleur de lignes
    Static CouleurLignesPaires As Long                    'représente une couleur de lignes paires
    Static CouleurLignesImpaires As Long                'représente une couleur de lignes impaires
    Static CouleurLignesNumChargePont As Long    'représente une couleur de lignes pour la charge sur le pont
    
    Dim TexteCellule As String                                  'représente le texte d'une cellule
    
    Select Case EtatSouhaite

        Case GESTION_GRILLES.GG_INITIALISATION
            '--- initialisation de la grille des détails ---
            With VSFGEnCours
                
                .Redraw = False
                
                .Clear
                
                .Editable = flexEDNone                                        'grille non éditable
                
                .BackColorFixed = COULEURS.BLEU_5              'couleur de fond des titres
                .ForeColorFixed = COULEURS.BLANC                 'couleur de premier plan des titres
                
                CouleurLignesPaires = COULEURS.JAUNE_0
                CouleurLignesImpaires = COULEURS.VERT_0
                CouleurLignesNumChargePont = COULEURS.ROUGE_1
                
                .BackColor = CouleurLignesPaires                      'couleur de fond de la grille
                .BackColorAlternate = CouleurLignesImpaires    'couleur de fond en alternance des lignes de la grille
                
                .ForeColor = COULEURS.NOIR                             'couleur de premier plan
                
                .FixedCols = 1
                .FixedRows = 1
                .Rows = NBR_LIGNES_DETAILS_EN_COURS + .FixedRows
                .Cols = NBR_COLONNES_DETAILS_EN_COURS + .FixedCols
                .RowHeight(0) = 750                                        'épaisseur des titres
                .RowHeightMin = 315                                      'épaisseur mini des lignes
                .MergeCells = flexMergeNever                       'type de mélange de cellules
                .Row = 0
            
                '--- paramétrages de chaque colonne ---
                .Col = COLONNES_DETAILS_EN_COURS.C_NUM_LIGNES
                .ColWidth(.Col) = 3 * EPAISSEUR_CARACTERE: .Text = ""
                .ColAlignment(.Col) = flexAlignRightCenter

                .Col = COLONNES_DETAILS_EN_COURS.C_NUM_COMMANDE_INTERNE
                .ColWidth(.Col) = 10 * EPAISSEUR_CARACTERE: .Text = "Numéro de pointage"
                .ColAlignment(.Col) = flexAlignCenterCenter
                
                .Col = COLONNES_DETAILS_EN_COURS.C_POSTE
                .ColWidth(.Col) = 10 * EPAISSEUR_CARACTERE: .Text = "Poste ou pont"
                .ColAlignment(.Col) = flexAlignCenterCenter
                
                .Col = COLONNES_DETAILS_EN_COURS.C_CODE_CLIENT
                .ColWidth(.Col) = 20 * EPAISSEUR_CARACTERE: .Text = "Code client"
                .ColAlignment(.Col) = flexAlignLeftCenter
                
                .Col = COLONNES_DETAILS_EN_COURS.C_NUM_BARRE
                .ColWidth(.Col) = 8 * EPAISSEUR_CARACTERE: .Text = "N° de la barre"
                .ColAlignment(.Col) = flexAlignLeftCenter

                .Col = COLONNES_DETAILS_EN_COURS.C_NBR_PIECES
                .ColWidth(.Col) = 8 * EPAISSEUR_CARACTERE: .Text = "Nombre de pièces"
                .ColAlignment(.Col) = flexAlignRightCenter

                .Col = COLONNES_DETAILS_EN_COURS.C_DESIGNATION
                .ColWidth(.Col) = 30 * EPAISSEUR_CARACTERE: .Text = "Désignation"
                .ColAlignment(.Col) = flexAlignLeftCenter
                
                .Col = COLONNES_DETAILS_EN_COURS.C_MATIERE
                .ColWidth(.Col) = 30 * EPAISSEUR_CARACTERE: .Text = "Matière"
                .ColAlignment(.Col) = flexAlignLeftCenter
                
                '--- centrage des titres ---
                .Row = 0
                For a = 1 To Pred(.Cols)
                    .Col = a
                    .CellAlignment = flexAlignCenterCenter                'tout centré
                Next a
            
                '--- N° de lignes, vidage des champs ---
                For a = 1 To Pred(.Rows)
                    .Col = COLONNES_DETAILS_EN_COURS.C_NUM_LIGNES
                    .Row = a
                    .Text = CStr(a)
                Next a
            
                '--- fixer le curseur ---
                .Row = 1
                .Col = COLONNES_DETAILS_EN_COURS.C_NUM_LIGNES
                
                .Redraw = True
            
            End With
                
        Case GESTION_GRILLES.GG_AFFICHAGE
            '--- affichage de la grille ---
            With VSFGEnCours
                
                '--- mémorisation des valeurs ligne, colonne ---
                MemLigne = .Row
                MemColonne = .Col
                .FocusRect = flexFocusNone                      'pas de focus durant l'affichage

                '--- bloquer l'affichage de la grille ---
                .Redraw = False
            
                For a = CHARGES.C_NUM_MINI To CHARGES.C_NUM_MAXI
                    
                    With TEtatsCharges(a)
                    
                        If ExistenceNumeroCharge(a) = False Then
                        
                            '***************************** EFFACEMENT SI LA CHARGE N'EST PAS EN LIGNE *****************************
                            VSFGEnCours.Cell(flexcpText, a, COLONNES_DETAILS_EN_COURS.C_NUM_COMMANDE_INTERNE, _
                                                                              a, COLONNES_DETAILS_EN_COURS.C_MATIERE) = ""
                    
                        Else
                        
                            '************************************** AFFICHAGE DE LA CHARGE *********************************************
                            If .NbrPostesTraites > 0 Or TEtatsPonts(PONTS.P_1).NumCharge = a Or TEtatsPonts(PONTS.P_2).NumCharge = a Then
                            
                                '--- affectation du numéro de poste de la charge ---
                                NumPoste = 0
                                If .NbrPostesTraites > 0 Then
                                    NumPoste = .TDetailsFichesProduction(.NbrPostesTraites).NumPoste
                                End If
                                
                                '--- N° de commande interne ---
                                TexteCellule = TEtatsCharges(a).TDetailsCharges(1).NumCommandeInterne
                                AffichageTexteMatrice VSFGEnCours, a, COLONNES_DETAILS_EN_COURS.C_NUM_COMMANDE_INTERNE, TexteCellule
                                
                                '--- libellé du poste ---
                                If TEtatsPonts(PONTS.P_1).NumCharge = a Then
                                    TexteCellule = "PONT 1"
                                ElseIf TEtatsPonts(PONTS.P_2).NumCharge = a Then
                                    TexteCellule = "PONT 2"
                                Else
                                    If NumPoste > 0 Then
                                        TexteCellule = UCase(TEtatsPostes(NumPoste).DefinitionPoste.NomPoste)
                                    Else
                                        TexteCellule = "-"
                                    End If
                                End If
                                AffichageTexteMatrice VSFGEnCours, a, COLONNES_DETAILS_EN_COURS.C_POSTE, TexteCellule
                            
                                '--- code client ---
                                TexteCellule = TEtatsCharges(a).TDetailsCharges(1).CodeClient
                                AffichageTexteMatrice VSFGEnCours, a, COLONNES_DETAILS_EN_COURS.C_CODE_CLIENT, TexteCellule
                                
                                '--- numéro de barre ---
                                TexteCellule = TBarres(TEtatsCharges(a).NumBarre).Libelle
                                
                                AffichageTexteMatrice VSFGEnCours, a, COLONNES_DETAILS_EN_COURS.C_NUM_BARRE, TexteCellule
                            
                                '--- nombre de pièces ---
                                TexteCellule = TEtatsCharges(a).TDetailsCharges(1).NbrPieces
                                AffichageTexteMatrice VSFGEnCours, a, COLONNES_DETAILS_EN_COURS.C_NBR_PIECES, TexteCellule
                                
                                '--- désignation ---
                                TexteCellule = TEtatsCharges(a).TDetailsCharges(1).Designation
                                AffichageTexteMatrice VSFGEnCours, a, COLONNES_DETAILS_EN_COURS.C_DESIGNATION, TexteCellule
                                
                                '--- matière ---
                                TexteCellule = TEtatsCharges(a).TDetailsCharges(1).Matiere
                                AffichageTexteMatrice VSFGEnCours, a, COLONNES_DETAILS_EN_COURS.C_MATIERE, TexteCellule
                                
                            End If
                    
                        End If
                    
                        '--- changement de couleur de la ligne si la charge structée est celle sur le pont ---
                        If TEtatsPonts(PONTS.P_1).NumCharge = a Or TEtatsPonts(PONTS.P_2).NumCharge = a Then
                            CouleurLignes = CouleurLignesNumChargePont
                        Else
                            If a Mod 2 = 0 Then
                                CouleurLignes = CouleurLignesPaires
                            Else
                                CouleurLignes = CouleurLignesImpaires
                            End If
                        End If
                        VSFGEnCours.Cell(flexcpBackColor, a, COLONNES_DETAILS_EN_COURS.C_NUM_COMMANDE_INTERNE, _
                                                                                   a, COLONNES_DETAILS_EN_COURS.C_MATIERE) = CouleurLignes
                    
                    End With

                Next a
                
                '--- rafraichissement ---
                .Redraw = True
                
                '--- restitution des valeurs ligne, colonne ---
                .Row = MemLigne
                .Col = MemColonne

            End With

        Case Else

    End Select

End Sub

'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Rôle      : Affiche la totalité des données des redresseurs sur le synoptique
' Entrées :
' Retours :
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Private Sub AffichageDonneesRedresseursSurSynoptique()
    
    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next
    
    '--- constantes privées ---
    
    '--- déclaration ---
    Dim a As Integer                'pour les boucles FOR...NEXT
    Dim TempsRestantRedresseur As Integer

    For a = REDRESSEURS.R_C13 To REDRESSEURS.R_C19
    
        With TEtatsRedresseurs(a)
        
            '--- état du redresseur ---
            Select Case .EtatRedresseur
                
                Case ETATS_REDRESSEUR.ER_ARRET To ETATS_REDRESSEUR.ER_EXCLUSION
                    OCXRedresseurs(a).Etat = .EtatRedresseur
                
                Case Else
                    OCXRedresseurs(a).Etat = ETAT_NON_DEFINI
            
            End Select
                
            '--- mode du redresseur ---
            Select Case .ModeRedresseur
                
                Case MODES_REDRESSEUR.MR_MANUEL
                    OCXRedresseurs(a).Mode = MODE_MANUEL
                
                Case MODES_REDRESSEUR.MR_AUTOMATIQUE
                    OCXRedresseurs(a).Mode = MODE_AUTOMATIQUE
                
                Case Else
                    OCXRedresseurs(a).Mode = MODE_NON_DEFINI
            
            End Select
            
            '--- forçage du sens ---
            Select Case a
                Case REDRESSEURS.R_C13 To REDRESSEURS.R_C16: .SensRedresseur = SENS_ANODIQUE
                Case REDRESSEURS.R_C19: .SensRedresseur = SENS_SPECTRO
                Case Else
           End Select
        
            '--- sens ---
            OCXRedresseurs(a).Sens = .SensRedresseur
            
            '--- tension ---
            If .EtatRedresseur = ETATS_REDRESSEUR.ER_ARRET Then
                OCXRedresseurs(a).Tension = 0
            Else
                OCXRedresseurs(a).Tension = .U
            End If
            
            '--- intensité ---
            If .EtatRedresseur = ETATS_REDRESSEUR.ER_ARRET Then
                OCXRedresseurs(a).Intensite = 0
            Else
                OCXRedresseurs(a).Intensite = .I
            End If
            
            '--- ah ---
            If .EtatRedresseur = ETATS_REDRESSEUR.ER_ARRET Then
                OCXRedresseurs(a).Ah = 0
            Else
                OCXRedresseurs(a).Ah = .Ah
            End If
            
            '--- temps restant de la phase (99:59 possible) ---
            If .TempsPhaseEnCours > 0 And .TempsEcoulePhaseEnCours > 0 Then
                OCXRedresseurs(a).TempsRestantPhase = CTemps3(Abs(.TempsPhaseEnCours - .TempsEcoulePhaseEnCours))
            Else
                OCXRedresseurs(a).TempsRestantPhase = "-"
            End If
            
            '--- vu-mètre de la phase en cours ---
            Select Case .NumPhaseEnCours
                
                Case PHASES_GAMMES_REDRESSEURS.PH_T1 To PHASES_GAMMES_REDRESSEURS.PH_T4
                    OCXRedresseurs(a).Phase = .NumPhaseEnCours
                
                Case Else
                    OCXRedresseurs(a).Phase = ETEINT
            
            End Select
            
        End With

    Next a

End Sub

'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Rôle      : Visualise tous les états de la ligne
' Détails  :
' Retours :
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Sub VisualisationEtatsLigne()
        
    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next

    '--- constantes privées ---
    Const PONT_MAINTENANCE As String = "MAINTENANCE"
    Const PONT_MANUEL As String = "MANUEL"
    Const PONT_SEMI_AUTOMATIQUE As String = "SEMI-AUTO."
    Const PONT_AUTOMATIQUE As String = "AUTOMATIQUE"
    
    Const SEQUENCE_INCONNU = "-" & vbCrLf & "-"
    Const SEQUENCE_CYCLIQUE_PAR_IMPULSIONS = "MODE CYCLIQUE" & vbCrLf & "PAR IMPULSIONS"
    Const SEQUENCE_CYCLIQUE_OPTIMISE = "MODE CYCLIQUE" & vbCrLf & "OPTIMISE"
    Const SEQUENCE_ALEATOIRE = "MODE" & vbCrLf & "<< ALEATOIRE >>"
    
    Const REDRESSEUR_MANUEL As String = "Manu."
    Const REDRESSEUR_AUTOMATIQUE As String = "Auto."
    Const REDRESSEUR_ANODIQUE As String = "Anod..."
    Const REDRESSEUR_CATHODIQUE As String = "Catho..."
    Const REDRESSEUR_POLARISATION As String = "Polari."
    Const REDRESSEUR_AMORCAGE As String = "Amorç."
        
    Const RECTANGLE_VERT As String = "rectangle vert"
    Const RECTANGLE_ORANGE As String = "rectangle orange"
    Const RECTANGLE_ROUGE As String = "rectangle rouge"
    
    Const CHARIOT_PRESENT As String = "chariot present"
    Const CHARIOT_PRESENT_VERROUILLE As String = "chariot present verrouille"
    Const CHARIOT_PRESENT_VERROUILLE_CHARGE As String = "chariot present verrouille charge"
    
    Const CROIX_DE_CONDAMNATION As String = "croix de condamnation 1"
    
    Const CHRONOMETRE_BLANC As String = "chronometre fond blanc"
    Const CHRONOMETRE_CYAN As String = "chronometre fond cyan"
    Const CHRONOMETRE_ORANGE As String = "chronometre fond orange"
    
    Const CHARGE_NUMERO As String = "Charge "
    
    '--- déclaration ---
    Dim a As Integer, _
            b As Integer
    Dim IndexImage As Integer, _
            NumCuve As Integer
    Dim ClignotantNormal As Integer, _
           ClignotantRapide As Integer, _
           RetourControleTemperature As Integer, _
           NumCharge As Integer, _
           NumBarre As Integer
    Static CptAppels As Integer
    Dim CouleurFond As Long, CouleurPlan As Long
    Dim Texte As String
        
    '*************************************************************************************************
    '                                             CONTROLE DU CLIGNOTEMENT
    '*************************************************************************************************
    CptAppels = IIf(CptAppels > 11, 1, CptAppels + 1)
    ClignotantNormal = Choose(CptAppels, 1, 1, 0, 0, 1, 1, 0, 0, 1, 1, 0, 0)
    ClignotantRapide = Choose(CptAppels, 0, 1, 0, 1, 0, 1, 0, 1, 0, 1, 0, 1)
    
    '*************************************************************************************************
    '                                                  GESTION DES EN COURS
    '*************************************************************************************************
    GestionEnCours GESTION_GRILLES.GG_AFFICHAGE
    
    '*************************************************************************************************
    '                                                ETAT GENERAL DE LA LIGNE
    '*************************************************************************************************
    '--- marche générale ---
    With LMarcheGenerale
        If TEtatsLigne.MarcheGenerale = True And .BackColor <> COULEURS.VERT_2 Then
            
            '--- lancement de l'initialisation sur la marche générale ---
            InitialisationSurMarcheGenerale
            
            '--- changement des couleurs pour la marche générale ---
            .BackColor = COULEURS.VERT_2
            .ForeColor = COULEURS.NOIR
            .Refresh
        
        End If
        If TEtatsLigne.MarcheGenerale = False And .BackColor <> COULEURS.ROUGE_3 Then
            
            '--- changement des couleurs pour l'arrêt général ---
            .BackColor = COULEURS.ROUGE_3
            .ForeColor = COULEURS.JAUNE_3
            .Refresh
        
        End If
    End With
    
    '*************************************************************************************************
    '                                                         POIDS SOULEVES
    '*************************************************************************************************
    For a = LBound(TEtatsPonts()) To UBound(TEtatsPonts())
        AffichageTexte LPoidsSouleve(a), Format(TEtatsPonts(a).PoidsSouleve, FORMAT_POIDS_SOULEVE)
    Next a
    
    '*************************************************************************************************
    '                                                          REDRESSEURS
    '*************************************************************************************************
    AffichageDonneesRedresseursSurSynoptique    'ensemble des redresseurs
    
    '*************************************************************************************************
    '                      VISIBILITE OU NON DES ZONES MANUEL DES PONTS
    '*************************************************************************************************
    PBManuelP1.Visible = (TEtatsPonts(PONTS.P_1).ModePont = MODES_PONTS.M_MANUEL)
    PBManuelP2.Visible = (TEtatsPonts(PONTS.P_2).ModePont = MODES_PONTS.M_MANUEL)
    
    '*************************************************************************************************
    '                                                     ETATS DES PONTS
    '*************************************************************************************************
    For a = LBound(TEtatsPonts()) To UBound(TEtatsPonts())
        
        With TEtatsPonts(a)

            '----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
            
            '--- de maintenance à semi-automatique ---
            Select Case .ModePont
                
                Case MODES_PONTS.M_MAINTENANCE
                    '--- maintenance ---
                    With LMaintenanceAAutomatique(a)
                        If ClignotantRapide = 0 Then
                            .BackColor = COULEURS.BLANC: .ForeColor = COULEURS.BLANC
                        Else
                            .BackColor = COULEURS.ORANGE_3: .ForeColor = COULEURS.JAUNE_3
                        End If
                        If .Caption <> PONT_MAINTENANCE Then
                            .Caption = PONT_MAINTENANCE
                            .Refresh
                        End If
                    End With
                                    
                    '--- forcer le contrôle opérateur sur faux ---
                    .ControleParOperateur = False
                    
                    '--- forcer le type de séquence inconnu ---
                    .TypeSequence = TYPES_SEQUENCES.TS_INCONNU
                
                Case MODES_PONTS.M_MANUEL
                    '--- manuel ---
                    With LMaintenanceAAutomatique(a)
                        If ClignotantRapide = 0 Then
                            .BackColor = COULEURS.BLANC: .ForeColor = COULEURS.BLANC
                        Else
                            .BackColor = COULEURS.ORANGE_3: .ForeColor = COULEURS.JAUNE_3
                        End If
                        If .Caption <> PONT_MANUEL Then
                            .Caption = PONT_MANUEL
                            .Refresh
                        End If
                    End With
                    
                    '--- forcer le contrôle opérateur sur faux ---
                    .ControleParOperateur = False
                    
                    '--- forcer le type de séquence inconnu ---
                    .TypeSequence = TYPES_SEQUENCES.TS_INCONNU
                   
                Case MODES_PONTS.M_SEMI_AUTOMATIQUE
                    '--- semi-automatique ---
                    With LMaintenanceAAutomatique(a)
                        If ClignotantRapide = 0 Then
                            .BackColor = COULEURS.BLANC: .ForeColor = COULEURS.BLANC
                        Else
                            .BackColor = COULEURS.VERT_3: .ForeColor = COULEURS.NOIR
                        End If
                        If .Caption <> PONT_SEMI_AUTOMATIQUE Then
                            .Caption = PONT_SEMI_AUTOMATIQUE
                            .Refresh
                        End If
                    End With
                    
                    '--- forcer le contrôle opérateur sur faux ---
                    .ControleParOperateur = False

                    '--- forcer le type de séquence sur cyclique ---
                    .TypeSequence = TYPES_SEQUENCES.TS_CYCLIQUE_PAR_IMPULSIONS
                
                Case MODES_PONTS.M_AUTOMATIQUE
                    '--- automatique ---
                    With LMaintenanceAAutomatique(a)
                        If .Caption <> PONT_AUTOMATIQUE Then
                            .BackColor = COULEURS.VERT_3
                            .ForeColor = COULEURS.NOIR
                            .Caption = PONT_AUTOMATIQUE
                            .Refresh
                        End If
                    End With
                
                Case Else
            
            End Select

            '----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

            '--- priorité à l'indication de la condamnation ---
            If .Condamnation = True Then
                
                '--- affichage de la croix de condamnation ---
                If IEtatsPonts(a).Picture <> ILOutilsDivers.ListImages(CROIX_DE_CONDAMNATION).Picture Then
                    Set IEtatsPonts(a).Picture = ILOutilsDivers.ListImages(CROIX_DE_CONDAMNATION).Picture
                End If
                    
                '--- type de séquence inconnu dans le cas de la condamnation ---
                With LTypeSequence(a)
                    If .Caption <> SEQUENCE_INCONNU Then
                        .BackColor = COULEURS.BLANC
                        .ForeColor = COULEURS.NOIR
                        .Caption = SEQUENCE_INCONNU
                        .Refresh
                    End If
                End With
            
            Else
                
                '--- indicateur signalant au moins un défaut ---
                If TEtatsPonts(a).UnDefautAuMoinsSignale = True Then
                    If IEtatsPonts(a).Picture <> ILOutilsDivers.ListImages(RECTANGLE_ROUGE).Picture Then
                        Set IEtatsPonts(a).Picture = ILOutilsDivers.ListImages(RECTANGLE_ROUGE).Picture
                    End If
                Else
                    If IEtatsPonts(a).Picture <> ILOutilsDivers.ListImages(RECTANGLE_VERT).Picture Then
                        Set IEtatsPonts(a).Picture = ILOutilsDivers.ListImages(RECTANGLE_VERT).Picture
                    End If
                End If
                                
                '--- type de séquence du pont (de inconnu à aléatoire) ---
                Select Case .TypeSequence
            
                    Case TYPES_SEQUENCES.TS_INCONNU
                        '--- type de séquence inconnu (cas de la maintenance ou du manuel) ---
                        With LTypeSequence(a)
                            If .Caption <> SEQUENCE_INCONNU Then
                                .BackColor = COULEURS.BLANC
                                .ForeColor = COULEURS.NOIR
                                .Caption = SEQUENCE_INCONNU
                                .Refresh
                            End If
                        End With
            
                    Case TYPES_SEQUENCES.TS_CYCLIQUE_PAR_IMPULSIONS
                        '--- type de séquence cyclique par impulsions (cas du semi-automatique) ---
                        With LTypeSequence(a)
                            If .Caption <> SEQUENCE_CYCLIQUE_PAR_IMPULSIONS Then
                                .BackColor = COULEURS.VERT_3
                                .ForeColor = COULEURS.NOIR
                                .Caption = SEQUENCE_CYCLIQUE_PAR_IMPULSIONS
                                .Refresh
                            End If
                        End With
                    
                    Case TYPES_SEQUENCES.TS_CYCLIQUE_OPTIMISE
                        '--- type de séquence cyclique optimisé (automatique sans nécessité l'IA) ---
                        With LTypeSequence(a)
                            If .Caption <> SEQUENCE_CYCLIQUE_OPTIMISE Then
                                .BackColor = COULEURS.VERT_3
                                .ForeColor = COULEURS.NOIR
                                .Caption = SEQUENCE_CYCLIQUE_OPTIMISE
                                .Refresh
                            End If
                        End With
                    
                    Case TYPES_SEQUENCES.TS_ALEATOIRE
                        '--- type de séquence aléatoire (cas du pilotage par IA) ---
                        With LTypeSequence(a)
                            If .Caption <> SEQUENCE_ALEATOIRE Then
                                .BackColor = COULEURS.CYAN_2
                                .ForeColor = COULEURS.NOIR
                                .Caption = SEQUENCE_ALEATOIRE
                                .Refresh
                            End If
                        End With
            
                    Case Else
            
                End Select
        
            End If

        End With
            
        '----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
        
        '--- contrôle par l'opérateur ---
        With LControleOperateurPonts(a)
            If TEtatsPonts(a).ControleParOperateur = False Then
                If .Visible <> False Then .Visible = False
            Else
                If ClignotantNormal = 0 Then
                    If .Visible <> False Then .Visible = False
                Else
                    If .Visible <> True Then .Visible = True
                End If
            End If
        End With
    
    Next a
    
    '----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    
    '--- cadre de l'affichage du contrôle opérateur ---
    With SControleOperateur
        If TEtatsPonts(PONTS.P_1).ControleParOperateur = True Or TEtatsPonts(PONTS.P_2).ControleParOperateur = True Then
            If .Visible <> True Then .Visible = True
        Else
            If .Visible <> False Then .Visible = False
        End If
    End With
    
    '*************************************************************************************************
    '                                                     ETATS DES POSTES
    '*************************************************************************************************
    For a = POSTES.P_CHGT_1 To DERNIER_POSTE
    
        With TEtatsPostes(a)
            
            '----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
            
            '--- priorité à l'indication de la condamnation ---
            If .Condamnation = True Then
                
                '--- affichage de la croix de condamnation ---
                If IEtatsPostes(a).Picture <> ILOutilsDivers.ListImages(CROIX_DE_CONDAMNATION).Picture Then
                    Set IEtatsPostes(a).Picture = ILOutilsDivers.ListImages(CROIX_DE_CONDAMNATION).Picture
                End If
                
            Else
                    
                '--- indicateur d'un défaut sur une cuve ---
                NumCuve = CorrespondancePostesCuvesAPI(a)
                If NumCuve > 0 Then
                    If TEtatsCuves(NumCuve).UnDefautAuMoinsSignale = True Then
                        '--- passage en couleur rouge ---
                        If IEtatsPostes(a).Picture <> ILOutilsDivers.ListImages(RECTANGLE_ROUGE).Picture Then
                            IEtatsPostes(a).Picture = ILOutilsDivers.ListImages(RECTANGLE_ROUGE).Picture
                        End If
                    Else
                        '--- passage en couleur verte ---
                        If IEtatsPostes(a).Picture <> ILOutilsDivers.ListImages(RECTANGLE_VERT).Picture Then
                            IEtatsPostes(a).Picture = ILOutilsDivers.ListImages(RECTANGLE_VERT).Picture
                        End If
                    End If
                Else
                    '--- passage en couleur verte ---
                    If IEtatsPostes(a).Picture <> ILOutilsDivers.ListImages(RECTANGLE_VERT).Picture Then
                        IEtatsPostes(a).Picture = ILOutilsDivers.ListImages(RECTANGLE_VERT).Picture
                    End If
                End If
                
            End If
                
            '----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
                
            '--- numéros des charges / numéros de barres ---
            NumCharge = .NumCharge
            If NumCharge >= CHARGES.C_NUM_MINI And NumCharge <= CHARGES.C_NUM_MAXI Then
                NumBarre = TEtatsCharges(.NumCharge).NumBarre
            Else
                NumBarre = 0
            End If
            If ModeAffichageSynoptique = MA_NUM_BARRES And _
               NumBarre >= BARRES.B_NUM_MINI And NumBarre <= BARRES.B_NUM_MAXI Then
                AffichageTexte LNumCharges(a), TBarres(NumBarre).Libelle, COULEURS.VERT_3
            Else
                If NumCharge >= CHARGES.C_NUM_MINI And NumCharge <= CHARGES.C_NUM_MAXI Then
                    AffichageTexte LNumCharges(a), TBarres(NumBarre).Libelle, COULEURS.JAUNE_3
                Else
                    AffichageTexte LNumCharges(a), "", COULEURS.BLANC
                End If
            End If

            '----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
            
            Select Case a

                Case POSTES.P_CHGT_1 To POSTES.P_CHGT_2
                    '--- états des chariots pour les postes de chargement ---
                    Select Case .EtatsChariots

                        Case ETATS_CHARIOTS.E_ABSENT
                            '--- chariot absent ---
                            If IEtatsChariots(a).Picture <> ILOutilsDivers.ListImages(RECTANGLE_VERT).Picture Then
                                Set IEtatsChariots(a).Picture = ILOutilsDivers.ListImages(RECTANGLE_VERT).Picture
                            End If

                        Case ETATS_CHARIOTS.E_PRESENT
                            '--- chariot présent ---
                            If IEtatsChariots(a).Picture <> ILOutilsDivers.ListImages(CHARIOT_PRESENT).Picture Then
                                Set IEtatsChariots(a).Picture = ILOutilsDivers.ListImages(CHARIOT_PRESENT).Picture
                            End If

                        Case ETATS_CHARIOTS.E_PRESENT_VERROUILLE
                            '--- chariot présent verrouillé avec ou sans charge ---
                            NumCharge = .NumCharge
                            If NumCharge >= CHARGES.C_NUM_MINI And NumCharge <= CHARGES.C_NUM_MAXI Then
                                If IEtatsChariots(a).Picture <> ILOutilsDivers.ListImages(CHARIOT_PRESENT_VERROUILLE_CHARGE).Picture Then
                                    Set IEtatsChariots(a).Picture = ILOutilsDivers.ListImages(CHARIOT_PRESENT_VERROUILLE_CHARGE).Picture
                                End If
                            Else
                                If IEtatsChariots(a).Picture <> ILOutilsDivers.ListImages(CHARIOT_PRESENT_VERROUILLE).Picture Then
                                    Set IEtatsChariots(a).Picture = ILOutilsDivers.ListImages(CHARIOT_PRESENT_VERROUILLE).Picture
                                End If
                            End If
                
                        Case Else
                    End Select
                
                Case POSTES.P_D1 To POSTES.P_D2
                    '--- états des chariots pour les postes de déchargement ---
                    Select Case .EtatsChariots

                        Case ETATS_CHARIOTS.E_ABSENT
                            '--- chariot absent ---
                            If IEtatsChariots(a).Picture <> ILOutilsDivers.ListImages(RECTANGLE_VERT).Picture Then
                                Set IEtatsChariots(a).Picture = ILOutilsDivers.ListImages(RECTANGLE_VERT).Picture
                            End If

                        Case ETATS_CHARIOTS.E_PRESENT
                            '--- chariot présent ---
                            If IEtatsChariots(a).Picture <> ILOutilsDivers.ListImages(CHARIOT_PRESENT).Picture Then
                                Set IEtatsChariots(a).Picture = ILOutilsDivers.ListImages(CHARIOT_PRESENT).Picture
                            End If

                        Case ETATS_CHARIOTS.E_PRESENT_VERROUILLE
                            '--- chariot présent verrouillé avec ou sans charge ---
                            NumCharge = .NumCharge
                            If NumCharge >= CHARGES.C_NUM_MINI And NumCharge <= CHARGES.C_NUM_MAXI Then
                                If IEtatsChariots(a).Picture <> ILOutilsDivers.ListImages(CHARIOT_PRESENT_VERROUILLE_CHARGE).Picture Then
                                    Set IEtatsChariots(a).Picture = ILOutilsDivers.ListImages(CHARIOT_PRESENT_VERROUILLE_CHARGE).Picture
                                End If
                            Else
                                If IEtatsChariots(a).Picture <> ILOutilsDivers.ListImages(CHARIOT_PRESENT_VERROUILLE).Picture Then
                                    Set IEtatsChariots(a).Picture = ILOutilsDivers.ListImages(CHARIOT_PRESENT_VERROUILLE).Picture
                                End If
                            End If

                        Case Else
                    End Select
                
                Case Else
                    '--- autres postes ---
            
            End Select
    
        End With
    
    Next a
    
    '*************************************************************************************************
    '                                                CUVES GEREES PAR l'API
    '*************************************************************************************************
    '--- programmateur cyclique et températures des cuves ---
    For a = CUVES_REGULATION.C_C00 To DERNIERE_CUV_REGULATION
        
        '----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
        
        '--- mode la régulation ---
        Select Case TEtatsCuves(a).ModeRegulation
        
            Case MODES_REGULATION.MR_MANUEL
                '--- mode manuel de la régulation ---
                Texte = "MANUEL"
                If ClignotantRapide = 0 Then
                    CouleurFond = COULEURS.BLANC: CouleurPlan = COULEURS.BLANC
                Else
                    CouleurFond = COULEURS.ORANGE_2: CouleurPlan = COULEURS.NOIR
                End If
            
            Case MODES_REGULATION.MR_AUTOMATIQUE
                '--- mode automatique de la régulation ---
                Texte = "AUTO."
                CouleurFond = COULEURS.VERT_3: CouleurPlan = COULEURS.NOIR
            
            Case Else
        End Select
        
        '--- affichage ---
        AffichageTexte LManuAutoRegulation(TEtatsCuves(a).IndexAutomate), Texte, CouleurFond, CouleurPlan
        
        '----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
        
        '--- mode de production de chaque cuve ---
        Select Case TEtatsCuves(a).ModeProduction
        
            Case MODES_PRODUCTION.M_ARRET
                '--- mode arrêt ---
                If IMProgrammateurCycliqueCuves(TEtatsCuves(a).IndexAutomate).Picture <> ILProgrammateurCycliqueCuves.ListImages(CHRONOMETRE_BLANC).Picture Then
                    Set IMProgrammateurCycliqueCuves(TEtatsCuves(a).IndexAutomate).Picture = ILProgrammateurCycliqueCuves.ListImages(CHRONOMETRE_BLANC).Picture
                End If
        
            Case MODES_PRODUCTION.M_VEILLE
                '--- mode veille ---
                If IMProgrammateurCycliqueCuves(TEtatsCuves(a).IndexAutomate).Picture <> ILProgrammateurCycliqueCuves.ListImages(CHRONOMETRE_CYAN).Picture Then
                    Set IMProgrammateurCycliqueCuves(TEtatsCuves(a).IndexAutomate).Picture = ILProgrammateurCycliqueCuves.ListImages(CHRONOMETRE_CYAN).Picture
                End If
        
            Case MODES_PRODUCTION.M_PRODUCTION
                '--- mode de production ---
                If IMProgrammateurCycliqueCuves(TEtatsCuves(a).IndexAutomate).Picture <> ILProgrammateurCycliqueCuves.ListImages(CHRONOMETRE_ORANGE).Picture Then
                    Set IMProgrammateurCycliqueCuves(TEtatsCuves(a).IndexAutomate).Picture = ILProgrammateurCycliqueCuves.ListImages(CHRONOMETRE_ORANGE).Picture
                End If
        
            Case Else
        End Select

        '----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

        '--- températures ---
        With TEtatsCuves(a).Temperatures
            RetourControleTemperature = ControleTemperature(a)
            Select Case RetourControleTemperature
                Case CONTROLES_TEMPERATURES.C_PAS_DE_CONTROLE: CouleurFond = COULEURS.VERT_3: CouleurPlan = COULEURS.NOIR
                Case CONTROLES_TEMPERATURES.C_TEMPERATURE_NORMALE: CouleurFond = ORANGE_CUVE: CouleurPlan = COULEURS.NOIR
                Case CONTROLES_TEMPERATURES.C_TEMPERATURE_INFERIEURE, _
                         CONTROLES_TEMPERATURES.C_TEMPERATURE_SUPERIEURE, _
                         CONTROLES_TEMPERATURES.C_DEFAUT_PT100
                    If ClignotantRapide = 0 Then
                        CouleurFond = COULEURS.BLANC: CouleurPlan = COULEURS.BLANC
                    Else
                        CouleurFond = ROUGE_DEFAUT: CouleurPlan = COULEURS.JAUNE_3
                    End If
                Case Else
            End Select
            Select Case RetourControleTemperature
                Case CONTROLES_TEMPERATURES.C_DEFAUT_PT100: Texte = "PT100"
                Case Else: Texte = Format(.TempActuelle, FORMAT_TEMPERATURE_COMPACTE_1_DECIMALE_UNITE)
            End Select
            AffichageTexte OccFSynoptique.LTemperatures(TEtatsCuves(a).IndexAutomate), Texte, CouleurFond, CouleurPlan
        End With
    
        '----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
        
        '--- niveaux ---
        If TEtatsCuves(a).DefinitionCuve.PresenceNiveauBas = True Or TEtatsCuves(a).DefinitionCuve.PresenceNiveauHaut = True Then
            
            Select Case TEtatsCuves(a).EtatsNiveaux
                
                Case ETATS_NIVEAUX.E_TRES_BAS
                    Texte = "TRES BAS"
                
                Case ETATS_NIVEAUX.E_INTERMEDIAIRE_BAS
                    Texte = "INT. BAS"
                    CouleurFond = COULEURS.ORANGE_2: CouleurPlan = COULEURS.NOIR
                
                Case ETATS_NIVEAUX.E_NORMAL
                    Texte = "NORMAL"
                    CouleurFond = COULEURS.VERT_3: CouleurPlan = COULEURS.NOIR
                
                Case ETATS_NIVEAUX.E_INTERMEDIAIRE_HAUT
                    Texte = "INT. HAUT"
                    CouleurFond = COULEURS.VERT_3: CouleurPlan = COULEURS.NOIR
                
                Case ETATS_NIVEAUX.E_TRES_HAUT
                    Texte = "TRES HAUT"
                
                Case Else
            End Select
            
            '--- couleurs clignotantes pour les niveaux importants ---
            If TEtatsCuves(a).EtatsNiveaux = ETATS_NIVEAUX.E_TRES_BAS Or TEtatsCuves(a).EtatsNiveaux = ETATS_NIVEAUX.E_TRES_HAUT Then
                If ClignotantRapide = 0 Then
                    CouleurFond = COULEURS.BLANC: CouleurPlan = COULEURS.BLANC
                Else
                    CouleurFond = ROUGE_DEFAUT: CouleurPlan = COULEURS.JAUNE_3
                End If
            Else
                
            End If
            
            '--- affichage ---
            AffichageTexte LNiveaux(TEtatsCuves(a).IndexAutomate), Texte, CouleurFond, CouleurPlan
        
        End If
    
    Next a
    
    '*************************************************************************************************
    '                                  ENTRETIEN DES GRAPHES DE PRODUCTION
    '*************************************************************************************************
        
    With CBEntretienGraphesProduction
    
        If EntretienGraphesProduction = True Then
            If ClignotantRapide = 0 Then
                .BackColor = COULEURS.JAUNE_1
            Else
                .BackColor = COULEURS.ROUGE_2
            End If
        End If

        .Visible = EntretienGraphesProduction

    End With
    
    '*************************************************************************************************
    '                                                              ANNEXES
    '*************************************************************************************************
    
    'With OccFSynoptique
    
        '--- ventilation ---
        'With .LEtatsAnnexes(INDEX_CHAMPS.IDX_CHAMP_VENTILATION)
        '    If TEtatsAnnexes.EtatsVentilation = ETATS_VENTILATION.E_DEFAUT Then
        '        If .BackColor <> ROUGE_DEFAUT Then
        '            .BackColor = ROUGE_DEFAUT
        '            .Refresh
        '        End If
        '    Else
        '        If .BackColor <> COULEURS.VERT_3 Then
        '            .BackColor = COULEURS.VERT_3
        '            .Refresh
        '        End If
        '    End If
        'End With
        
        '--- volet de compensation ---
        'With .LEtatsAnnexes(INDEX_CHAMPS.IDX_CHAMP_VOLET_COMPENSATION)
        '    If TEtatsAnnexes.EtatsVoletCompensation = ETATS_VOLET_COMPENSATION.E_DEFAUT Then
        '        If .BackColor <> ROUGE_DEFAUT Then
        '            .BackColor = ROUGE_DEFAUT
        '            .Refresh
        '        End If
        '    Else
        '        If .BackColor <> COULEURS.VERT_3 Then
        '            .BackColor = COULEURS.VERT_3
        '            .Refresh
        '        End If
        '    End If
        'End With
        
        '--- air comprimé ---
        'With .LEtatsAnnexes(INDEX_CHAMPS.IDX_CHAMP_AIR_COMPRIME)
        '    If TEtatsLigne.ManqueAir = True Then
        '        If .BackColor <> ROUGE_DEFAUT Then
        '            .BackColor = ROUGE_DEFAUT
        '            .Refresh
        '        End If
        '    Else
        '        If .BackColor <> COULEURS.VERT_3 Then
        '            .BackColor = COULEURS.VERT_3
        '            .Refresh
        '        End If
        '    End If
        'End With
        
        '--- surpresseur d'air ---
        'With .LEtatsAnnexes(INDEX_CHAMPS.IDX_CHAMP_SURPRESSEUR_AIR)
        '    If TEtatsAnnexes.EtatsSurpresseurAir = ETATS_SURPRESSEUR_AIR.E_DEFAUT Then
        '        If .BackColor <> ROUGE_DEFAUT Then
        '            .BackColor = ROUGE_DEFAUT
        '            .Refresh
        '        End If
        '    Else
        '        If .BackColor <> COULEURS.VERT_3 Then
        '            .BackColor = COULEURS.VERT_3
        '            .Refresh
        '        End If
        '    End If
        'End With
        
    'End With
    
End Sub



VERSION 5.00
Object = "{0AFE7BE0-11B7-4A3E-978D-D4501E9A57FE}#1.0#0"; "c1sizer.ocx"
Object = "{1C0489F8-9EFD-423D-887A-315387F18C8F}#1.0#0"; "vsflex8l.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form FModeCyclique 
   Caption         =   "MODE CYCLIQUE"
   ClientHeight    =   12900
   ClientLeft      =   930
   ClientTop       =   1050
   ClientWidth     =   20730
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FModeCyclique.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   12900
   ScaleWidth      =   20730
   WindowState     =   2  'Maximized
   Begin VB.PictureBox PBDeplacementFenetre 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   13590
      Index           =   0
      Left            =   0
      ScaleHeight     =   13590
      ScaleWidth      =   20730
      TabIndex        =   3
      Top             =   375
      Width           =   20730
      Begin VB.PictureBox PBDeplacementFenetre 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   13365
         Index           =   1
         Left            =   0
         ScaleHeight     =   13365
         ScaleWidth      =   28740
         TabIndex        =   4
         Top             =   0
         Width           =   28740
         Begin C1SizerLibCtl.C1Tab CTOnglets 
            Height          =   12435
            Left            =   180
            TabIndex        =   9
            Top             =   180
            Width           =   28365
            _cx             =   50033
            _cy             =   21934
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
            Caption         =   "Général"
            Align           =   0
            CurrTab         =   0
            FirstTab        =   0
            Style           =   1
            Position        =   0
            AutoSwitch      =   -1  'True
            AutoScroll      =   -1  'True
            TabPreview      =   -1  'True
            ShowFocusRect   =   0   'False
            TabsPerPage     =   10
            BorderWidth     =   0
            BoldCurrent     =   -1  'True
            DogEars         =   -1  'True
            MultiRow        =   -1  'True
            MultiRowOffset  =   0
            CaptionStyle    =   0
            TabHeight       =   420
            TabCaptionPos   =   4
            TabPicturePos   =   0
            CaptionEmpty    =   ""
            Separators      =   0   'False
            AccessibleName  =   ""
            AccessibleDescription=   ""
            AccessibleValue =   ""
            AccessibleRole  =   37
            Begin VB.PictureBox PBOnglets 
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   11925
               Index           =   3
               Left            =   29010
               ScaleHeight     =   11865
               ScaleWidth      =   28215
               TabIndex        =   11
               Top             =   465
               Width           =   28275
            End
            Begin VB.PictureBox PBOnglets 
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   11925
               Index           =   0
               Left            =   45
               ScaleHeight     =   11865
               ScaleWidth      =   28215
               TabIndex        =   10
               Top             =   465
               Width           =   28275
               Begin VSFlex8LCtl.VSFlexGrid VSFGModificationNumChargesPostes 
                  Height          =   9975
                  Left            =   180
                  TabIndex        =   14
                  Top             =   1140
                  Width           =   8955
                  _cx             =   15796
                  _cy             =   17595
                  Appearance      =   0
                  BorderStyle     =   1
                  Enabled         =   -1  'True
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  MousePointer    =   0
                  BackColor       =   16777152
                  ForeColor       =   12582912
                  BackColorFixed  =   65280
                  ForeColorFixed  =   -2147483640
                  BackColorSel    =   255
                  ForeColorSel    =   -2147483633
                  BackColorBkg    =   12640511
                  BackColorAlternate=   16777152
                  GridColor       =   -2147483633
                  GridColorFixed  =   -2147483633
                  TreeColor       =   -2147483633
                  FloodColor      =   -2147483633
                  SheetBorder     =   16777215
                  FocusRect       =   3
                  HighLight       =   0
                  AllowSelection  =   0   'False
                  AllowBigSelection=   0   'False
                  AllowUserResizing=   0
                  SelectionMode   =   0
                  GridLines       =   2
                  GridLinesFixed  =   2
                  GridLineWidth   =   1
                  Rows            =   4
                  Cols            =   3
                  FixedRows       =   2
                  FixedCols       =   2
                  RowHeightMin    =   0
                  RowHeightMax    =   0
                  ColWidthMin     =   0
                  ColWidthMax     =   0
                  ExtendLastCol   =   0   'False
                  FormatString    =   $"FModeCyclique.frx":014A
                  ScrollTrack     =   0   'False
                  ScrollBars      =   2
                  ScrollTips      =   0   'False
                  MergeCells      =   0
                  MergeCompare    =   0
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
                  TabBehavior     =   1
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
                  ForeColorFrozen =   -2147483641
                  WallPaperAlignment=   10
                  AccessibleName  =   ""
                  AccessibleDescription=   ""
                  AccessibleValue =   ""
                  AccessibleRole  =   24
               End
               Begin VB.CommandButton CBTransfertNumChargesPontsVersAPI 
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
                  Left            =   14880
                  Style           =   1  'Graphical
                  TabIndex        =   21
                  ToolTipText     =   " Transfère les valeurs dans l'automate "
                  Top             =   11220
                  Width           =   1515
               End
               Begin VB.CommandButton CBAnnulerTransfertNumChargesPontsVersAPI 
                  BackColor       =   &H00C0FFFF&
                  Caption         =   "Annuler TOUT"
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
                  Left            =   9480
                  Style           =   1  'Graphical
                  TabIndex        =   20
                  ToolTipText     =   " Annule l'entrée des données "
                  Top             =   11220
                  Width           =   1515
               End
               Begin VB.CommandButton CBTransfertNumChargesPostesVersAPI 
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
                  Left            =   7500
                  Style           =   1  'Graphical
                  TabIndex        =   13
                  ToolTipText     =   " Transfère les valeurs dans l'automate "
                  Top             =   11220
                  Width           =   1515
               End
               Begin VB.CommandButton CBAnnulerTransfertNumChargesPostesVersAPI 
                  BackColor       =   &H00C0FFFF&
                  Caption         =   "Annuler TOUT"
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
                  Left            =   300
                  Style           =   1  'Graphical
                  TabIndex        =   12
                  ToolTipText     =   " Annule l'entrée des données "
                  Top             =   11220
                  Width           =   1515
               End
               Begin VSFlex8LCtl.VSFlexGrid VSFGModificationNumChargesPonts 
                  Height          =   9975
                  Left            =   9360
                  TabIndex        =   19
                  Top             =   1140
                  Width           =   7155
                  _cx             =   12621
                  _cy             =   17595
                  Appearance      =   0
                  BorderStyle     =   1
                  Enabled         =   -1  'True
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  MousePointer    =   0
                  BackColor       =   16777152
                  ForeColor       =   12582912
                  BackColorFixed  =   128
                  ForeColorFixed  =   -2147483639
                  BackColorSel    =   255
                  ForeColorSel    =   -2147483633
                  BackColorBkg    =   12640511
                  BackColorAlternate=   16777152
                  GridColor       =   -2147483633
                  GridColorFixed  =   -2147483633
                  TreeColor       =   -2147483633
                  FloodColor      =   -2147483633
                  SheetBorder     =   16777215
                  FocusRect       =   3
                  HighLight       =   0
                  AllowSelection  =   0   'False
                  AllowBigSelection=   0   'False
                  AllowUserResizing=   0
                  SelectionMode   =   0
                  GridLines       =   2
                  GridLinesFixed  =   2
                  GridLineWidth   =   1
                  Rows            =   4
                  Cols            =   3
                  FixedRows       =   2
                  FixedCols       =   2
                  RowHeightMin    =   0
                  RowHeightMax    =   0
                  ColWidthMin     =   0
                  ColWidthMax     =   0
                  ExtendLastCol   =   0   'False
                  FormatString    =   $"FModeCyclique.frx":01A5
                  ScrollTrack     =   0   'False
                  ScrollBars      =   0
                  ScrollTips      =   0   'False
                  MergeCells      =   0
                  MergeCompare    =   0
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
                  TabBehavior     =   1
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
                  ForeColorFrozen =   -2147483639
                  WallPaperAlignment=   10
                  AccessibleName  =   ""
                  AccessibleDescription=   ""
                  AccessibleValue =   ""
                  AccessibleRole  =   24
               End
               Begin VB.Shape SDecorationModificationNumChargesPonts 
                  BackColor       =   &H00000080&
                  BackStyle       =   1  'Opaque
                  Height          =   615
                  Left            =   9360
                  Top             =   11100
                  Width           =   7155
               End
               Begin VB.Label LTitreModificationNumChargesPonts 
                  Alignment       =   2  'Center
                  Appearance      =   0  'Flat
                  BackColor       =   &H00000080&
                  BorderStyle     =   1  'Fixed Single
                  Caption         =   "MODIFICATION des NUMEROS de CHARGES pour les PONTS"
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
                  Left            =   9360
                  TabIndex        =   18
                  Top             =   900
                  Width           =   7155
                  WordWrap        =   -1  'True
               End
               Begin VB.Label LLegendeModificationNumCharges 
                  Alignment       =   2  'Center
                  Appearance      =   0  'Flat
                  BackColor       =   &H000000C0&
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
                  ForeColor       =   &H0000FFFF&
                  Height          =   555
                  Left            =   180
                  TabIndex        =   16
                  Top             =   180
                  Width           =   16335
               End
               Begin VB.Shape SDecorationModificationNumChargesPostes 
                  BackColor       =   &H00404000&
                  BackStyle       =   1  'Opaque
                  Height          =   615
                  Left            =   180
                  Top             =   11100
                  Width           =   8955
               End
               Begin VB.Label LTitreModificationNumChargesPostes 
                  Alignment       =   2  'Center
                  Appearance      =   0  'Flat
                  BackColor       =   &H0000C0C0&
                  BorderStyle     =   1  'Fixed Single
                  Caption         =   "MODIFICATION des NUMEROS de CHARGES pour les POSTES"
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00000000&
                  Height          =   255
                  Left            =   180
                  TabIndex        =   15
                  Top             =   900
                  Width           =   8955
                  WordWrap        =   -1  'True
               End
               Begin VB.Shape SFocusModificationNumChargesPostes 
                  BorderColor     =   &H000000FF&
                  BorderWidth     =   5
                  Height          =   435
                  Left            =   120
                  Top             =   720
                  Visible         =   0   'False
                  Width           =   615
               End
               Begin VB.Shape SFocusModificationNumChargesPonts 
                  BorderColor     =   &H000000FF&
                  BorderWidth     =   5
                  Height          =   435
                  Left            =   9300
                  Top             =   720
                  Visible         =   0   'False
                  Width           =   615
               End
            End
         End
      End
   End
   Begin VB.PictureBox PBRenseignementsFenetre 
      Align           =   1  'Align Top
      BackColor       =   &H00FF0000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   0
      Picture         =   "FModeCyclique.frx":0200
      ScaleHeight     =   315
      ScaleWidth      =   20670
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   0
      Width           =   20730
      Begin VB.Label LRenseignementsFenetre 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00008000&
         BackStyle       =   0  'Transparent
         Caption         =   "MODE CYCLIQUE"
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
         Left            =   600
         TabIndex        =   2
         Top             =   0
         Width           =   11415
         WordWrap        =   -1  'True
      End
   End
   Begin VB.PictureBox PBBoutons 
      Align           =   2  'Align Bottom
      DrawStyle       =   2  'Dot
      DrawWidth       =   16891
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   0
      ScaleHeight     =   1035
      ScaleWidth      =   20670
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   11805
      Width           =   20730
      Begin VB.PictureBox PBOutilsDeplacementFenetre 
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   915
         Left            =   0
         ScaleHeight     =   855
         ScaleWidth      =   1155
         TabIndex        =   5
         Top             =   0
         Visible         =   0   'False
         Width           =   1215
         Begin VB.CommandButton CBAgrandirFenetre 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Agrandir"
            DownPicture     =   "FModeCyclique.frx":24B42
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
            Picture         =   "FModeCyclique.frx":24CEC
            Style           =   1  'Graphical
            TabIndex        =   8
            ToolTipText     =   " Agrandissement de la fenêtre "
            Top             =   0
            UseMaskColor    =   -1  'True
            Width           =   900
         End
         Begin VB.HScrollBar HSDeplacementFenetre 
            Height          =   240
            LargeChange     =   300
            Left            =   0
            SmallChange     =   100
            TabIndex        =   7
            Top             =   615
            Width           =   900
         End
         Begin VB.VScrollBar VSDeplacementFenetre 
            Height          =   855
            LargeChange     =   300
            Left            =   900
            SmallChange     =   100
            TabIndex        =   6
            Top             =   0
            Width           =   255
         End
      End
      Begin VB.Timer TimerSequencesAutomatique 
         Enabled         =   0   'False
         Interval        =   500
         Left            =   1320
         Top             =   480
      End
      Begin MSComctlLib.ImageList ILOnglets 
         Left            =   1920
         Top             =   120
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   182
         ImageHeight     =   16
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   5
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FModeCyclique.frx":24E96
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FModeCyclique.frx":27128
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FModeCyclique.frx":293BA
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FModeCyclique.frx":2B64C
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FModeCyclique.frx":2D8DE
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin MSComctlLib.ImageList ILGrillesDonnees 
         Left            =   2580
         Top             =   120
         _ExtentX        =   794
         _ExtentY        =   794
         BackColor       =   -2147483643
         ImageWidth      =   12
         ImageHeight     =   12
         MaskColor       =   16711935
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   26
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FModeCyclique.frx":2FB70
               Key             =   "fleche noire"
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FModeCyclique.frx":2FD7C
               Key             =   "fleche blanche"
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FModeCyclique.frx":2FF88
               Key             =   "fleche grise"
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FModeCyclique.frx":30194
               Key             =   "fleche rouge"
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FModeCyclique.frx":303A0
               Key             =   "fleche jaune"
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FModeCyclique.frx":305AC
               Key             =   "fleche verte"
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FModeCyclique.frx":307B8
               Key             =   "fleche cyan"
            EndProperty
            BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FModeCyclique.frx":309C4
               Key             =   "fleche bleue"
            EndProperty
            BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FModeCyclique.frx":30BD0
               Key             =   "etoile noire"
            EndProperty
            BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FModeCyclique.frx":30DDC
               Key             =   "etoile blanche"
            EndProperty
            BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FModeCyclique.frx":30FE8
               Key             =   "etoile grise"
            EndProperty
            BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FModeCyclique.frx":311F4
               Key             =   "etoile rouge"
            EndProperty
            BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FModeCyclique.frx":31400
               Key             =   "etoile jaune"
            EndProperty
            BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FModeCyclique.frx":3160C
               Key             =   "etoile verte"
            EndProperty
            BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FModeCyclique.frx":31818
               Key             =   "etoile cyan"
            EndProperty
            BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FModeCyclique.frx":31A24
               Key             =   "etoile bleue"
            EndProperty
            BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FModeCyclique.frx":31C30
               Key             =   "modification noire"
            EndProperty
            BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FModeCyclique.frx":31E34
               Key             =   "modification blanche"
            EndProperty
            BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FModeCyclique.frx":32038
               Key             =   "modification grise"
            EndProperty
            BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FModeCyclique.frx":3223C
               Key             =   "modification rouge"
            EndProperty
            BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FModeCyclique.frx":32440
               Key             =   "modification jaune"
            EndProperty
            BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FModeCyclique.frx":32644
               Key             =   "modification vert"
            EndProperty
            BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FModeCyclique.frx":32848
               Key             =   "modification cyan"
            EndProperty
            BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FModeCyclique.frx":32A4C
               Key             =   "modification bleue"
            EndProperty
            BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FModeCyclique.frx":32C50
               Key             =   "indicateur vert"
            EndProperty
            BeginProperty ListImage26 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FModeCyclique.frx":32E54
               Key             =   "indicateur rouge"
            EndProperty
         EndProperty
      End
      Begin VB.CommandButton CBQuitter 
         BackColor       =   &H00FFFFFF&
         Cancel          =   -1  'True
         Caption         =   "Echap = &QUITTER"
         DownPicture     =   "FModeCyclique.frx":33058
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   13320
         MaskColor       =   &H00FF00FF&
         Picture         =   "FModeCyclique.frx":3375A
         Style           =   1  'Graphical
         TabIndex        =   17
         ToolTipText     =   " Quitter cette fenêtre "
         Top             =   120
         UseMaskColor    =   -1  'True
         Width           =   2175
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
End
Attribute VB_Name = "FModeCyclique"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Rôle                    : Fenêtre gérant le mode cyclique
' Nom                    : FModeCyclique.frm
' Date de création : 07/10/2010
' Détails                :
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

'--- déclarations obligatoires ---
Option Explicit

'--- options générales ---
Option Base 1
DefVar A-Z

'--- constantes privées ---
Private Const NBR_COLONNES_MODIFICATION_NUM_CHARGES_POSTES As Integer = 5
Private Const NBR_COLONNES_MODIFICATION_NUM_CHARGES_PONT As Integer = 4

Private Const TITRE_FENETRE As String = "MODE CYCLIQUE"
Private Const TITRE_MESSAGES As String = TITRE_FENETRE

'--- énumérations privées ---
Private Enum ONGLETS
    O_GENERAL = 0
End Enum

Private Enum COLONNES_MODIFICATION_NUM_CHARGES_POSTES
    C_NUM_POSTE = 0
    C_NOM_POSTE = 1
    C_LIBELLE_POSTE = 2
    C_NUM_CHARGE = 3
    C_NOUVEAU_NUM_CHARGE = 4
End Enum

Private Enum COLONNES_MODIFICATION_NUM_CHARGES_PONTS
    C_NUM_PONT = 0
    C_LIBELLE_PONT = 1
    C_NUM_CHARGE = 2
    C_NOUVEAU_NUM_CHARGE = 3
End Enum

'--- variables privées ---
Private PremiereActivation As Boolean
Private InterdireEvenements As Boolean                  'pour interdire certains évènements

Private MemDernierBouton As Long                          'mémoire du dernier bouton

'--- tableaux privés ---

'--- variables publiques ---
Public NumFenetre As Long                                        'numéro de la fenêtre lorsqu'elle devient active
    
Private Sub CBAgrandirFENETRE_Click()
    On Error Resume Next
    Me.WindowState = vbMaximized
End Sub

Private Sub CBAnnulerTransfertNumChargesPontsVersAPI_Click()
    
    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next
    
    '--- vidage puis réaffichage des valeurs ---
    GestionModificationNumChargesPonts GESTION_GRILLES.GG_VIDAGE

    '--- gestion des boutons ---
    CBAnnulerTransfertNumChargesPontsVersAPI.Enabled = False
    CBTransfertNumChargesPontsVersAPI.Enabled = False

End Sub

Private Sub CBAnnulerTransfertNumChargesPostesVersAPI_Click()
    
    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next
    
    '--- vidage puis réaffichage des valeurs ---
    GestionModificationNumChargesPostes GESTION_GRILLES.GG_VIDAGE

    '--- gestion des boutons ---
    CBAnnulerTransfertNumChargesPostesVersAPI.Enabled = False
    CBTransfertNumChargesPostesVersAPI.Enabled = False

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

Private Sub CBTransfertNumChargesPontsVersAPI_Click()
    
    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next
    
    '--- enregistre les annexes dans l'automate ---
    If AppelFenetre(F_MESSAGE, _
                            TITRE_MESSAGES, _
                            MESSAGE_4, _
                            TYPES_MESSAGES.T_ATTENTION, _
                            TYPES_BOUTONS.T_OUI_NON, _
                            EMPLACEMENT_FOCUS.E_SUR_NON) = vbYes Then
    
        '--- enregistre les numéros de charges dans l'automate ---
        If EnregistreNumChargesPontsVersAutomate = True Then
        
            '--- vidage de la grille ---
            GestionModificationNumChargesPonts GG_VIDAGE
        
            '--- gestion des boutons ---
            CBAnnulerTransfertNumChargesPontsVersAPI.Enabled = False
            CBTransfertNumChargesPontsVersAPI.Enabled = False
        
        Else
        
            '--- lancer un message d'erreur ---
            Bidon = MessageErreur(TITRE_MESSAGES, MESSAGE_500)
        
        End If

    End If

End Sub

Private Sub CBTransfertNumChargesPostesVersAPI_Click()

    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next
    
    '--- enregistre les annexes dans l'automate ---
    If AppelFenetre(F_MESSAGE, _
                             TITRE_MESSAGES, _
                             MESSAGE_4, _
                             TYPES_MESSAGES.T_ATTENTION, _
                             TYPES_BOUTONS.T_OUI_NON, _
                             EMPLACEMENT_FOCUS.E_SUR_NON) = vbYes Then
    
        '--- enregistre les numéros de charges dans l'automate ---
        If EnregistreNumChargesPostesVersAutomate = True Then
        
            '--- vidage de la grille ---
            GestionModificationNumChargesPostes GG_VIDAGE
        
            '--- gestion des boutons ---
            CBAnnulerTransfertNumChargesPostesVersAPI.Enabled = False
            CBTransfertNumChargesPostesVersAPI.Enabled = False
        
        Else
            
            '--- lancer un message d'erreur ---
            Bidon = MessageErreur(TITRE_MESSAGES, MESSAGE_500)
        
        End If

    End If

End Sub

Private Sub CTOnglets_Click()
    
    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next
    
    '--- déclaration ---

    '--- placement du focus en fonction de l'onglet ---
    Select Case CTOnglets.CurrTab
        
        Case ONGLETS.O_GENERAL
        
        Case Else
    End Select

End Sub

Private Sub Form_Activate()
    
    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next
    
    '--- renseigne la fenêtre principale ---
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

Private Sub LTitreModificationNumChargesPonts_Click()
    On Error Resume Next
    VSFGModificationNumChargesPonts.SetFocus
End Sub

Private Sub LTitreModificationNumChargesPostes_Click()
    On Error Resume Next
    VSFGModificationNumChargesPostes.SetFocus
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
    
    '--- déclaration ---
    
    '--- calculs des emplacements ---
    With PBRenseignementsFenetre
        LRenseignementsFenetre.Left = .ScaleLeft
        LRenseignementsFenetre.Top = .ScaleTop + 30
        LRenseignementsFenetre.Width = .ScaleWidth
        LRenseignementsFenetre.Height = .ScaleHeight
    End With

End Sub

Private Sub TimerSequencesAutomatique_Timer()
    
    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next

    '--- appel de la routine --
    TimerSequencesAutomatique.Enabled = False
    GestionSequencesAutomatique
    TimerSequencesAutomatique.Enabled = True
    
    '--- bip de passage dans la routine UNIQUEMENT POUR LES TESTS ---
    'If PROGRAMME_AVEC_AUTOMATE = False Then Beep
    
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

Private Sub VSDeplacementFENETRE_Change()
    On Error Resume Next
    PBDeplacementFenetre(ZONES_DEPLACEMENT_FENETRE.Z_FILLE).Top = -VSDeplacementFenetre.Value
End Sub

'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Rôle      : Initialise la fenêtre (chargement ou en vue de la rendre visible)
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Sub InitialisationFenetre()
    
    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next

    '--- déclaration ---
    Dim a As Integer                'pour les boucles FOR...NEXT

    '--- affectation ---
  
    '--- divers sur la fenêtre ---
    With Me
        .Caption = UCase(TITRE_FENETRE)
        .WindowState = vbMaximized
    End With
    PBDeplacementFenetre(ZONES_DEPLACEMENT_FENETRE.Z_FILLE).Picture = ImgFondOrange1
    PBBoutons.Picture = ImgFondDesBoutons
    
    '--- renseignements de la fenêtre ---
    LRenseignementsFenetre.Caption = UCase(TITRE_FENETRE)
    
    '--- images de fond dans les cardes ---
    PBOnglets(0).Picture = ImgFondDeFenetreXP
    
    '--- onglet par défaut à l'ouverture ---
    CTOnglets.CurrTab = ONGLETS.O_GENERAL

 End Sub

'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Rôle      : Effectue le paramètrage de la fenêtre
' Entrées :
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Sub ParametrageFenetre()
    
    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next
    
    '--- affichage de la légende pour l'onglet général ---
    LLegendeModificationNumCharges.Caption = "0 = Effacement d'une charge, " & _
                                                                               CHARGES.C_NUM_MINI & _
                                                                              " à " & CHARGES.C_NUM_MAXI & " = Numéro d'une charge" & vbCrLf & _
                                                                              "VOTRE RESPONSABILITE EST ENGAGEE LORS DE TOUT CHANGEMENT"
    
    '--- initialisation des grilles ---
    GestionModificationNumChargesPostes GG_INITIALISATION
    GestionModificationNumChargesPonts GG_INITIALISATION
    
    '--- lancement du timer ---
    TimerSequencesAutomatique.Enabled = True
                
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
    With TimerSequencesAutomatique
        .Enabled = False
        .Interval = 0
    End With

    '--- déchargement de la fenêtre ---
    Me.Visible = False
    Unload Me
    Set OccFModeCyclique = Nothing

End Sub

'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Rôle      : Gère l'appui des touches du clavier
' Entrées :
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Private Sub GestionTouches(KeyCode As Integer, Shift As Integer)
    
    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next
    
    '--- action en fonction des touches ---
    Select Case KeyCode
        
        Case Asc("A")
            '--- active l'onglet général ---
            KeyCode = 0

        Case Asc("B")
            '--- active l'onglet du CycleActuel ---
            KeyCode = 0
        
        Case Asc("C")
            '--- active l'onglet du taqueur ---
            KeyCode = 0
        
        Case Asc("D")
            '--- active l'onglet du stockeur ---
            KeyCode = 0
        
        Case Asc("E")
            '--- active l'onglet du charreur ---
            KeyCode = 0
        
        Case Else
    End Select

End Sub

Private Sub CBQuitter_Click()
    On Error Resume Next
    DechargeFenetre
End Sub

'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Rôle      : Gére l'états des boutons après une action de l'opèrateur
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Private Sub GestionBoutons(ByVal Situation As ETATS_BOUTONS)
    
    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next

    Select Case Situation
        
        Case ETATS_BOUTONS.E_CHARGEMENT_FENETRE
            '--- au chargement de la fenêtre ---
        
        Case ETATS_BOUTONS.E_DECHARGEMENT_FENETRE
            '--- au déchargement de la fenêtre ---
        
        Case ETATS_BOUTONS.E_AVANT_VALIDER
            '--- avant valider ---
        
        Case ETATS_BOUTONS.E_APRES_VALIDER
            '--- après valider ---

        Case ETATS_BOUTONS.E_AVANT_ANNULER
            '--- avant annuler ---
        
        Case ETATS_BOUTONS.E_APRES_ANNULER
            '--- après annuler ---

        Case ETATS_BOUTONS.E_AVANT_ACTUALISER
            '--- avant actualiser ---
        
        Case ETATS_BOUTONS.E_APRES_ACTUALISER
            '--- après actualiser ---
        
        Case ETATS_BOUTONS.E_MODIFICATION_EN_COURS
            '--- après modifier (à ne pas traiter si nouvel enregistrement) ---

        Case ETATS_BOUTONS.E_AVANT_NOUVEAU
            '--- avant nouveau ---
        
        Case ETATS_BOUTONS.E_APRES_NOUVEAU
            '--- après nouveau ---
        
        Case ETATS_BOUTONS.E_AVANT_SUPPRIMER
            '--- avant supprimer ---
        
        Case ETATS_BOUTONS.E_APRES_SUPPRIMER
            '--- après supprimer ---

        Case Else
    End Select

    '--- affectation ---
    MemDernierBouton = Situation

End Sub

Private Sub VSFGModificationNumChargesPonts_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    
    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next
    
    '--- déclaration ---
    Dim ValeurCellule As Integer                  'valeur numérique de la cellule
    Dim TexteCellule  As String                    'représente le texte de la cellule

    '--- analyse des limites ---
    With VSFGModificationNumChargesPonts
    
        '--- affectation ---
        TexteCellule = .TextMatrix(Row, Col)

        If IsNumeric(TexteCellule) = True Then

            '--- affectation ---
            ValeurCellule = CInt(TexteCellule)

           Select Case Col
                   
               Case COLONNES_MODIFICATION_NUM_CHARGES_PONTS.C_NOUVEAU_NUM_CHARGE
                    '--- limites des nouveaux numéros ---
                    If ValeurCellule >= 0 And ValeurCellule <= CHARGES.C_NUM_MAXI Then
                    Else
                        .TextMatrix(Row, Col) = ""
                    End If
                
               Case Else
           End Select
                
        End If
                
    End With

End Sub

Private Sub VSFGModificationNumChargesPonts_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    
    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next

    '--- choix du masque d'édition ---
    With VSFGModificationNumChargesPonts

        Select Case .Col
                
            Case COLONNES_MODIFICATION_NUM_CHARGES_PONTS.C_NOUVEAU_NUM_CHARGE
                '--- fixer le masque d'édition ---
                .ColEditMask(Col) = "##"
     
            Case Else
        End Select
                
    End With

End Sub

Private Sub VSFGModificationNumChargesPonts_GotFocus()
    
    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next

    '--- déplacement du focus sur la grille ---
    With SFocusModificationNumChargesPonts
        .Left = LTitreModificationNumChargesPonts.Left
        .Top = LTitreModificationNumChargesPonts.Top
        .Height = ActiveControl.Height + LTitreModificationNumChargesPonts.Height + SDecorationModificationNumChargesPonts.Height - 2 * Screen.TwipsPerPixelX
        .Width = ActiveControl.Width
        .Visible = True
    End With

End Sub

Private Sub VSFGModificationNumChargesPonts_KeyPress(KeyAscii As Integer)
    On Error Resume Next
    Select Case KeyAscii
        Case vbKeyReturn               'pour changer de ligne après la touche Entrée
            With VSFGModificationNumChargesPonts
                .SetFocus
                If .Row + 1 < .Rows Then .Row = .Row + 1
                KeyAscii = 0
            End With
        Case Else
    End Select
End Sub

Private Sub VSFGModificationNumChargesPonts_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
    
    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next
    
    Select Case KeyAscii
        
        Case vbKeyReturn
            '--- pour changer de ligne après la touche Entrée ---
            With VSFGModificationNumChargesPonts
                .SetFocus
                If .Row + 1 < .Rows Then .Row = .Row + 1
            End With
        
        Case Else
            '--- nombre naturel seulement admis ---
            FiltreToucheASCII KeyAscii, DONNEES.D_NBR_NATURELS
            
            '--- gestion des boutons ---
            If KeyAscii <> 0 Then
                CBAnnulerTransfertNumChargesPontsVersAPI.Enabled = True
                CBTransfertNumChargesPontsVersAPI.Enabled = True
            End If

    End Select

End Sub

Private Sub VSFGModificationNumChargesPonts_LostFocus()
    On Error Resume Next
    SFocusModificationNumChargesPonts.Visible = False
End Sub

Private Sub VSFGModificationNumChargesPostes_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    
    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next
    
    '--- déclaration ---
    Dim ValeurCellule As Integer                  'valeur numérique de la cellule
    Dim TexteCellule  As String                    'représente le texte de la cellule

    '--- analyse des limites ---
    With VSFGModificationNumChargesPostes
    
        '--- affectation ---
        TexteCellule = .TextMatrix(Row, Col)

        If IsNumeric(TexteCellule) = True Then

            '--- affectation ---
            ValeurCellule = CInt(TexteCellule)

           Select Case Col
                   
               Case COLONNES_MODIFICATION_NUM_CHARGES_POSTES.C_NOUVEAU_NUM_CHARGE
                    '--- limites des nouveaux numéros ---
                    If ValeurCellule >= 0 And ValeurCellule <= CHARGES.C_NUM_MAXI Then
                    Else
                        .TextMatrix(Row, Col) = ""
                    End If
                
               Case Else
           End Select
                
        End If
                
    End With

End Sub

Private Sub VSFGModificationNumChargesPostes_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    
    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next

    '--- choix du masque d'édition ---
    With VSFGModificationNumChargesPostes

        Select Case .Col
                
            Case COLONNES_MODIFICATION_NUM_CHARGES_POSTES.C_NOUVEAU_NUM_CHARGE
                '--- fixer le masque d'édition ---
                .ColEditMask(Col) = "##"
     
            Case Else
        End Select
                
    End With

End Sub

Private Sub VSFGModificationNumChargesPostes_GotFocus()
    
    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next

    '--- déplacement du focus sur la grille ---
    With SFocusModificationNumChargesPostes
        .Left = LTitreModificationNumChargesPostes.Left
        .Top = LTitreModificationNumChargesPostes.Top
        .Height = ActiveControl.Height + LTitreModificationNumChargesPostes.Height + SDecorationModificationNumChargesPostes.Height - 2 * Screen.TwipsPerPixelX
        .Width = ActiveControl.Width
        .Visible = True
    End With

End Sub

Private Sub VSFGModificationNumChargesPostes_KeyPress(KeyAscii As Integer)
    On Error Resume Next
    Select Case KeyAscii
        Case vbKeyReturn               'pour changer de ligne après la touche Entrée
            With VSFGModificationNumChargesPostes
                .SetFocus
                If .Row + 1 < .Rows Then .Row = .Row + 1
                KeyAscii = 0
            End With
        Case Else
    End Select
End Sub

Private Sub VSFGModificationNumChargesPostes_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
    
    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next
    
    Select Case KeyAscii
        
        Case vbKeyReturn
            '--- pour changer de ligne après la touche Entrée ---
            With VSFGModificationNumChargesPostes
                .SetFocus
                If .Row + 1 < .Rows Then .Row = .Row + 1
            End With
        
        Case Else
            '--- nombre naturels seulement admis ---
            FiltreToucheASCII KeyAscii, DONNEES.D_NBR_NATURELS
            
            '--- gestion des boutons ---
            If KeyAscii <> 0 Then
                CBAnnulerTransfertNumChargesPostesVersAPI.Enabled = True
                CBTransfertNumChargesPostesVersAPI.Enabled = True
            End If

    End Select

End Sub

Private Sub VSFGModificationNumChargesPostes_LostFocus()
    On Error Resume Next
    SFocusModificationNumChargesPostes.Visible = False
End Sub

'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Rôle      : Gestion de la modification des numéros de charges pour les postes
' Entrées :
' Retours :
' Détails  :
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Private Sub GestionModificationNumChargesPostes(ByVal EtatSouhaite As GESTION_GRILLES)
    
    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next
    
    '--- déclaration ---
    Dim a  As Integer                               'pour les boucles FOR...NEXT
    Dim NumColonne As Long                 'représente un numéro de colonne
    Dim Texte As String                           'représente un texte quelconque

    Select Case EtatSouhaite
            
        Case GESTION_GRILLES.GG_INITIALISATION
            '--- initialisation de la grille des détails ---
            With VSFGModificationNumChargesPostes
            
                .Redraw = False
             
                .Clear
                
                .FixedRows = 1                                                                                                 'nombre de lignes fixes
                .FixedCols = 3                                                                                                   'nombre de colonnes fixes
                .Rows = DERNIER_POSTE + .FixedRows                                                              'nombre de lignes au total
                .Cols = NBR_COLONNES_MODIFICATION_NUM_CHARGES_POSTES            'nombre de colonnes au total
                .RowHeight(0) = 500                                                                                          'épaisseur des titres
                .RowHeightMin = 350                                                                                        'épaisseur mini des lignes
                .Editable = flexEDKbdMouse                                                                            'mode d'étidion des cellules

                '--- mélange des cellules identiques pour obtenir une seule cellule ---
                .MergeCells = flexMergeFixedOnly                                                    'uniquement les cellules fixes
                .MergeRow(0) = True                                                                           'mélange des cellules identiques sur la ligne

                .Row = 0                                                                                               'pointer la ligne 0
                
                '--- paramétrages de chaque colonne ---
                .Col = COLONNES_MODIFICATION_NUM_CHARGES_POSTES.C_NUM_POSTE
                .ColWidth(.Col) = 7 * EPAISSEUR_CARACTERE
                .Text = "N° du POSTE"
                .ColAlignment(.Col) = flexAlignCenterCenter
                .CellBackColor = COULEURS.BLANC
                'couleur de fond de la cellule
                
                .Col = COLONNES_MODIFICATION_NUM_CHARGES_POSTES.C_NOM_POSTE
                .ColWidth(.Col) = 7 * EPAISSEUR_CARACTERE
                .Text = "NOM du POSTE"
                .ColAlignment(.Col) = flexAlignCenterCenter
                .CellBackColor = COULEURS.BLANC                                                 'couleur de fond de la cellule
                
                .Col = COLONNES_MODIFICATION_NUM_CHARGES_POSTES.C_LIBELLE_POSTE
                .ColWidth(.Col) = 25.9 * EPAISSEUR_CARACTERE
                .Text = "LIBELLE du POSTE"
                .ColAlignment(.Col) = flexAlignCenterCenter
                .CellBackColor = COULEURS.BLANC                                                 'couleur de fond de la cellule

                .Col = COLONNES_MODIFICATION_NUM_CHARGES_POSTES.C_NUM_CHARGE
                .ColWidth(.Col) = 11 * EPAISSEUR_CARACTERE
                .Text = "N° de CHARGE ACTUEL"
                .ColAlignment(.Col) = flexAlignCenterCenter
                .CellBackColor = COULEURS.BLANC                                                 'couleur de fond de la cellule
               
                .Col = COLONNES_MODIFICATION_NUM_CHARGES_POSTES.C_NOUVEAU_NUM_CHARGE
                .ColWidth(.Col) = 11 * EPAISSEUR_CARACTERE
                .Text = "NOUVEAU" & vbCr & "N° de CHARGE"
                .ColAlignment(.Col) = flexAlignCenterCenter
                .CellBackColor = COULEURS.BLANC                                                 'couleur de fond de la cellule
                
                '--- changement de couleur pour la colonne des valeurs en cours ---
                .Select 1, _
                            COLONNES_MODIFICATION_NUM_CHARGES_POSTES.C_NUM_CHARGE, _
                            Pred(.Rows), _
                            COLONNES_MODIFICATION_NUM_CHARGES_POSTES.C_NOUVEAU_NUM_CHARGE
                .FillStyle = flexFillRepeat                                                            'force la répétition des instructions dans un groupe de cellules sélectionnées
                .CellBackColor = COULEURS.JAUNE_0
                'couleur de fond de la cellule
                .CellForeColor = COULEURS.BLEU_3                                         'couleur de premier plan de la cellule
                .FillStyle = flexFillSingle                                                             'supprime la répétition
                .Select 1, 1                                                                                   'supprime la sélection multiple

                '--- alignements des libellés ---
                For a = 1 To Pred(.Rows)
                    .Row = a
                    .Col = COLONNES_MODIFICATION_NUM_CHARGES_POSTES.C_NUM_POSTE
                    .CellAlignment = flexAlignCenterCenter                            'mode de centrage des valeurs dans la cellule
                    .Col = COLONNES_MODIFICATION_NUM_CHARGES_POSTES.C_NOM_POSTE
                    .CellAlignment = flexAlignCenterCenter                            'mode de centrage des valeurs dans la cellule
                    .Col = COLONNES_MODIFICATION_NUM_CHARGES_POSTES.C_LIBELLE_POSTE
                    .CellAlignment = flexAlignLeftCenter                                 'mode de centrage des valeurs dans la cellule
                    .Col = COLONNES_MODIFICATION_NUM_CHARGES_POSTES.C_NUM_CHARGE
                    .CellAlignment = flexAlignCenterCenter                            'mode de centrage des valeurs dans la cellule
                    .Col = COLONNES_MODIFICATION_NUM_CHARGES_POSTES.C_NOUVEAU_NUM_CHARGE
                    .CellAlignment = flexAlignCenterCenter                            'mode de centrage des valeurs dans la cellule
                Next a

                '--- remplissage des libellés ---
                For a = 1 To Pred(.Rows)
                    
                    .Row = a                                                                            'ligne en cours
                            
                    '--- affectation ---
                    .Col = COLONNES_MODIFICATION_NUM_CHARGES_POSTES.C_NUM_POSTE
                    .Text = TEtatsPostes(a).DefinitionPoste.NumPoste
                    .Col = COLONNES_MODIFICATION_NUM_CHARGES_POSTES.C_NOM_POSTE
                    .Text = TEtatsPostes(a).DefinitionPoste.NomPoste
                    .Col = COLONNES_MODIFICATION_NUM_CHARGES_POSTES.C_LIBELLE_POSTE
                    .Text = UN_ESPACE & TEtatsPostes(a).DefinitionPoste.LibellePoste
                
                Next a
                
                '--- fixer le curseur ---
                .Row = 1
                .Col = COLONNES_MODIFICATION_NUM_CHARGES_POSTES.C_NOUVEAU_NUM_CHARGE
                
                .Redraw = True
            
            End With
        
        Case GESTION_GRILLES.GG_VIDAGE
                '--- vidage de la grille ---
                With VSFGModificationNumChargesPostes
                    .Redraw = False
                    .Col = COLONNES_MODIFICATION_NUM_CHARGES_POSTES.C_NOUVEAU_NUM_CHARGE
                    For a = 1 To Pred(.Rows)
                        .TextMatrix(a, .Col) = ""
                    Next a
                    .Redraw = True
                End With
            
        Case GESTION_GRILLES.GG_AFFICHAGE
            '--- affichage ---
            With VSFGModificationNumChargesPostes
                
                '--- bloquer l'affichage de la grille ---
                .Redraw = False
                
                '--- affectation du numéro de colonne ---
                NumColonne = COLONNES_MODIFICATION_NUM_CHARGES_POSTES.C_NUM_CHARGE
                
                For a = 1 To Pred(.Rows)
                    
                    '--- affectation ---
                    Texte = CStr(TEtatsPostes(a).NumCharge)
                
                    '--- affichage en changeant la couleur des cellules ---
                    If Texte = "0" Then
                        .Cell(flexcpBackColor, a, NumColonne, a, NumColonne) = COULEURS.CYAN_0                               'couleur normale si 0
                        .TextMatrix(a, COLONNES_MODIFICATION_NUM_CHARGES_POSTES.C_NUM_CHARGE) = ""        'vider le champ de la cellule
                    Else
                        .Cell(flexcpBackColor, a, NumColonne, a, NumColonne) = COULEURS.JAUNE_2                              'changement de couleur de la cellule
                        .TextMatrix(a, COLONNES_MODIFICATION_NUM_CHARGES_POSTES.C_NUM_CHARGE) = Texte  'afficher le numéro de charge
                    End If
                
                Next a
                
                '--- rafraichissement ---
                .Redraw = True
            
            End With
        
        Case Else
    End Select
    
End Sub

'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Rôle      : Gestion de la modification des numéros de charges pour les ponts
' Entrées :
' Retours :
' Détails  :
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Private Sub GestionModificationNumChargesPonts(ByVal EtatSouhaite As GESTION_GRILLES)
    
    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next

    '--- déclaration ---
    Dim a  As Integer                               'pour les boucles FOR...NEXT
    Dim NumColonne As Long                 'représente un numéro de colonne
    Dim Texte As String                           'représente un texte quelconque

    Select Case EtatSouhaite

        Case GESTION_GRILLES.GG_INITIALISATION
            '--- initialisation de la grille des détails ---
            With VSFGModificationNumChargesPonts

                .Redraw = False

                .Clear

                .FixedRows = 1                                                                                   'nombre de lignes fixes
                .FixedCols = 3                                                                                     'nombre de colonnes fixes
                .Rows = PONTS.P_2 + .FixedRows                                                     'nombre de lignes au total
                .Cols = NBR_COLONNES_MODIFICATION_NUM_CHARGES_PONT  'nombre de colonnes au total
                .RowHeight(0) = 500                                                                            'épaisseur des titres
                .RowHeightMin = 350                                                                          'épaisseur mini des lignes
                .Editable = flexEDKbdMouse                                                              'mode d'étidion des cellules

                '--- mélange des cellules identiques pour obtenir une seule cellule ---
                .MergeCells = flexMergeFixedOnly                                                    'uniquement les cellules fixes
                .MergeRow(0) = True                                                                           'mélange des cellules identiques sur la ligne

                .Row = 0                                                                                               'pointer la ligne 0

                '--- paramétrages de chaque colonne ---
                .Col = COLONNES_MODIFICATION_NUM_CHARGES_PONTS.C_NUM_PONT
                .ColWidth(.Col) = 3 * EPAISSEUR_CARACTERE
                .Text = "LIBELLE"
                .ColAlignment(.Col) = flexAlignCenterCenter
                .CellBackColor = COULEURS.ROUGE_5                                              'couleur de fond de la cellule

                .Col = COLONNES_MODIFICATION_NUM_CHARGES_PONTS.C_LIBELLE_PONT
                .ColWidth(.Col) = 25.8 * EPAISSEUR_CARACTERE
                .Text = "LIBELLE"
                .ColAlignment(.Col) = flexAlignCenterCenter
                .CellBackColor = COULEURS.ROUGE_5                                              'couleur de fond de la cellule

                .Col = COLONNES_MODIFICATION_NUM_CHARGES_PONTS.C_NUM_CHARGE
                .ColWidth(.Col) = 11 * EPAISSEUR_CARACTERE
                .Text = "N° de CHARGE ACTUEL"
                .ColAlignment(.Col) = flexAlignCenterCenter
                .CellBackColor = COULEURS.ROUGE_5                                              'couleur de fond de la cellule

                .Col = COLONNES_MODIFICATION_NUM_CHARGES_PONTS.C_NOUVEAU_NUM_CHARGE
                .ColWidth(.Col) = 11 * EPAISSEUR_CARACTERE
                .Text = "NOUVEAU" & vbCr & "N° de CHARGE"
                .ColAlignment(.Col) = flexAlignCenterCenter
                .CellBackColor = COULEURS.ROUGE_5                                              'couleur de fond de la cellule

                '--- changement de couleur pour la colonne des valeurs en cours ---
                .Select 1, _
                            COLONNES_MODIFICATION_NUM_CHARGES_PONTS.C_NUM_CHARGE, _
                            Pred(.Rows), _
                            COLONNES_MODIFICATION_NUM_CHARGES_PONTS.C_NOUVEAU_NUM_CHARGE
                .FillStyle = flexFillRepeat                                                            'force la répétition des instructions dans un groupe de cellules sélectionnées
                .CellBackColor = COULEURS.CYAN_0                                        'couleur de fond de la cellule
                .CellForeColor = COULEURS.BLEU_3                                         'couleur de premier plan de la cellule
                .FillStyle = flexFillSingle                                                             'supprime la répétition
                .Select 1, 1                                                                                   'supprime la sélection multiple

                '--- alignements des libellés ---
                For a = 1 To Pred(.Rows)
                    .Row = a
                    .Col = COLONNES_MODIFICATION_NUM_CHARGES_PONTS.C_NUM_PONT
                    .CellAlignment = flexAlignCenterCenter                            'mode de centrage des valeurs dans la cellule
                    .Col = COLONNES_MODIFICATION_NUM_CHARGES_PONTS.C_LIBELLE_PONT
                    .CellAlignment = flexAlignLeftCenter                                 'mode de centrage des valeurs dans la cellule
                    .Col = COLONNES_MODIFICATION_NUM_CHARGES_PONTS.C_NUM_CHARGE
                    .CellAlignment = flexAlignCenterCenter                            'mode de centrage des valeurs dans la cellule
                    .Col = COLONNES_MODIFICATION_NUM_CHARGES_PONTS.C_NOUVEAU_NUM_CHARGE
                    .CellAlignment = flexAlignCenterCenter                            'mode de centrage des valeurs dans la cellule
                Next a

                '--- remplissage des libellés ---
                For a = 1 To Pred(.Rows)

                    .Row = a                                                                                              'ligne en cours

                    '--- pour le pont ---
                    Texte = "Pont " & a
                    .Col = COLONNES_MODIFICATION_NUM_CHARGES_PONTS.C_NUM_PONT
                    .Text = Texte
                    .Col = COLONNES_MODIFICATION_NUM_CHARGES_PONTS.C_LIBELLE_PONT
                    .CellAlignment = flexAlignCenterCenter                                              'mode de centrage des valeurs dans la cellule
                    .Text = Texte
                    .MergeRow(a) = True                                                                           'mélange des cellules identiques sur la ligne

                Next a

                '--- fixer le curseur ---
                .Row = 1
                .Col = COLONNES_MODIFICATION_NUM_CHARGES_PONTS.C_NOUVEAU_NUM_CHARGE

                .Redraw = True

            End With

        Case GESTION_GRILLES.GG_VIDAGE
                '--- vidage de la grille ---
                With VSFGModificationNumChargesPonts
                    .Redraw = False
                    .Col = COLONNES_MODIFICATION_NUM_CHARGES_PONTS.C_NOUVEAU_NUM_CHARGE
                    For a = 1 To Pred(.Rows)
                        .TextMatrix(a, .Col) = ""
                    Next a
                    .Redraw = True
                End With

        Case GESTION_GRILLES.GG_AFFICHAGE
            '--- affichage ---
            With VSFGModificationNumChargesPonts

                '--- bloquer l'affichage de la grille ---
                .Redraw = False

                '--- affectation du numéro de colonne ---
                NumColonne = COLONNES_MODIFICATION_NUM_CHARGES_PONTS.C_NUM_CHARGE

                For a = 1 To Pred(.Rows)

                    '--- affectation ---
                    Texte = CStr(TEtatsPonts(a).NumCharge)

                    '--- affichage en changeant la couleur des cellules ---
                    If Texte = "0" Then
                        .Cell(flexcpBackColor, a, NumColonne, a, NumColonne) = COULEURS.CYAN_0                                 'couleur normale si 0
                        .TextMatrix(a, COLONNES_MODIFICATION_NUM_CHARGES_PONTS.C_NUM_CHARGE) = ""            'vider le champ de la cellule
                    Else
                        .Cell(flexcpBackColor, a, NumColonne, a, NumColonne) = COULEURS.JAUNE_2                               'changement de couleur de la cellule
                        .TextMatrix(a, COLONNES_MODIFICATION_NUM_CHARGES_PONTS.C_NUM_CHARGE) = Texte      'afficher le numéro de charge
                    End If

                Next a

                '--- rafraichissement ---
                .Redraw = True

            End With

        Case Else
    End Select
    
End Sub

'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Rôle      : Enregistre les numéros de charge pour les postes dans l'automate
' Entrées :
' Retours : EnregistreNumChargesVersAutomate -> FALSE = Incident de communication
'                                                                                   TRUE = Aucun incident
' Détails  :
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Private Function EnregistreNumChargesPostesVersAutomate() As Boolean
    
    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next

    '--- constantes privées ---
    Const NOM_GROUPE = "SUIVI_LIGNE"    'nom du groupe
    
    '--- déclaration ---
    Dim a  As Integer                                       'pour les boucles FOR...NEXT
    Dim ValeurRetourneeAPI As Long             'valeur retournée par une fonction concernant le dialogue avec l'automate
    Dim NomVariable As String                       'nom de la variable OPC
    Dim NouveauNumCharge  As String         'représente un nouveau numéro de charge
        
    '--- valeur par défaut ---
    EnregistreNumChargesPostesVersAutomate = True
   
    '--- transfert des valeurs ---
    If PROGRAMME_AVEC_AUTOMATE = True Then
    
        '--- curseur de la souris ---
        SourisEnAttente True
        
        '--- transfert des valeurs ---
        With VSFGModificationNumChargesPostes
                
            For a = 1 To Pred(.Rows)
                        
                '--- extraction de la cellule ---
                NouveauNumCharge = .TextMatrix(a, COLONNES_MODIFICATION_NUM_CHARGES_POSTES.C_NOUVEAU_NUM_CHARGE)
                NouveauNumCharge = Trim(NouveauNumCharge)
                                        
                                       
                '--- affectation du nom de la variable ---
                NomVariable = "NumChargePoste" & Right("00" & CStr(a), 2)
                    
                '--- écriture dans l'automate ---
                If NouveauNumCharge <> "" Then
                    ValeurRetourneeAPI = APIEcritureVariableNommee(NOM_GROUPE, NomVariable, NouveauNumCharge)
                    If ValeurRetourneeAPI <> 0 Then
                        EnregistreNumChargesPostesVersAutomate = False
                        Exit For
                    End If
                End If
                    
           Next a

        End With
        
        '--- curseur de la souris ---
        SourisEnAttente False

    End If
        
End Function

'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Rôle      : Enregistre les numéros de charge pour les ponts dans l'automate
' Entrées :
' Retours : EnregistreNumChargesVersAutomate -> FALSE = Incident de communication
'                                                                                   TRUE = Aucun incident
' Détails  :
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Private Function EnregistreNumChargesPontsVersAutomate() As Boolean
    
    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next

    '--- constantes privées ---
    Const NOM_GROUPE = "SUIVI_LIGNE"    'nom du groupe
    
    '--- déclaration ---
    Dim a  As Integer                                       'pour les boucles FOR...NEXT
    Dim ValeurRetourneeAPI As Long             'valeur retournée par une fonction concernant le dialogue avec l'automate
    Dim NomGroupe As String                         'représente un nom de groupe
    Dim NomVariable As String                       'nom de la variable OPC
    Dim NouveauNumCharge  As String         'représente un nouveau numéro de charge
        
    '--- valeur par défaut ---
    EnregistreNumChargesPontsVersAutomate = True

    '--- transfert des valeurs ---
    If PROGRAMME_AVEC_AUTOMATE = True Then
    
        '--- curseur de la souris ---
        SourisEnAttente True
        
        '--- transfert des valeurs ---
        With VSFGModificationNumChargesPonts
                
            For a = 1 To Pred(.Rows)
                        
                '--- extraction de la cellule ---
                NouveauNumCharge = .TextMatrix(a, COLONNES_MODIFICATION_NUM_CHARGES_PONTS.C_NOUVEAU_NUM_CHARGE)
                NouveauNumCharge = Trim(NouveauNumCharge)
                        
                '--- affectation du nom du groupe et de la variable ---
                NomVariable = "NumChargeP" & a
                   
                '--- écriture dans l'automate ---
                If NouveauNumCharge <> "" Then
                    ValeurRetourneeAPI = APIEcritureVariableNommee(NOM_GROUPE, NomVariable, NouveauNumCharge)
                    If ValeurRetourneeAPI <> 0 Then
                        EnregistreNumChargesPontsVersAutomate = False
                        Exit For
                    End If
                End If
                    
           Next a

        End With
        
        '--- curseur de la souris ---
        SourisEnAttente False

    End If
        
End Function

Private Sub GestionSequencesAutomatique()
    
    '--- gestion des erreurs ---
    On Error Resume Next
    
    '--- constantes privées ---
    
    '--- déclaration ---

    '**************************************************************** GENERAL ********************************************************************
    
    GestionModificationNumChargesPostes GG_AFFICHAGE
    GestionModificationNumChargesPonts GG_AFFICHAGE
    
End Sub


VERSION 5.00
Object = "{0AFE7BE0-11B7-4A3E-978D-D4501E9A57FE}#1.0#0"; "c1sizer.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form FMaintenance 
   ClientHeight    =   14595
   ClientLeft      =   2850
   ClientTop       =   1035
   ClientWidth     =   23175
   Icon            =   "FMaintenance.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   0
   ScaleWidth      =   0
   ShowInTaskbar   =   0   'False
   WindowState     =   2  'Maximized
   Begin VB.PictureBox PBDeplacementFenetre 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   12975
      Index           =   0
      Left            =   0
      ScaleHeight     =   12975
      ScaleWidth      =   23175
      TabIndex        =   3
      Top             =   375
      Width           =   23175
      Begin VB.PictureBox PBDeplacementFenetre 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   12735
         Index           =   1
         Left            =   0
         ScaleHeight     =   12735
         ScaleWidth      =   28740
         TabIndex        =   4
         Top             =   0
         Width           =   28740
         Begin C1SizerLibCtl.C1Tab CTOnglets 
            Height          =   12195
            Left            =   300
            TabIndex        =   14
            Top             =   300
            Width           =   28155
            _cx             =   49662
            _cy             =   21511
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
            Caption         =   "Touche A = VARIATEURS|Touche B = ENTREES / SORTIES|Touche C = LECTURE / ECRITURE|Touche D = MNEMONIQUES AUTOMATE"
            Align           =   0
            CurrTab         =   0
            FirstTab        =   0
            Style           =   1
            Position        =   0
            AutoSwitch      =   -1  'True
            AutoScroll      =   -1  'True
            TabPreview      =   -1  'True
            ShowFocusRect   =   0   'False
            TabsPerPage     =   4
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
               BackColor       =   &H00C0FFC0&
               Height          =   11685
               Index           =   1
               Left            =   28800
               ScaleHeight     =   11625
               ScaleWidth      =   28005
               TabIndex        =   20
               Top             =   465
               Width           =   28065
               Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFGEntreesSortiesAutomate 
                  Height          =   11235
                  Left            =   180
                  TabIndex        =   21
                  Top             =   180
                  Width           =   27675
                  _ExtentX        =   48816
                  _ExtentY        =   19817
                  _Version        =   393216
                  BackColor       =   16777215
                  ForeColor       =   12582912
                  Rows            =   100
                  Cols            =   6
                  BackColorFixed  =   8421376
                  ForeColorFixed  =   16777215
                  BackColorSel    =   16777215
                  BackColorBkg    =   12648447
                  GridColor       =   8421504
                  GridColorFixed  =   0
                  GridColorUnpopulated=   -2147483644
                  WordWrap        =   -1  'True
                  AllowBigSelection=   0   'False
                  FocusRect       =   0
                  HighLight       =   0
                  AllowUserResizing=   3
                  Appearance      =   0
                  BandDisplay     =   1
                  RowSizingMode   =   1
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  BeginProperty FontFixed {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  _NumberOfBands  =   1
                  _Band(0).Cols   =   6
                  _Band(0).GridLinesBand=   1
                  _Band(0).TextStyleBand=   0
                  _Band(0).TextStyleHeader=   0
               End
               Begin VB.Shape SFocusEntreesSortiesAutomate 
                  BorderColor     =   &H000000FF&
                  BorderWidth     =   4
                  Height          =   300
                  Left            =   0
                  Top             =   0
                  Visible         =   0   'False
                  Width           =   405
               End
            End
            Begin VB.PictureBox PBOnglets 
               BackColor       =   &H00FFC0C0&
               Height          =   11685
               Index           =   2
               Left            =   29100
               ScaleHeight     =   11625
               ScaleWidth      =   28005
               TabIndex        =   17
               Top             =   465
               Width           =   28065
               Begin VB.TextBox TBEditionLectureEcritureAutomate 
                  BackColor       =   &H0000FFFF&
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
                  Height          =   255
                  Left            =   480
                  TabIndex        =   18
                  Top             =   360
                  Visible         =   0   'False
                  Width           =   1035
               End
               Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFGLectureEcritureAutomate 
                  Height          =   11235
                  Left            =   180
                  TabIndex        =   19
                  Top             =   180
                  Width           =   27675
                  _ExtentX        =   48816
                  _ExtentY        =   19817
                  _Version        =   393216
                  BackColor       =   16777215
                  ForeColor       =   12582912
                  Rows            =   100
                  Cols            =   6
                  BackColorFixed  =   12582912
                  ForeColorFixed  =   16777215
                  BackColorSel    =   16777215
                  BackColorBkg    =   12648447
                  GridColor       =   8421504
                  GridColorFixed  =   0
                  GridColorUnpopulated=   -2147483644
                  WordWrap        =   -1  'True
                  AllowBigSelection=   0   'False
                  FocusRect       =   0
                  HighLight       =   0
                  AllowUserResizing=   3
                  Appearance      =   0
                  BandDisplay     =   1
                  RowSizingMode   =   1
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  BeginProperty FontFixed {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  _NumberOfBands  =   1
                  _Band(0).Cols   =   6
                  _Band(0).GridLinesBand=   1
                  _Band(0).TextStyleBand=   0
                  _Band(0).TextStyleHeader=   0
               End
               Begin VB.Shape SFocusLectureEcritureAutomate 
                  BorderColor     =   &H000000FF&
                  BorderWidth     =   4
                  Height          =   300
                  Left            =   120
                  Top             =   120
                  Visible         =   0   'False
                  Width           =   405
               End
            End
            Begin VB.PictureBox PBOnglets 
               BackColor       =   &H00C0FFFF&
               Height          =   11685
               Index           =   3
               Left            =   29400
               ScaleHeight     =   11625
               ScaleWidth      =   28005
               TabIndex        =   15
               Top             =   465
               Width           =   28065
               Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFGMnemoniquesAutomate 
                  Height          =   11235
                  Left            =   180
                  TabIndex        =   16
                  Top             =   180
                  Width           =   27675
                  _ExtentX        =   48816
                  _ExtentY        =   19817
                  _Version        =   393216
                  BackColor       =   16777215
                  ForeColor       =   12582912
                  Rows            =   100
                  Cols            =   6
                  BackColorFixed  =   192
                  ForeColorFixed  =   16777215
                  BackColorSel    =   16777215
                  BackColorBkg    =   12648447
                  GridColor       =   8421504
                  GridColorFixed  =   0
                  GridColorUnpopulated=   -2147483644
                  WordWrap        =   -1  'True
                  AllowBigSelection=   0   'False
                  FocusRect       =   0
                  HighLight       =   0
                  AllowUserResizing=   3
                  Appearance      =   0
                  BandDisplay     =   1
                  RowSizingMode   =   1
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  BeginProperty FontFixed {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  _NumberOfBands  =   1
                  _Band(0).Cols   =   6
                  _Band(0).GridLinesBand=   1
                  _Band(0).TextStyleBand=   0
                  _Band(0).TextStyleHeader=   0
               End
               Begin VB.Shape SFocusMnemoniquesAutomate 
                  BorderColor     =   &H000000FF&
                  BorderWidth     =   4
                  Height          =   300
                  Left            =   60
                  Top             =   120
                  Visible         =   0   'False
                  Width           =   405
               End
            End
            Begin VB.PictureBox PBOnglets 
               BackColor       =   &H00C0E0FF&
               Height          =   11685
               Index           =   0
               Left            =   45
               ScaleHeight     =   11625
               ScaleWidth      =   28005
               TabIndex        =   22
               Top             =   465
               Width           =   28065
               Begin VB.CommandButton CBInformationsDefautVariateurs 
                  BackColor       =   &H00C0C0FF&
                  Caption         =   "Cliquez ici pour obtenir des informations détaillés sur les  défauts"
                  BeginProperty Font 
                     Name            =   "Arial"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   1215
                  Left            =   9600
                  Style           =   1  'Graphical
                  TabIndex        =   35
                  Top             =   5580
                  Width           =   1455
               End
               Begin VB.OptionButton OBNomsVariateurs 
                  BackColor       =   &H00808000&
                  Caption         =   "Translation DROITE"
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
                  Height          =   375
                  Index           =   11
                  Left            =   6960
                  Style           =   1  'Graphical
                  TabIndex        =   34
                  Top             =   1020
                  Width           =   2655
               End
               Begin VB.OptionButton OBNomsVariateurs 
                  BackColor       =   &H00808000&
                  Caption         =   "Translation GAUCHE"
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
                  Height          =   375
                  Index           =   10
                  Left            =   6960
                  Style           =   1  'Graphical
                  TabIndex        =   33
                  Top             =   600
                  Width           =   2655
               End
               Begin VB.OptionButton OBNomsVariateurs 
                  BackColor       =   &H00808000&
                  Caption         =   "Levage"
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
                  Height          =   375
                  Index           =   2
                  Left            =   360
                  Style           =   1  'Graphical
                  TabIndex        =   32
                  Top             =   1440
                  Width           =   2655
               End
               Begin VB.OptionButton OBNomsVariateurs 
                  Appearance      =   0  'Flat
                  BackColor       =   &H00808000&
                  Caption         =   "Translation DROITE"
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
                  Height          =   375
                  Index           =   1
                  Left            =   360
                  Style           =   1  'Graphical
                  TabIndex        =   31
                  Top             =   1020
                  Width           =   2655
               End
               Begin VB.OptionButton OBNomsVariateurs 
                  Appearance      =   0  'Flat
                  BackColor       =   &H00808000&
                  Caption         =   "Translation GAUCHE"
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
                  Height          =   375
                  Index           =   0
                  Left            =   360
                  MaskColor       =   &H00FF00FF&
                  Style           =   1  'Graphical
                  TabIndex        =   30
                  Top             =   600
                  UseMaskColor    =   -1  'True
                  Width           =   2655
               End
               Begin VB.OptionButton OBNomsVariateurs 
                  BackColor       =   &H00808000&
                  Caption         =   "Bac anti égouttures GAUCHE"
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
                  Height          =   375
                  Index           =   3
                  Left            =   360
                  Style           =   1  'Graphical
                  TabIndex        =   29
                  Top             =   1860
                  Width           =   2655
               End
               Begin VB.OptionButton OBNomsVariateurs 
                  BackColor       =   &H00808000&
                  Caption         =   "Bac anti égouttures DROIT"
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
                  Height          =   375
                  Index           =   4
                  Left            =   360
                  Style           =   1  'Graphical
                  TabIndex        =   28
                  Top             =   2280
                  Width           =   2655
               End
               Begin VB.OptionButton OBNomsVariateurs 
                  Appearance      =   0  'Flat
                  BackColor       =   &H00808000&
                  Caption         =   "Translation GAUCHE"
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
                  Height          =   375
                  Index           =   5
                  Left            =   3660
                  MaskColor       =   &H00FF00FF&
                  Style           =   1  'Graphical
                  TabIndex        =   27
                  Top             =   600
                  UseMaskColor    =   -1  'True
                  Width           =   2655
               End
               Begin VB.OptionButton OBNomsVariateurs 
                  Appearance      =   0  'Flat
                  BackColor       =   &H00808000&
                  Caption         =   "Translation DROITE"
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
                  Height          =   375
                  Index           =   6
                  Left            =   3660
                  Style           =   1  'Graphical
                  TabIndex        =   26
                  Top             =   1020
                  Width           =   2655
               End
               Begin VB.OptionButton OBNomsVariateurs 
                  BackColor       =   &H00808000&
                  Caption         =   "Levage"
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
                  Height          =   375
                  Index           =   7
                  Left            =   3660
                  Style           =   1  'Graphical
                  TabIndex        =   25
                  Top             =   1440
                  Width           =   2655
               End
               Begin VB.OptionButton OBNomsVariateurs 
                  BackColor       =   &H00808000&
                  Caption         =   "Bac anti égouttures GAUCHE"
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
                  Height          =   375
                  Index           =   8
                  Left            =   3660
                  Style           =   1  'Graphical
                  TabIndex        =   24
                  Top             =   1860
                  Width           =   2655
               End
               Begin VB.OptionButton OBNomsVariateurs 
                  BackColor       =   &H00808000&
                  Caption         =   "Bac anti égouttures DROIT"
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
                  Height          =   375
                  Index           =   9
                  Left            =   3660
                  Style           =   1  'Graphical
                  TabIndex        =   23
                  Top             =   2280
                  Width           =   2655
               End
               Begin VB.Image ImgFdCLogiciels 
                  BorderStyle     =   1  'Fixed Single
                  Height          =   315
                  Left            =   1920
                  Picture         =   "FMaintenance.frx":014A
                  Top             =   4320
                  Width           =   315
               End
               Begin VB.Image ImgVerVarLiberation 
                  BorderStyle     =   1  'Fixed Single
                  Height          =   315
                  Left            =   6540
                  Picture         =   "FMaintenance.frx":0500
                  Top             =   4320
                  Width           =   315
               End
               Begin VB.Image ImgMA 
                  BorderStyle     =   1  'Fixed Single
                  Height          =   315
                  Left            =   5940
                  Picture         =   "FMaintenance.frx":08B6
                  Top             =   4320
                  Width           =   315
               End
               Begin VB.Image ImgMaintienPosition 
                  BorderStyle     =   1  'Fixed Single
                  Height          =   315
                  Left            =   5640
                  Picture         =   "FMaintenance.frx":0C6C
                  Top             =   4320
                  Width           =   315
               End
               Begin VB.Image ImgCommutationRampes 
                  BorderStyle     =   1  'Fixed Single
                  Height          =   315
                  Left            =   5340
                  Picture         =   "FMaintenance.frx":1022
                  Top             =   4320
                  Width           =   315
               End
               Begin VB.Image ImgCommutationJeuParametres 
                  BorderStyle     =   1  'Fixed Single
                  Height          =   315
                  Left            =   5040
                  Picture         =   "FMaintenance.frx":13D8
                  Top             =   4320
                  Width           =   315
               End
               Begin VB.Image ImgReset 
                  BorderStyle     =   1  'Fixed Single
                  Height          =   315
                  Left            =   4740
                  Picture         =   "FMaintenance.frx":178E
                  Top             =   4320
                  Width           =   315
               End
               Begin VB.Image ImgReservePFa 
                  BorderStyle     =   1  'Fixed Single
                  Height          =   315
                  Left            =   4440
                  Picture         =   "FMaintenance.frx":1B44
                  Top             =   4320
                  Width           =   315
               End
               Begin VB.Image ImgMAR 
                  BorderStyle     =   1  'Fixed Single
                  Height          =   315
                  Left            =   6240
                  Picture         =   "FMaintenance.frx":1EFA
                  Top             =   4320
                  Width           =   315
               End
               Begin VB.Image ImgStart 
                  BorderStyle     =   1  'Fixed Single
                  Height          =   315
                  Left            =   4020
                  Picture         =   "FMaintenance.frx":22B0
                  Top             =   4320
                  Width           =   315
               End
               Begin VB.Image ImgJoggMoins 
                  BorderStyle     =   1  'Fixed Single
                  Height          =   315
                  Left            =   3720
                  Picture         =   "FMaintenance.frx":2666
                  Top             =   4320
                  Width           =   315
               End
               Begin VB.Image ImgJoggPlus 
                  BorderStyle     =   1  'Fixed Single
                  Height          =   315
                  Left            =   3420
                  Picture         =   "FMaintenance.frx":2A1C
                  Top             =   4320
                  Width           =   315
               End
               Begin VB.Image ImgSelectionModeL 
                  BorderStyle     =   1  'Fixed Single
                  Height          =   315
                  Left            =   3120
                  Picture         =   "FMaintenance.frx":2DD2
                  Top             =   4320
                  Width           =   315
               End
               Begin VB.Image ImgSelectionModeH 
                  BorderStyle     =   1  'Fixed Single
                  Height          =   315
                  Left            =   2820
                  Picture         =   "FMaintenance.frx":3188
                  Top             =   4320
                  Width           =   315
               End
               Begin VB.Image ImgReservePFo 
                  BorderStyle     =   1  'Fixed Single
                  Height          =   315
                  Left            =   2520
                  Picture         =   "FMaintenance.frx":353E
                  Top             =   4320
                  Width           =   315
               End
               Begin VB.Image ImgCommutation 
                  BorderStyle     =   1  'Fixed Single
                  Height          =   315
                  Left            =   2220
                  Picture         =   "FMaintenance.frx":38F4
                  Top             =   4320
                  Width           =   315
               End
               Begin VB.Label LMotCommandeBit8 
                  Appearance      =   0  'Flat
                  BackColor       =   &H80000005&
                  BorderStyle     =   1  'Fixed Single
                  Caption         =   " 8 : Start"
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
                  Height          =   255
                  Left            =   660
                  TabIndex        =   83
                  Top             =   7080
                  Width           =   3615
               End
               Begin VB.Label LMotCommandeBit9 
                  Appearance      =   0  'Flat
                  BackColor       =   &H80000005&
                  BorderStyle     =   1  'Fixed Single
                  Caption         =   " 9 : Jogg -"
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
                  Height          =   255
                  Left            =   660
                  TabIndex        =   84
                  Top             =   6780
                  Width           =   3375
               End
               Begin VB.Label LMotCommandeBit10 
                  Appearance      =   0  'Flat
                  BackColor       =   &H80000005&
                  BorderStyle     =   1  'Fixed Single
                  Caption         =   " 10 : Jogg +"
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
                  Height          =   255
                  Left            =   660
                  TabIndex        =   68
                  Top             =   6480
                  Width           =   3075
               End
               Begin VB.Label LMotCommandeBit11 
                  Appearance      =   0  'Flat
                  BackColor       =   &H80000005&
                  BorderStyle     =   1  'Fixed Single
                  Caption         =   " 11 : Sélection Mode Low"
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
                  Height          =   255
                  Left            =   660
                  TabIndex        =   85
                  Top             =   6180
                  Width           =   2775
               End
               Begin VB.Label LMotCommandeBit12 
                  Appearance      =   0  'Flat
                  BackColor       =   &H80000005&
                  BorderStyle     =   1  'Fixed Single
                  Caption         =   " 12 : Sélection Mode High"
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
                  Height          =   255
                  Left            =   660
                  TabIndex        =   86
                  Top             =   5880
                  Width           =   2475
               End
               Begin VB.Label LMotCommandeBit13 
                  Appearance      =   0  'Flat
                  BackColor       =   &H80000005&
                  BorderStyle     =   1  'Fixed Single
                  Caption         =   " 13 : Réservé"
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
                  Height          =   255
                  Left            =   660
                  TabIndex        =   87
                  Top             =   5580
                  Width           =   2175
               End
               Begin VB.Label LMotCommandeBit14 
                  Appearance      =   0  'Flat
                  BackColor       =   &H80000005&
                  BorderStyle     =   1  'Fixed Single
                  Caption         =   " 14 : Commutation rampes"
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
                  Height          =   255
                  Left            =   660
                  TabIndex        =   88
                  Top             =   5280
                  Width           =   1875
               End
               Begin VB.Label LMotCommandeBit15 
                  Appearance      =   0  'Flat
                  BackColor       =   &H80000005&
                  BorderStyle     =   1  'Fixed Single
                  Caption         =   " 15 : /FdC logiciels"
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
                  Height          =   255
                  Left            =   660
                  TabIndex        =   89
                  Top             =   4980
                  Width           =   1575
               End
               Begin VB.Label LMotCommandeBit7 
                  Appearance      =   0  'Flat
                  BackColor       =   &H80000005&
                  BorderStyle     =   1  'Fixed Single
                  Caption         =   " 7 : Réservé"
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
                  Height          =   255
                  Left            =   4440
                  TabIndex        =   75
                  Top             =   7080
                  Width           =   4455
               End
               Begin VB.Label LMotCommandeBit6 
                  Appearance      =   0  'Flat
                  BackColor       =   &H80000005&
                  BorderStyle     =   1  'Fixed Single
                  Caption         =   " 6 : Reset"
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
                  Height          =   255
                  Left            =   4740
                  TabIndex        =   74
                  Top             =   6780
                  Width           =   4155
               End
               Begin VB.Label LMotCommandeBit5 
                  Appearance      =   0  'Flat
                  BackColor       =   &H80000005&
                  BorderStyle     =   1  'Fixed Single
                  Caption         =   " 5 : Commutation jeu paramètres"
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
                  Height          =   255
                  Left            =   5040
                  TabIndex        =   73
                  Top             =   6480
                  Width           =   3855
               End
               Begin VB.Label LMotCommandeBit4 
                  Appearance      =   0  'Flat
                  BackColor       =   &H80000005&
                  BorderStyle     =   1  'Fixed Single
                  Caption         =   " 4 : Commutation rampes "
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
                  Height          =   255
                  Left            =   5340
                  TabIndex        =   72
                  Top             =   6180
                  Width           =   3555
               End
               Begin VB.Label LMotCommandeBit3 
                  Appearance      =   0  'Flat
                  BackColor       =   &H80000005&
                  BorderStyle     =   1  'Fixed Single
                  Caption         =   " 3 : Maintien de position"
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
                  Height          =   255
                  Left            =   5640
                  TabIndex        =   71
                  Top             =   5880
                  Width           =   3255
               End
               Begin VB.Label LMotCommandeBit2 
                  Appearance      =   0  'Flat
                  BackColor       =   &H80000005&
                  BorderStyle     =   1  'Fixed Single
                  Caption         =   " 2 : Marche / Arrêt"
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
                  Height          =   255
                  Left            =   5940
                  TabIndex        =   93
                  Top             =   5580
                  Width           =   2955
               End
               Begin VB.Label LMotCommandeBit1 
                  Appearance      =   0  'Flat
                  BackColor       =   &H80000005&
                  BorderStyle     =   1  'Fixed Single
                  Caption         =   " 1 : Marche / Arrêt rapide"
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
                  Height          =   255
                  Left            =   6240
                  TabIndex        =   92
                  Top             =   5280
                  Width           =   2655
               End
               Begin VB.Label LMotCommandeBit0 
                  Appearance      =   0  'Flat
                  BackColor       =   &H80000005&
                  BorderStyle     =   1  'Fixed Single
                  Caption         =   " 0 : Verrouillage / Libération"
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
                  Height          =   255
                  Left            =   6540
                  TabIndex        =   91
                  Top             =   4980
                  Width           =   2355
               End
               Begin VB.Label LMotEtatBit7 
                  Appearance      =   0  'Flat
                  BackColor       =   &H80000005&
                  BorderStyle     =   1  'Fixed Single
                  Caption         =   " 7 : Fin de course à GAUCHE"
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
                  Height          =   255
                  Left            =   11100
                  TabIndex        =   82
                  Top             =   7080
                  Width           =   3675
               End
               Begin VB.Label LMotEtatBit6 
                  Appearance      =   0  'Flat
                  BackColor       =   &H80000005&
                  BorderStyle     =   1  'Fixed Single
                  Caption         =   " 6 : Fin de course à DROITE"
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
                  Height          =   255
                  Left            =   11400
                  TabIndex        =   81
                  Top             =   6780
                  Width           =   3375
               End
               Begin VB.Label LMotEtatBit5 
                  Appearance      =   0  'Flat
                  BackColor       =   &H80000005&
                  BorderStyle     =   1  'Fixed Single
                  Caption         =   " 5 : Défaut / Avertissement"
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
                  Height          =   255
                  Left            =   11700
                  TabIndex        =   80
                  Top             =   6480
                  Width           =   3075
               End
               Begin VB.Label LMotEtatBit4 
                  Appearance      =   0  'Flat
                  BackColor       =   &H80000005&
                  BorderStyle     =   1  'Fixed Single
                  Caption         =   " 4 : Frein libéré"
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
                  Height          =   255
                  Left            =   12000
                  TabIndex        =   79
                  Top             =   6180
                  Width           =   2775
               End
               Begin VB.Label LMotEtatBit3 
                  Appearance      =   0  'Flat
                  BackColor       =   &H80000005&
                  BorderStyle     =   1  'Fixed Single
                  Caption         =   " 3 : Position atteinte"
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
                  Height          =   255
                  Left            =   12300
                  TabIndex        =   78
                  Top             =   5880
                  Width           =   2475
               End
               Begin VB.Label LMotEtatBit2 
                  Appearance      =   0  'Flat
                  BackColor       =   &H80000005&
                  BorderStyle     =   1  'Fixed Single
                  Caption         =   " 2 : Axe référencé"
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
                  Height          =   255
                  Left            =   12600
                  TabIndex        =   90
                  Top             =   5580
                  Width           =   2175
               End
               Begin VB.Label LMotEtatBit1 
                  Appearance      =   0  'Flat
                  BackColor       =   &H80000005&
                  BorderStyle     =   1  'Fixed Single
                  Caption         =   " 1 : Variateur prêt"
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
                  Height          =   255
                  Left            =   12900
                  TabIndex        =   77
                  Top             =   5280
                  Width           =   1875
               End
               Begin VB.Label LMotEtatBit0 
                  Appearance      =   0  'Flat
                  BackColor       =   &H80000005&
                  BorderStyle     =   1  'Fixed Single
                  Caption         =   " 0 : Moteur tourne"
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
                  Height          =   255
                  Left            =   13200
                  TabIndex        =   76
                  Top             =   4980
                  Width           =   1575
               End
               Begin VB.Label LNumDefaut 
                  Alignment       =   2  'Center
                  Appearance      =   0  'Flat
                  BackColor       =   &H00FFFFC0&
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
                  Height          =   255
                  Left            =   10320
                  TabIndex        =   94
                  Top             =   4410
                  Width           =   735
               End
               Begin VB.Image ImgFdCGauche 
                  BorderStyle     =   1  'Fixed Single
                  Height          =   315
                  Left            =   11100
                  Picture         =   "FMaintenance.frx":3CAA
                  Top             =   4380
                  Width           =   315
               End
               Begin VB.Image ImgFdCDroite 
                  BorderStyle     =   1  'Fixed Single
                  Height          =   315
                  Left            =   11400
                  Picture         =   "FMaintenance.frx":4060
                  Top             =   4380
                  Width           =   315
               End
               Begin VB.Image ImgDefautAvertissement 
                  BorderStyle     =   1  'Fixed Single
                  Height          =   315
                  Left            =   11700
                  Picture         =   "FMaintenance.frx":4416
                  Top             =   4380
                  Width           =   315
               End
               Begin VB.Image ImgFreinLibere 
                  BorderStyle     =   1  'Fixed Single
                  Height          =   315
                  Left            =   12000
                  Picture         =   "FMaintenance.frx":47CC
                  Top             =   4380
                  Width           =   315
               End
               Begin VB.Image ImgPositionAtteinte 
                  BorderStyle     =   1  'Fixed Single
                  Height          =   315
                  Left            =   12300
                  Picture         =   "FMaintenance.frx":4B82
                  Top             =   4380
                  Width           =   315
               End
               Begin VB.Image ImgAxeReference 
                  BorderStyle     =   1  'Fixed Single
                  Height          =   315
                  Left            =   12600
                  Picture         =   "FMaintenance.frx":4F38
                  Top             =   4380
                  Width           =   315
               End
               Begin VB.Image ImgVariateurPret 
                  BorderStyle     =   1  'Fixed Single
                  Height          =   315
                  Left            =   12900
                  Picture         =   "FMaintenance.frx":52EE
                  Top             =   4380
                  Width           =   315
               End
               Begin VB.Image ImgMoteurTourne 
                  BorderStyle     =   1  'Fixed Single
                  Height          =   315
                  Left            =   13200
                  Picture         =   "FMaintenance.frx":56A4
                  Top             =   4380
                  Width           =   315
               End
               Begin VB.Shape SDecoration 
                  BackColor       =   &H00FFFFFF&
                  BackStyle       =   1  'Opaque
                  Height          =   2415
                  Index           =   1
                  Left            =   180
                  Top             =   420
                  Width           =   3075
               End
               Begin VB.Line Line2 
                  X1              =   10680
                  X2              =   10680
                  Y1              =   4620
                  Y2              =   4980
               End
               Begin VB.Line LineVE0 
                  X1              =   13350
                  X2              =   13350
                  Y1              =   4560
                  Y2              =   5100
               End
               Begin VB.Line LineVE1 
                  X1              =   13050
                  X2              =   13050
                  Y1              =   4560
                  Y2              =   5400
               End
               Begin VB.Line LineVE2 
                  X1              =   12750
                  X2              =   12750
                  Y1              =   4500
                  Y2              =   5640
               End
               Begin VB.Line LineVE3 
                  X1              =   12450
                  X2              =   12450
                  Y1              =   4560
                  Y2              =   6060
               End
               Begin VB.Line LineVE4 
                  X1              =   12150
                  X2              =   12150
                  Y1              =   4560
                  Y2              =   6360
               End
               Begin VB.Line LineVE5 
                  X1              =   11850
                  X2              =   11850
                  Y1              =   4560
                  Y2              =   6540
               End
               Begin VB.Line LineVE6 
                  X1              =   11550
                  X2              =   11550
                  Y1              =   4560
                  Y2              =   6900
               End
               Begin VB.Line LineVE7 
                  X1              =   11250
                  X2              =   11250
                  Y1              =   4560
                  Y2              =   7260
               End
               Begin VB.Line Line3 
                  X1              =   4590
                  X2              =   4590
                  Y1              =   4560
                  Y2              =   7080
               End
               Begin VB.Line LineV7 
                  X1              =   4890
                  X2              =   4890
                  Y1              =   4500
                  Y2              =   6780
               End
               Begin VB.Line LineV5 
                  X1              =   5190
                  X2              =   5190
                  Y1              =   4500
                  Y2              =   6480
               End
               Begin VB.Line LineV4 
                  X1              =   5490
                  X2              =   5490
                  Y1              =   4500
                  Y2              =   6180
               End
               Begin VB.Line LineV3 
                  X1              =   5790
                  X2              =   5790
                  Y1              =   4500
                  Y2              =   5910
               End
               Begin VB.Line LineV2 
                  X1              =   6090
                  X2              =   6090
                  Y1              =   4560
                  Y2              =   5580
               End
               Begin VB.Line LineV1 
                  X1              =   6390
                  X2              =   6390
                  Y1              =   4500
                  Y2              =   5310
               End
               Begin VB.Line LineV15 
                  X1              =   2070
                  X2              =   2070
                  Y1              =   4560
                  Y2              =   5010
               End
               Begin VB.Line LineV14 
                  X1              =   2370
                  X2              =   2370
                  Y1              =   4500
                  Y2              =   5310
               End
               Begin VB.Line LineV13 
                  X1              =   2670
                  X2              =   2670
                  Y1              =   4500
                  Y2              =   5610
               End
               Begin VB.Line LineV12 
                  X1              =   2970
                  X2              =   2970
                  Y1              =   4500
                  Y2              =   5910
               End
               Begin VB.Line LineV11 
                  X1              =   3270
                  X2              =   3270
                  Y1              =   4500
                  Y2              =   6210
               End
               Begin VB.Line LineV10 
                  X1              =   3570
                  X2              =   3570
                  Y1              =   4560
                  Y2              =   6480
               End
               Begin VB.Line LineV9 
                  X1              =   3870
                  X2              =   3870
                  Y1              =   4500
                  Y2              =   6810
               End
               Begin VB.Line Line1 
                  X1              =   4170
                  X2              =   4170
                  Y1              =   4500
                  Y2              =   7110
               End
               Begin VB.Line LineV0 
                  X1              =   6690
                  X2              =   6690
                  Y1              =   4500
                  Y2              =   5010
               End
               Begin VB.Shape SFocusVariateurs 
                  BorderColor     =   &H000000FF&
                  BorderWidth     =   4
                  Height          =   300
                  Left            =   120
                  Top             =   120
                  Visible         =   0   'False
                  Width           =   405
               End
               Begin VB.Label LCaracteristiquesVariateur 
                  Alignment       =   2  'Center
                  Appearance      =   0  'Flat
                  BackColor       =   &H00FF0000&
                  BorderStyle     =   1  'Fixed Single
                  Caption         =   "-"
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H0000FFFF&
                  Height          =   255
                  Left            =   180
                  TabIndex        =   95
                  Top             =   3060
                  Width           =   15075
               End
               Begin VB.Label LEP1 
                  Alignment       =   2  'Center
                  Appearance      =   0  'Flat
                  BackColor       =   &H0080C0FF&
                  BorderStyle     =   1  'Fixed Single
                  Caption         =   "Mot d'état"
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
                  Left            =   10320
                  TabIndex        =   70
                  Top             =   4020
                  Width           =   3195
               End
               Begin VB.Label LSP1 
                  Alignment       =   2  'Center
                  Appearance      =   0  'Flat
                  BackColor       =   &H0080C0FF&
                  BorderStyle     =   1  'Fixed Single
                  Caption         =   "Mot de commande"
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
                  Left            =   1920
                  TabIndex        =   69
                  Top             =   3960
                  Width           =   4935
               End
               Begin VB.Label Label25 
                  Alignment       =   2  'Center
                  Appearance      =   0  'Flat
                  BackColor       =   &H80000005&
                  BorderStyle     =   1  'Fixed Single
                  Caption         =   "Etat du variateur / Code défaut"
                  BeginProperty Font 
                     Name            =   "Arial"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   -1  'True
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H80000008&
                  Height          =   435
                  Left            =   9600
                  TabIndex        =   67
                  Top             =   4980
                  Width           =   1455
               End
               Begin VB.Label LEntreeProcess 
                  Alignment       =   2  'Center
                  Appearance      =   0  'Flat
                  BackColor       =   &H00FF0000&
                  BorderStyle     =   1  'Fixed Single
                  Caption         =   "DONNEES EMISES PAR LE VARIATEUR"
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H0000FFFF&
                  Height          =   255
                  Left            =   9360
                  TabIndex        =   66
                  Top             =   3540
                  Width           =   5655
               End
               Begin VB.Label LLibellePositionCible 
                  BackColor       =   &H0080C0FF&
                  BackStyle       =   0  'Transparent
                  Caption         =   "SP2/3 : Position cible"
                  BeginProperty Font 
                     Name            =   "Arial"
                     Size            =   9.75
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H8000000D&
                  Height          =   255
                  Left            =   1860
                  TabIndex        =   65
                  Top             =   7620
                  Width           =   2895
               End
               Begin VB.Label LLibelleConsigneVitesse 
                  BackColor       =   &H0080C0FF&
                  BackStyle       =   0  'Transparent
                  Caption         =   "SP4 : Consigne de vitesse"
                  BeginProperty Font 
                     Name            =   "Arial"
                     Size            =   9.75
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H8000000D&
                  Height          =   255
                  Left            =   1860
                  TabIndex        =   64
                  Top             =   8040
                  Width           =   2895
               End
               Begin VB.Label LPositionCible 
                  Alignment       =   2  'Center
                  BackColor       =   &H8000000E&
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
                  ForeColor       =   &H00FF0000&
                  Height          =   285
                  Left            =   4800
                  TabIndex        =   63
                  Top             =   7620
                  Width           =   1425
               End
               Begin VB.Label LConsigneVitesse 
                  Alignment       =   2  'Center
                  BackColor       =   &H8000000E&
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
                  ForeColor       =   &H00FF0000&
                  Height          =   285
                  Left            =   4800
                  TabIndex        =   62
                  Top             =   8040
                  Width           =   1425
               End
               Begin VB.Label LLibelleRampeAcceleration 
                  BackColor       =   &H0080C0FF&
                  BackStyle       =   0  'Transparent
                  Caption         =   "SP5 : Rampe d'accélération"
                  BeginProperty Font 
                     Name            =   "Arial"
                     Size            =   9.75
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H8000000D&
                  Height          =   315
                  Left            =   1860
                  TabIndex        =   61
                  Top             =   8460
                  Width           =   2895
               End
               Begin VB.Label LUniteRampeAcceleration 
                  Alignment       =   2  'Center
                  BorderStyle     =   1  'Fixed Single
                  Caption         =   "ms"
                  BeginProperty Font 
                     Name            =   "Arial"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   285
                  Left            =   6420
                  TabIndex        =   60
                  Top             =   8460
                  Width           =   975
               End
               Begin VB.Label LRampeAcceleration 
                  Alignment       =   2  'Center
                  BackColor       =   &H8000000E&
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
                  ForeColor       =   &H00FF0000&
                  Height          =   285
                  Left            =   4800
                  TabIndex        =   59
                  Top             =   8460
                  Width           =   1425
               End
               Begin VB.Label LLibelleRampeDeceleration 
                  BackColor       =   &H0080C0FF&
                  BackStyle       =   0  'Transparent
                  Caption         =   "SP6 : Rampe de décélération"
                  BeginProperty Font 
                     Name            =   "Arial"
                     Size            =   9.75
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H8000000D&
                  Height          =   315
                  Left            =   1860
                  TabIndex        =   58
                  Top             =   8880
                  Width           =   2895
               End
               Begin VB.Label LUnitePositionCible 
                  Alignment       =   2  'Center
                  BorderStyle     =   1  'Fixed Single
                  Caption         =   "mm"
                  BeginProperty Font 
                     Name            =   "Arial"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   285
                  Left            =   6420
                  TabIndex        =   57
                  Top             =   7620
                  Width           =   975
               End
               Begin VB.Label LUniteConsigneVitesse 
                  Alignment       =   2  'Center
                  BorderStyle     =   1  'Fixed Single
                  Caption         =   "mm / s"
                  BeginProperty Font 
                     Name            =   "Arial"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   285
                  Left            =   6420
                  TabIndex        =   56
                  Top             =   8040
                  Width           =   975
               End
               Begin VB.Label LRampeDeceleration 
                  Alignment       =   2  'Center
                  BackColor       =   &H8000000E&
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
                  ForeColor       =   &H00FF0000&
                  Height          =   285
                  Left            =   4800
                  TabIndex        =   55
                  Top             =   8880
                  Width           =   1425
               End
               Begin VB.Label LUniteRampeDeceleration 
                  Alignment       =   2  'Center
                  BorderStyle     =   1  'Fixed Single
                  Caption         =   "ms"
                  BeginProperty Font 
                     Name            =   "Arial"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   285
                  Left            =   6420
                  TabIndex        =   54
                  Top             =   8880
                  Width           =   975
               End
               Begin VB.Label LLibellePositionActuelle 
                  BackColor       =   &H0080C0FF&
                  BackStyle       =   0  'Transparent
                  Caption         =   "EP2/3 : Position actuelle"
                  BeginProperty Font 
                     Name            =   "Arial"
                     Size            =   9.75
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H8000000D&
                  Height          =   255
                  Left            =   9660
                  TabIndex        =   53
                  Top             =   7620
                  Width           =   2415
               End
               Begin VB.Label LUnitePositionActuelle 
                  Alignment       =   2  'Center
                  BorderStyle     =   1  'Fixed Single
                  Caption         =   "mm"
                  BeginProperty Font 
                     Name            =   "Arial"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   285
                  Left            =   13740
                  TabIndex        =   52
                  Top             =   7620
                  Width           =   975
               End
               Begin VB.Label LLibelleVitesseReelle 
                  BackColor       =   &H0080C0FF&
                  BackStyle       =   0  'Transparent
                  Caption         =   "EP4 : Vitesse réelle"
                  BeginProperty Font 
                     Name            =   "Arial"
                     Size            =   9.75
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H8000000D&
                  Height          =   315
                  Left            =   9660
                  TabIndex        =   51
                  Top             =   8040
                  Width           =   2295
               End
               Begin VB.Label LUniteVitesseReelle 
                  Alignment       =   2  'Center
                  BorderStyle     =   1  'Fixed Single
                  Caption         =   "mm / s"
                  BeginProperty Font 
                     Name            =   "Arial"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   285
                  Left            =   13740
                  TabIndex        =   50
                  Top             =   8040
                  Width           =   975
               End
               Begin VB.Label LLibelleCourantActif 
                  BackColor       =   &H0080C0FF&
                  BackStyle       =   0  'Transparent
                  Caption         =   "EP5 : Courant actif"
                  BeginProperty Font 
                     Name            =   "Arial"
                     Size            =   9.75
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H8000000D&
                  Height          =   315
                  Left            =   9660
                  TabIndex        =   49
                  Top             =   8460
                  Width           =   2415
               End
               Begin VB.Label LUniteCourantActif 
                  Alignment       =   2  'Center
                  BorderStyle     =   1  'Fixed Single
                  Caption         =   "%"
                  BeginProperty Font 
                     Name            =   "Arial"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   285
                  Left            =   13740
                  TabIndex        =   48
                  Top             =   8460
                  Width           =   975
               End
               Begin VB.Label LLibelleFacteurCharge 
                  BackColor       =   &H0080C0FF&
                  BackStyle       =   0  'Transparent
                  Caption         =   "EP6 : Facteur de charge"
                  BeginProperty Font 
                     Name            =   "Arial"
                     Size            =   9.75
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H8000000D&
                  Height          =   285
                  Left            =   9660
                  TabIndex        =   47
                  Top             =   8880
                  Width           =   2295
               End
               Begin VB.Label LUniteFacteurCharge 
                  Alignment       =   2  'Center
                  BorderStyle     =   1  'Fixed Single
                  Caption         =   "%"
                  BeginProperty Font 
                     Name            =   "Arial"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   285
                  Left            =   13740
                  TabIndex        =   46
                  Top             =   8880
                  Width           =   975
               End
               Begin VB.Label LFacteurCharge 
                  Alignment       =   2  'Center
                  BackColor       =   &H00FFFFFF&
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
                  ForeColor       =   &H00FF0000&
                  Height          =   285
                  Left            =   12120
                  TabIndex        =   45
                  Top             =   8880
                  Width           =   1425
               End
               Begin VB.Label LCourantActif 
                  Alignment       =   2  'Center
                  BackColor       =   &H00FFFFFF&
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
                  ForeColor       =   &H00FF0000&
                  Height          =   285
                  Left            =   12120
                  TabIndex        =   44
                  Top             =   8460
                  Width           =   1425
               End
               Begin VB.Label LVitesseReelle 
                  Alignment       =   2  'Center
                  BackColor       =   &H00FFFFFF&
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
                  ForeColor       =   &H00FF0000&
                  Height          =   285
                  Left            =   12120
                  TabIndex        =   43
                  Top             =   8040
                  Width           =   1425
               End
               Begin VB.Label LPositionActuelle 
                  Alignment       =   2  'Center
                  BackColor       =   &H00FFFFFF&
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
                  ForeColor       =   &H00FF0000&
                  Height          =   285
                  Left            =   12120
                  TabIndex        =   42
                  Top             =   7620
                  Width           =   1425
               End
               Begin VB.Label LSortiesProcess 
                  Alignment       =   2  'Center
                  Appearance      =   0  'Flat
                  BackColor       =   &H00FF0000&
                  BorderStyle     =   1  'Fixed Single
                  Caption         =   "DONNEES RECUES PAR LE VARIATEUR"
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H0000FFFF&
                  Height          =   255
                  Left            =   420
                  TabIndex        =   41
                  Top             =   3540
                  Width           =   8715
               End
               Begin VB.Label LLibelles 
                  Alignment       =   2  'Center
                  Appearance      =   0  'Flat
                  BackColor       =   &H00FF0000&
                  BorderStyle     =   1  'Fixed Single
                  Caption         =   "PONT 1"
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H0000FFFF&
                  Height          =   255
                  Index           =   0
                  Left            =   180
                  TabIndex        =   40
                  Top             =   180
                  Width           =   3075
               End
               Begin VB.Label LLibelles 
                  Alignment       =   2  'Center
                  Appearance      =   0  'Flat
                  BackColor       =   &H00FF0000&
                  BorderStyle     =   1  'Fixed Single
                  Caption         =   "PONT 2"
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H0000FFFF&
                  Height          =   255
                  Index           =   1
                  Left            =   3480
                  TabIndex        =   39
                  Top             =   180
                  Width           =   3075
               End
               Begin VB.Shape SDecoration 
                  BackColor       =   &H00FFFFFF&
                  BackStyle       =   1  'Opaque
                  Height          =   2415
                  Index           =   2
                  Left            =   3480
                  Top             =   420
                  Width           =   3075
               End
               Begin VB.Label LLibelles 
                  Alignment       =   2  'Center
                  Appearance      =   0  'Flat
                  BackColor       =   &H00FF0000&
                  BorderStyle     =   1  'Fixed Single
                  Caption         =   "DEGRAISSAGE"
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H0000FFFF&
                  Height          =   255
                  Index           =   2
                  Left            =   6780
                  TabIndex        =   38
                  Top             =   180
                  Width           =   3075
               End
               Begin VB.Shape SDecoration 
                  BackColor       =   &H00FFFFFF&
                  BackStyle       =   1  'Opaque
                  Height          =   2415
                  Index           =   4
                  Left            =   6780
                  Top             =   420
                  Width           =   3075
               End
               Begin VB.Label LLibelles 
                  Alignment       =   2  'Center
                  Appearance      =   0  'Flat
                  BackColor       =   &H00FF0000&
                  BorderStyle     =   1  'Fixed Single
                  Caption         =   "-"
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H0000FFFF&
                  Height          =   255
                  Index           =   3
                  Left            =   10080
                  TabIndex        =   37
                  Top             =   180
                  Width           =   5175
               End
               Begin VB.Label LLibelles 
                  Alignment       =   2  'Center
                  Appearance      =   0  'Flat
                  BackColor       =   &H00FF0000&
                  BorderStyle     =   1  'Fixed Single
                  Caption         =   "-"
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H0000FFFF&
                  Height          =   255
                  Index           =   4
                  Left            =   15480
                  TabIndex        =   36
                  Top             =   180
                  Width           =   5535
               End
               Begin VB.Shape SDecoration 
                  BackColor       =   &H00C0E0FF&
                  BackStyle       =   1  'Opaque
                  Height          =   5595
                  Index           =   3
                  Left            =   9360
                  Top             =   3780
                  Width           =   5655
               End
               Begin VB.Shape SDecoration 
                  BackColor       =   &H00FFFFFF&
                  BackStyle       =   1  'Opaque
                  Height          =   9195
                  Index           =   6
                  Left            =   15480
                  Top             =   420
                  Width           =   5535
               End
               Begin VB.Shape SDecoration 
                  BackColor       =   &H00FFFFFF&
                  BackStyle       =   1  'Opaque
                  Height          =   2415
                  Index           =   5
                  Left            =   10080
                  Top             =   420
                  Width           =   5175
               End
               Begin VB.Shape SDecoration 
                  BackColor       =   &H00C0E0FF&
                  BackStyle       =   1  'Opaque
                  Height          =   5595
                  Index           =   0
                  Left            =   420
                  Top             =   3780
                  Width           =   8715
               End
               Begin VB.Shape SDecoration 
                  BackColor       =   &H0080C0FF&
                  BackStyle       =   1  'Opaque
                  Height          =   6315
                  Index           =   7
                  Left            =   180
                  Top             =   3300
                  Width           =   15075
               End
            End
         End
      End
   End
   Begin VB.PictureBox PBRenseignementsFenetre 
      Align           =   1  'Align Top
      BackColor       =   &H00FF0000&
      Height          =   375
      Left            =   0
      Picture         =   "FMaintenance.frx":5A5A
      ScaleHeight     =   315
      ScaleWidth      =   23115
      TabIndex        =   1
      Top             =   0
      Width           =   23175
      Begin VB.Label LRenseignementsFenetre 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00008000&
         BackStyle       =   0  'Transparent
         Caption         =   "MAINTENANCE GEREE"
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
      Height          =   1095
      Left            =   0
      ScaleHeight     =   1035
      ScaleWidth      =   23115
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   13500
      Width           =   23175
      Begin VB.PictureBox PBOutilsDeplacementFenetre 
         BackColor       =   &H00E0E0E0&
         Height          =   1035
         Left            =   0
         ScaleHeight     =   975
         ScaleWidth      =   1155
         TabIndex        =   10
         Top             =   0
         Visible         =   0   'False
         Width           =   1215
         Begin VB.CommandButton CBAgrandirFenetre 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Agrandir"
            DownPicture     =   "FMaintenance.frx":2A39C
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
            Picture         =   "FMaintenance.frx":2A546
            Style           =   1  'Graphical
            TabIndex        =   13
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
            TabIndex        =   12
            Top             =   0
            Width           =   255
         End
         Begin VB.HScrollBar HSDeplacementFenetre 
            Height          =   255
            LargeChange     =   300
            Left            =   0
            SmallChange     =   100
            TabIndex        =   11
            Top             =   720
            Width           =   915
         End
      End
      Begin VB.CommandButton CBListeDefautsComAutomate 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Liste des défauts de communication avec l'automate"
         DownPicture     =   "FMaintenance.frx":2A6F0
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   9660
         MaskColor       =   &H00FF00FF&
         Picture         =   "FMaintenance.frx":2ADF2
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   120
         UseMaskColor    =   -1  'True
         Width           =   5955
      End
      Begin VB.CommandButton CBAnnuler 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Annuler"
         DownPicture     =   "FMaintenance.frx":2B4F4
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
         Left            =   17400
         MaskColor       =   &H00FF00FF&
         Picture         =   "FMaintenance.frx":2BBF6
         Style           =   1  'Graphical
         TabIndex        =   8
         ToolTipText     =   " Annuler les dernières modifications "
         Top             =   120
         UseMaskColor    =   -1  'True
         Width           =   1515
      End
      Begin VB.CommandButton CBValider 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Valider"
         DownPicture     =   "FMaintenance.frx":2C2F8
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
         Left            =   19140
         MaskColor       =   &H00FF00FF&
         Picture         =   "FMaintenance.frx":2C9FA
         Style           =   1  'Graphical
         TabIndex        =   7
         ToolTipText     =   " Valider l'enregistrement "
         Top             =   120
         UseMaskColor    =   -1  'True
         Width           =   1515
      End
      Begin VB.CommandButton CBActualiser 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Caption         =   "Actualise&r"
         DownPicture     =   "FMaintenance.frx":2D0FC
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
         Left            =   15720
         MaskColor       =   &H00FF00FF&
         Picture         =   "FMaintenance.frx":2D7FE
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   " Actualiser les données "
         Top             =   120
         UseMaskColor    =   -1  'True
         Width           =   1515
      End
      Begin VB.CommandButton CBQuitter 
         BackColor       =   &H00FFFFFF&
         Cancel          =   -1  'True
         Caption         =   "&Quitter"
         DownPicture     =   "FMaintenance.frx":2DF00
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
         Left            =   20880
         MaskColor       =   &H00FF00FF&
         Picture         =   "FMaintenance.frx":2E602
         Style           =   1  'Graphical
         TabIndex        =   5
         ToolTipText     =   " Quitter cette fenêtre "
         Top             =   120
         UseMaskColor    =   -1  'True
         Width           =   1515
      End
      Begin VB.Timer TimerMaintenance 
         Enabled         =   0   'False
         Interval        =   300
         Left            =   1680
         Top             =   120
      End
      Begin MSComctlLib.ImageList ILEntreesSortiesAutomate 
         Left            =   6180
         Top             =   180
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   11
         ImageHeight     =   11
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   2
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FMaintenance.frx":2ED04
               Key             =   "niveau 0"
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FMaintenance.frx":2EEE2
               Key             =   "niveau 1"
            EndProperty
         EndProperty
      End
      Begin MSComctlLib.ImageList ILVoyantVariateurs 
         Left            =   4740
         Top             =   180
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   17
         ImageHeight     =   18
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   2
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FMaintenance.frx":2F0C0
               Key             =   "VoyantVertAllume"
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FMaintenance.frx":2F486
               Key             =   "VoyantVertEtteint"
            EndProperty
         EndProperty
      End
      Begin MSComctlLib.ImageList ILOutilsGrilles 
         Left            =   5460
         Top             =   180
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   32
         ImageHeight     =   32
         MaskColor       =   16711935
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   2
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FMaintenance.frx":2F84C
               Key             =   "indicateur rouge"
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FMaintenance.frx":3049E
               Key             =   "indicateur vert"
            EndProperty
         EndProperty
      End
      Begin VB.Shape SFocus 
         BorderColor     =   &H000000FF&
         BorderWidth     =   5
         Height          =   315
         Left            =   2280
         Top             =   120
         Visible         =   0   'False
         Width           =   780
      End
   End
End
Attribute VB_Name = "FMaintenance"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Rôle                    : Fenêtre gérant la maintenance de la ligne
' Nom                    : FMaintenance.frm
' Date de création : 17/07/2001
' Détails                :
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

'--- déclarations obligatoires ---
Option Explicit

'--- options générales ---
Option Base 1
DefVar A-Z

'--- constantes privées ---
Private Const TITRE_FENETRE As String = "MAINTENANCE / DIALOGUES AVEC L'AUTOMATE"
Private Const TITRE_MESSAGES As String = TITRE_FENETRE

'--- variables privées ---
Private PremiereActivation As Boolean
Private InterdireEvenements As Boolean                                  'pour interdire certains évènements
Private AdrAPI_FicheRedresseur As String

'--- tableaux privés ---

'--- variables publiques ---
Public NumFenetre As Long                             'numéro de la fenêtre lorsqu'elle devient active
    
Private Sub CBActualiser_GotFocus()
    
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

Private Sub CBActualiser_LostFocus()
    On Error Resume Next
    SFocus.Visible = False
End Sub

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

Private Sub CBInformationsDefautVariateurs_Click()
    On Error Resume Next
    AppelFenetre F_INFORMATIONS_DEFAUTS_VARIATEURS
End Sub

Private Sub CBListeDefautsComAutomate_Click()
    On Error Resume Next
    AppelFenetre F_INFORMATIONS_DEFAUTS_COMMUNICATION_AUTOMATE
End Sub

Private Sub CBQuitter_Click()
    On Error Resume Next
    If CBValider.Enabled = True Then
        Select Case AppelFenetre(F_MESSAGE, _
                                                 TITRE_MESSAGES, _
                                                 MESSAGE_1, _
                                                 TYPES_MESSAGES.T_AVERTISSEMENT, _
                                                 TYPES_BOUTONS.T_OUI_NON_ANNULER, _
                                                 EMPLACEMENT_FOCUS.E_SUR_OUI)
            Case vbYes
                'CBValider_Click
                DechargeFenetre
            Case vbNo
                'CBAnnuler_Click
                DechargeFenetre
            Case Else
        End Select
    Else
        DechargeFenetre
    End If
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
        Me.Refresh
        'If TBCommencantPar.Visible = True Then TBCommencantPar.SetFocus
        PremiereActivation = True
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

Private Sub PBBoutons_Resize()
    
    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next
    
    '--- déclaration ---
    
    '--- calculs de l'emplacement des boutons ---
    CBQuitter.Left = PBBoutons.ScaleWidth - MARGES.M_BORD_DROIT - CBQuitter.Width
    CBValider.Left = CBQuitter.Left - MARGES.M_ENTRE_BOUTONS - CBValider.Width
    CBAnnuler.Left = CBValider.Left - MARGES.M_ENTRE_BOUTONS - CBAnnuler.Width
    CBActualiser.Left = CBAnnuler.Left - MARGES.M_ENTRE_BOUTONS - CBActualiser.Width
    CBListeDefautsComAutomate.Left = CBActualiser.Left - MARGES.M_ENTRE_BOUTONS - CBListeDefautsComAutomate.Width
    
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
                'PBDetailsCharge.Move .Left, .Top, .Width, .Height
            
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
    With TimerMaintenance
        .Enabled = False
        .Interval = 0
    End With

    '--- déchargement de la fenêtre ---
    Me.Visible = False
    Unload Me
    Set OccFMaintenance = Nothing

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

    '--- affectation ---
  
    '--- divers sur la fenêtre ---
    With Me
        .Caption = TITRE_FENETRE
        .WindowState = vbMaximized
    End With
    
    '--- couleurs des fonds ---
    PBDeplacementFenetre(ZONES_DEPLACEMENT_FENETRE.Z_FILLE).Picture = ImgFondVert1
    
    '--- fond de l'image des boutons ---
    PBBoutons.Picture = ImgFondDesBoutons
    
End Sub

